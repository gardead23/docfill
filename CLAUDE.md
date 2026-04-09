# DocFill — Word Add-In

Microsoft Word task pane add-in. Pure static files, no build step. Hosted on Cloudflare Pages.

**Live URL:** https://docfill.smplhq.com
**Hosting:** Cloudflare Pages — auto-deploys on push to `main` (repo: `gardead23/docfill`)

## Architecture

- `taskpane.html` — task pane shell, loads Office.js from CDN
- `taskpane.js` — all add-in logic (scan, render, fill, reset, create, localStorage)
- `taskpane.css` — styles
- `manifest.xml` — Office add-in descriptor pointing to Cloudflare Pages URL
- `commands.html` — required empty shell for Office command surface
- `privacy.html` — privacy policy (required for AppSource submission)
- `support.html` — support/help page (required for AppSource submission)

## UI Structure

Two-tab layout. Default tab: **Fill** (preserves existing behaviour). **Create** tab for template authoring.

```
┌─────────────────────────────────┐
│  [ Create ]  [ Fill ]           │  ← tab bar
├─────────────────────────────────┤
│  Fill tab:                      │
│  Empty state → Scan Document    │
│  Fields list → Fill Document    │
│                                 │
│  Create tab:                    │
│  Selection preview (live)       │
│  Placeholder name input         │
│  Replace with Placeholder btn   │
│  Created so far: chips          │
│  Done — Fill This Template btn  │
└─────────────────────────────────┘
```

## Key Architectural Decisions

### State model — Fill mode
Five top-level state variables:
- `currentFields` — ordered array of `{ key, label, type, dateFormat? }` from last scan
- `currentStorageKey` — localStorage key for field config persistence
- `originalOoxml` — OOXML snapshot of the document; used for full document restore on reset
- `hasFilled` — whether any fills have been applied in this session
- `lastFilledValues` — `Record<string, string>` mapping field keys to their last-filled values

### State model — Create mode
- `activeTab` — `'fill' | 'create'`
- `lastSelectedText` — text stored from the last `DocumentSelectionChanged` event
- `lastSuggestedName` — tracks the last auto-suggested placeholder name so manual edits aren't overwritten
- `createdPlaceholders` — `{ name: string, count: number }[]` list of created placeholders this session
- `pendingCreateText / pendingCreateName` — held during the multi-occurrence confirmation flow
- `chipNavIndex` — `Record<string, number>` tracks which occurrence to navigate to next per placeholder
- `selectionDebounceTimer / selectionFetchInProgress` — guard against overlapping Word.run calls

### OOXML snapshot (`originalOoxml`)
- Captured in two places:
  1. `scanDocument()` via `body.getOoxml()`, **only when `lastFilledValues` is empty** — i.e., at the true template state before any fills
  2. **Right before the first fill** in `fillDocument()`, again only when `lastFilledValues` is empty — this captures any text the user added to the document *after* scanning, so reset restores to the complete pre-fill state
- Used by `confirmReset()` to restore the full document via `body.insertOoxml(..., Word.InsertLocation.replace)`
- Stays in JS memory only; never sent to any server; clears when task pane closes

### Content controls for fill tracking
When `fillDocument()` replaces a `{{placeholder}}` with a value, it wraps the inserted text in a hidden content control (`appearance: "hidden"`) tagged with the placeholder key (`cc.tag = key`). This provides stable document anchors for refill and reset operations.

**First fill:** Search for `{{key}}` text, replace with value, wrap in tagged content control.
**Refill:** Find content controls by tag (`contentControls.getByTag(key)`), replace their text directly. No text-search needed.
**Per-field reset:** Find content control by tag, replace text with `{{key}}`, unwrap control (`cc.delete(false)`).
**Full document reset:** Still uses the OOXML snapshot approach (replaces entire body).

Content controls persist across save/reopen (they are native Word OOXML elements: `<w:sdt>`). Tags are not unique -- multiple occurrences of the same placeholder share the same tag, and `getByTag()` returns all of them.

`lastFilledValues` is still maintained for UI state (form values, reset button visibility) but is no longer used for document-level text searching.

### Selection monitoring in Create mode
- `DocumentSelectionChanged` event fires on every cursor move — always registered, but handler returns immediately if `activeTab !== 'create'`
- Handler debounces with 250ms timeout then calls `fetchCurrentSelection()`
- `fetchCurrentSelection()` guards against overlapping calls with `selectionFetchInProgress` flag
- Multi-paragraph selections (Word returns `\r` between paragraphs) are discarded — only single-line selections are usable as placeholder text
- **Selection loss problem:** Clicking task pane buttons loses the document selection. Solved by storing `lastSelectedText` from the debounced event; the button handler uses the stored value, not the live selection
- **Occurrence targeting:** When the selected text appears multiple times, `fetchCurrentSelection()` uses `Range.compareLocationWith()` (WordApi 1.3) to determine which occurrence index the selection corresponds to, stored as `lastSelectedOccurrenceIndex`. This allows "This occurrence (#N)" to replace the correct instance instead of always targeting `items[0]`. Falls back to first occurrence if the index cannot be determined.

### Range proxy lifetime
Word Range objects only live within their `Word.run` context — they cannot be persisted across calls. The create flow stores selected text as a string and uses `body.search()` at replacement time to find and replace all matching ranges.

### Chip navigation
Clicking a `{{name}}` chip in "Created so far" calls `navigateToChip(name)`, which:
1. Searches the doc for `{{name}}` occurrences
2. Syncs the chip's displayed count if the actual occurrence count has changed (e.g., user added one manually)
3. Calls `results.items[targetIdx].select()` to scroll to and highlight that occurrence
4. Cycles through occurrences on successive clicks (`chipNavIndex[name]`)
Uses `Range.select()` (WordApi 1.1).

### Multi-occurrence replacement confirmation
When `createPlaceholder()` finds more than one occurrence, it stores `pendingCreateText` / `pendingCreateName` and renders an inline confirmation with three options:
- **First occurrence only** — replaces `results.items[0]` (the first match in document order)
- **All N occurrences** — replaces all items
- **Cancel** — clears pending state

`window.confirm()` is blocked in Office add-in webviews on Mac (silently returns `false`). All confirmation UX must use inline HTML rendered into a `#status` or `#create-status` div.

### `confirm()` is blocked in Office add-in webviews
`window.confirm()` silently returns `false` in the Office add-in webview on Mac (and likely Windows). **Never use native confirm/alert/prompt dialogs.** All confirmation UX must use inline HTML rendered into `#status` or `#create-status`.

### Field types and date formatting
Three field types: `text` (default), `date` (month/day/year dropdowns + format selector), `paragraph` (textarea). Auto-detected from placeholder key name via `guessFieldType()`. Type pills are always visible on each field card (no toggle/expand needed).

Date input uses three `<select>` dropdowns (Month, Day, Year) plus a "Today" button inside a `.date-dropdowns` container. The container div gets `id="val-${field.key}"` so `collectValues()` can read the selected values. Year range is current year +/- 5 years (21 options), defaulting to the current year. The "Today" button calls `setDateToday(key)` to populate all three dropdowns with the current date. Format selector is wrapped in a `.date-format-row` with a "Format:" label.

**Date validation:** Day options are dynamically constrained by the selected month and year via `updateDayOptions(key)`. Changing month or year triggers `onchange` which rebuilds the day `<option>` list using `daysInMonth(month, year)` (handles leap years). If the previously selected day exceeds the new max, it clamps to the last valid day. `collectValues()` also applies a safety clamp via `daysInMonth()` before formatting.

Date format system:
- **Global default** stored in `localStorage` under `docfill:dateFormat` (default: `"long"`)
- **Per-field override** stored as `dateFormat` property in the field config object (in `localStorage` under the `template-filler:` key)
- `formatDate(isoDate, format)` handles 5 formats: `long`, `abbr`, `short-us`, `short-intl`, `iso`
- `collectValues()` resolves: per-field override → global default → `"long"`
- Global selector shown above the fields list when any date fields exist; per-field dropdown shown below each date field

Legacy migration: old `type: "number"` values are silently converted to `"text"` on load.

### Field order preservation on rescan
After a fill, rescanning finds fewer placeholders (consumed ones are gone). To preserve order:
1. Merge `docKeys` (found in doc) with `Object.keys(lastFilledValues)` (already filled)
2. Reconstruct order using `orderedExisting` (filter `currentFields` against the merged set) + `brandNewKeys` (truly new keys appended at the end)

### Ctrl+Z state sync
When `scanDocument()` runs, it checks both raw `{{placeholder}}` text in the document body and the presence of content controls. For any key that appears as `{{placeholder}}` text, that key is deleted from `lastFilledValues` (the fill was undone). Additionally, any key in `lastFilledValues` that no longer has a corresponding content control is also removed. This handles cases where the user undid a fill via Ctrl+Z (which removes both the text change and the content control), keeping DocFill's state consistent with the document.

### localStorage persistence
Field labels, types, and per-field date formats are saved keyed by the sorted+joined set of placeholder keys. Same template shape → same configs on next open. Prefix: `template-filler:` (kept as-is to avoid breaking existing data). Global date format stored separately under `docfill:dateFormat`.

## Icons

Source file: `DocFill Icon.png` (1080x1080, RGBA, transparent background).
Generated sizes: 16, 32, 64, 80, 128px. The 64px and 128px sizes are required by AppSource (`<IconUrl>` and `<HighResolutionIconUrl>`).

To regenerate all icons from the source:

```bash
python3 << 'EOF'
from PIL import Image

src = "DocFill Icon.png"
img = Image.open(src).convert("RGBA")
img = img.crop(img.getbbox())  # trim transparent borders
w, h = img.size
size = max(w, h)
square = Image.new("RGBA", (size, size), (0, 0, 0, 0))
square.paste(img, ((size - w) // 2, (size - h) // 2), img)

for out_size, name in [(16, "icon-16.png"), (32, "icon-32.png"), (64, "icon-64.png"), (80, "icon-80.png"), (128, "icon-128.png")]:
    square.resize((out_size, out_size), Image.LANCZOS).save(name)
    print(f"Saved {name}")
EOF
```

Requires Pillow (`pip3 install Pillow`).

## Deployment

Hosted on Cloudflare Pages with GitHub integration. Every push to `main` deploys automatically to https://docfill.smplhq.com — no manual deploy step needed.

## Known Limitations

- **Document body only:** All scan/fill/create operations use `context.document.body`. Headers, footers, text boxes, and other story ranges are not scanned or filled. Documented in README and support page.
- **Restore is full-document rollback:** "Restore Original Document" replaces the entire document body with the pre-fill OOXML snapshot. Any edits made after filling (even unrelated prose) are lost. The confirmation dialog warns about this explicitly.
- **Occurrence targeting fallback:** In Create mode, when the selected text has multiple occurrences, `fetchCurrentSelection()` attempts to identify the exact selected occurrence via `Range.compareLocationWith()`. If it succeeds, the button shows "This occurrence (#N)." If it fails (e.g., selection changed between detection and button click), it falls back to the first occurrence.

## Office.js Gotchas

- `body.search()` is case-sensitive when `matchCase: true` — use consistently
- All Word API calls must be inside `Word.run(async context => { ... await context.sync() })`
- `body.getOoxml()` returns a proxy object; read `.value` after `context.sync()`
- `Word.InsertLocation.replace` is the correct enum for in-place text substitution
- `Range.select()` scrolls to and highlights a range in the document (WordApi 1.1)
- `window.confirm()` is silently blocked in Office add-in webviews — never use it
- Minimum API requirement: `WordApi 1.3` (set in manifest `<Requirements>`)

## Fonts

Uses system font stack: `"Segoe UI", -apple-system, BlinkMacSystemFont, system-ui, sans-serif`. No external font dependencies. Google Fonts was removed to eliminate a third-party dependency (security hardening for AppSource).

## AppSource Submission

**Status:** In progress. Preparing for Microsoft AppSource (Office Add-in Store) listing.

### Completed
- Privacy policy page (`privacy.html`) — describes local-only data handling
- Support page (`support.html`) — quick start guide, FAQ, contact email (support@smplhq.com)
- Manifest updated: `<SupportUrl>`, correct icon sizes (64/128). Passes `npx office-addin-manifest validate`.
- Icon sizes generated: 16, 32, 64, 80, 128px
- **Note:** `<LongDescription>` and `<PrivacyUrl>` are NOT valid in the XML manifest schema — they must be entered in Partner Center during submission, not in manifest.xml.

### Remaining Before Submission
- [ ] Create Microsoft Partner Center account ($19 individual / $99 company)
- [ ] Complete identity verification and tax/payout profile
- [ ] Take screenshots (1366x768 recommended) showing scan, fill, and create workflows
- [ ] Write testing instructions for Microsoft reviewer
- [ ] Enter long description and privacy policy URL in Partner Center (not in manifest.xml — schema rejects them there)
- [ ] Test add-in in Word for the web (reviewers often test there first)
- [ ] Submit via Partner Center > Marketplace offers > Office Add-ins

### Distribution Strategy
- **AppSource** for public discovery and free install
- **smplhq.com** for licensing/billing — add-in checks license against SMPL HQ backend
- Direct sideload option available for enterprise users via manifest URL

## Future: AI Phase

Planned addition to the same task pane — "Fill with AI" button:
- User pastes source text (email, brief, notes) into a textarea
- Calls Claude API (client-side or thin proxy) → returns `{ field_key: value }` JSON
- Populates the same form; user reviews and clicks "Fill Document" as usual
- No architectural changes needed to existing fill/reset logic
