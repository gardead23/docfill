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
Four top-level state variables:
- `currentFields` — ordered array of `{ key, label, type, dateFormat? }` from last scan
- `currentStorageKey` — localStorage key for field config persistence (includes document fingerprint)
- `hasFilled` — whether any fills have been applied in this session
- `lastFilledValues` — `Record<string, string>` mapping field keys to their last-filled values

### State model — Create mode
- `activeTab` — `'fill' | 'create'`
- `lastSelectedText` — text stored from the last `DocumentSelectionChanged` event
- `lastSuggestedName` — tracks the last auto-suggested placeholder name so manual edits aren't overwritten
- `createdPlaceholders` — `{ name: string, count: number }[]` list of created placeholders this session
- `pendingCreateText / pendingCreateName / pendingCreateIndex` — held during the multi-occurrence confirmation flow
- `chipNavIndex` — `Record<string, number>` tracks which occurrence to navigate to next per placeholder
- `selectionDebounceTimer / selectionFetchInProgress` — guard against overlapping Word.run calls

### Content controls as first-class field objects

DocFill uses Word content controls as permanent, stable field anchors. `{{placeholder}}` text is only an authoring/import format -- on scan, each is converted into a DocFill CC.

**Tag convention:** Every DocFill CC has `tag = "docfill:{key}"`. The `docfill:` prefix distinguishes DocFill fields from other CCs. Helper functions: `isDocFillCC(cc)`, `ccTagToKey(tag)`, `keyToCCTag(key)`.

**CC properties:** `tag = "docfill:{key}"`, `title = "Title Case Label"`, `appearance = "boundingBox"`, `placeholderText = "{{key}}"`.

**Scan (Phase A):** Discovers existing DocFill CCs via tag prefix filter. Migrates old-style CCs (tag = raw key, no prefix) to `docfill:` prefix.
**Scan (Phase B):** Finds raw `{{key}}` text patterns across all bodies (body + headers/footers). Converts each to a DocFill CC by wrapping the range.
**Scan (Phase C):** Builds field list from all DocFill CCs. Hydrates `lastFilledValues` from CCs whose text is not their placeholder pattern.

**Fill:** All fields are CCs after scan. Fill simply finds CCs by tag (`contentControls.getByTag(keyToCCTag(key))`), replaces their text. Batched: one sync to load all collections, one sync to apply all updates.

**Reset:** Replaces CC text with `{{key}}`. CCs are **never deleted** -- they persist through fill/reset/reopen. This is the architectural invariant that makes the system stable.

**Create mode:** Inserts DocFill CCs directly (not raw `{{text}}`). The CC wraps the selected text with proper tag, title, appearance, and placeholderText.

**Chip navigation:** Uses `contentControls.getByTag(keyToCCTag(name))` instead of text search. Cycles through CCs by index.

Content controls persist across save/reopen (native Word OOXML elements: `<w:sdt>`). Tags are not unique -- multiple occurrences of the same placeholder share the same tag.

### Header/footer support
All scan, fill, and create operations process the full document: body + all section headers and footers (Primary, FirstPage, EvenPages). Helper functions `getAllBodies(context)` and `searchAllBodies(context, text, options)` enumerate all non-empty Body objects. Content controls in headers/footers are found by the same `contentControls.getByTag()` call (it's document-wide). Linked headers (Link to Previous) are handled by processing bodies sequentially for mutating operations -- placeholders already consumed by a linked copy are naturally skipped.

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
1. Finds DocFill CCs by tag (`contentControls.getByTag(keyToCCTag(name))`)
2. Syncs the chip's displayed count if the actual CC count has changed
3. Calls `ccs.items[targetIdx].select()` to scroll to and highlight that CC
4. Cycles through CCs on successive clicks (`chipNavIndex[name]`)

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
When `scanDocument()` runs, it rebuilds `lastFilledValues` entirely from DocFill CCs. A CC whose text matches `{{key}}` or is empty is "unfilled." A CC with any other text is "filled" and its value is hydrated into `lastFilledValues`. If a user undoes a fill via Ctrl+Z, the CC's text reverts to its previous state, and the next scan picks up the change. If a CC is deleted entirely (user manually removed it), the key disappears from the field list.

### localStorage persistence
Field labels, types, and per-field date formats are saved keyed by document fingerprint + sorted placeholder keys. The fingerprint is a djb2 hash of the first 200 non-placeholder characters of the document body, distinguishing templates that share the same placeholder names but have different surrounding text.

Key format: `template-filler:{fingerprint}:{sorted,keys}`. Legacy keys without a fingerprint are auto-migrated on first load via `loadFieldConfigsWithMigration()`.

Global date format stored separately under `docfill:dateFormat`.

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

- **Text boxes and shapes** are not accessible via the Word.js API. Placeholders inside text boxes, watermarks, or floating shapes cannot be scanned or filled.
- **Occurrence targeting fallback:** In Create mode, when the selected text has multiple occurrences, `fetchCurrentSelection()` attempts to identify the exact selected occurrence via `Range.compareLocationWith()`. If it succeeds, the button shows "This occurrence (#N)." If it fails (e.g., selection changed between detection and button click), it falls back to the first occurrence.

## Testing

Pure logic functions are extracted to `lib/pure.mjs` and tested with Vitest:

```bash
npm test        # run once
npm run test:watch  # watch mode
```

48 tests covering: `toTitleCase`, `escapeHtml`, `escapeAttr`, `guessFieldType`, `suggestPlaceholderName`, `daysInMonth`, `formatDate`, `buildStorageKey`, `isDocFillCC`, `ccTagToKey`, `keyToCCTag`, `placeholderText`, `isCCUnfilled`.

CI runs on every push/PR to `main` via GitHub Actions (`.github/workflows/ci.yml`): syntax check, manifest validation, and test suite.

## Office.js Gotchas

- `body.search()` is case-sensitive when `matchCase: true` — use consistently
- All Word API calls must be inside `Word.run(async context => { ... await context.sync() })`
- `Word.InsertLocation.replace` is the correct enum for in-place text substitution
- `Range.select()` scrolls to and highlights a range in the document (WordApi 1.1)
- `window.confirm()` is silently blocked in Office add-in webviews — never use it
- Minimum API requirement: `WordApi 1.3` (set in manifest `<Requirements>`)
- Headers/footers accessed via `section.getHeader(type)` / `section.getFooter(type)` — returns a `Body` with same API as main body
- Three header/footer types must be checked separately: `Primary`, `FirstPage`, `EvenPages`

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
