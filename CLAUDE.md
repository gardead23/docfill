# DocFill -- Word Add-In

Microsoft Word task pane add-in. Pure static files, no build step. Hosted on Cloudflare Pages.

**Live URL:** https://docfill.smplhq.com
**Hosting:** Cloudflare Pages -- auto-deploys on push to `main` (repo: `gardead23/docfill`)

## Architecture

- `taskpane.html` -- task pane shell, loads Office.js from CDN
- `taskpane.js` -- all add-in logic (scan, render, fill, reset, create, localStorage)
- `taskpane.css` -- styles
- `lib/pure.mjs` -- pure helper functions extracted for testability (duplicated in taskpane.js)
- `tests/pure.test.js` -- Vitest unit tests for pure helpers
- `manifest.xml` -- Office add-in descriptor pointing to Cloudflare Pages URL
- `commands.html` -- required empty shell for Office command surface
- `privacy.html` -- privacy policy (required for AppSource submission)
- `support.html` -- support/help page (required for AppSource submission)

## UI Structure

Two-tab layout. Default tab: **Fill** (preserves existing behaviour). **Create** tab for template authoring.

```
+-------------------------------------+
|  [ Create ]  [ Fill ]               |  <- tab bar (segmented control, iOS-style)
+-----------+--------------------------+
|  Fill tab:                          |
|  [Search bar] [Sort: Doc/A-Z]       |
|  Empty state -> Scan Document       |
|  Fields list -> Fill Document       |
|                                     |
|  Create tab:                        |
|  Scan status banner (if scanning)   |
|  Selection preview (live)           |
|  Placeholder name input             |
|  Convert to Placeholder btn         |
|  Created so far: scrollable list    |
|    with search, count badges, x     |
|  Done -- Fill This Template footer  |
+-------------------------------------+
```

## Key Architectural Decisions

### Content controls as first-class field objects

DocFill uses Word content controls (CCs) as permanent, stable field anchors. `{{placeholder}}` text is only an authoring/import format -- on scan, each raw `{{key}}` is converted into a DocFill CC. CCs are **never deleted** during fill or reset -- they persist permanently through the document lifecycle. This is the core architectural invariant.

**Tag convention:** Every DocFill CC has `tag = "docfill:{key}"` where `{key}` is always lowercase. The `docfill:` prefix distinguishes DocFill fields from other CCs in the document. Helper functions:
- `isDocFillCC(cc)` -- checks tag prefix
- `ccTagToKey(tag)` -- extracts lowercase key from tag
- `keyToCCTag(key)` -- builds tag from key

**CC properties on creation:**
- `tag = "docfill:{key}"` (lowercase)
- `title = "Title Case Label"` (via `toTitleCase()`)
- `appearance = Word.ContentControlAppearance.boundingBox`
- `placeholderText = "{{key}}"`

**Case insensitivity:** All keys are normalized to lowercase. The `ccTagToKey()` function always lowercases, and `keyToCCTag()` always lowercases the key before prefixing. `isPlaceholderTextForKey()` compares case-insensitively. Body search uses `matchCase: false` for raw `{{key}}` discovery.

**Persistence:** Content controls are native Word OOXML elements (`<w:sdt>`). They persist across save/reopen without any DocFill-side storage. Tags are not unique -- multiple occurrences of the same placeholder share the same tag.

### State model -- Fill mode

Six top-level state variables:
- `currentFields` -- ordered array of `{ key, label, type, dateFormat? }` from last scan
- `currentStorageKey` -- localStorage key for field config persistence (includes document fingerprint)
- `hasFilled` -- whether any fills have been applied in this session
- `hasScannedOnce` -- true after at least one full scan (body + HF) has completed; prevents redundant full scans
- `lastFilledValues` -- `Record<string, string>` mapping field keys to their last-filled values
- `hfScanInProgress` -- true while background header/footer scan is running

### State model -- Create mode

- `activeTab` -- `'fill' | 'create'`
- `lastSelectedText` -- text stored from the last `DocumentSelectionChanged` event
- `lastSuggestedName` -- tracks the last auto-suggested placeholder name so manual edits are not overwritten
- `createdPlaceholders` -- `{ name: string, count: number }[]` list of all placeholders in the document
- `pendingCreateText / pendingCreateName` -- held during the multi-occurrence confirmation flow
- `chipNavIndex` -- `Record<string, number>` tracks which occurrence to navigate to next per placeholder
- `selectionDebounceTimer / selectionFetchInProgress` -- guard against overlapping Word.run calls

### Scan phases

**Phase A -- Discover and migrate existing CCs:**
1. Loads all document content controls (`contentControls.load("items,tag,text")`)
2. Migrates old-style CCs (tag = raw key without `docfill:` prefix) via `migrateOldCC()`
3. Builds a `ccsByKey` map of existing DocFill CCs grouped by key
4. Records `hadExistingCCs` flag for downstream logic

**Phase B -- Discover raw `{{key}}` text in body and convert to CCs:**
1. Loads main body text cheaply
2. Regex-extracts all `{{key}}` patterns, canonicalizes to unique lowercase keys
3. For each key, searches body with `matchCase: false` to find all case variants
4. Batches parent-CC checks (`parentContentControlOrNullObject`) to skip ranges already inside a DocFill CC
5. Converts remaining raw ranges to CCs via `convertRangeToCC()`
6. Reloads all CCs if any conversions occurred

**Phase C -- Build field list and hydrate state:**
1. Rebuilds CC map from all current CCs (after Phase B conversions)
2. Hydrates `lastFilledValues` from CCs whose text is not their placeholder pattern (handles Ctrl+Z recovery and document reopen)
3. Builds `currentFields` array in document order (first occurrence of each key determines order)
4. Loads/migrates saved field configs from localStorage
5. Renders the form and scrolls to top

**Background HF scan (deferred, non-blocking):**
- Fires after Phase C completes; form renders immediately with body fields
- Shows a scan status banner: "You can enter values now. We're finalizing the document setup in the background."
- Fill button is disabled during HF scan with "Finishing setup..." label
- Loads all section headers/footers, checks for raw `{{key}}` patterns
- Only processes HF bodies that contain `{{}}` patterns (skip empty/irrelevant)
- Each HF body is processed in its own try/catch (one failure does not block others)
- Converts raw text to CCs with parent-CC check
- If new fields found: preserves user's draft values (`collectRawDrafts()`), rebuilds form from all CCs, restores drafts (`restoreRawDrafts()`)
- Sets `hasScannedOnce = true` on completion
- Re-enables Fill button and hides scan banner

**Incremental scan on Fill tab switch (`checkForNewPlaceholders()`):**
- Fires when switching back to Fill tab after `hasScannedOnce` is true
- Fast body-only check: loads body text + all CCs in one sync
- Compares current CC keys against `currentFields` to detect additions/deletions
- Searches body for raw `{{key}}` text not inside existing CCs and converts
- If changes detected: preserves drafts, rebuilds form, restores drafts
- No HF loading -- this is a quick consistency check

### Fill behavior

- CC-only updates: finds CCs by tag (`contentControls.getByTag(keyToCCTag(key))`), replaces their text
- Two syncs total regardless of field count: one to load all CC collections, one to apply all updates
- `\n` in textarea values is converted to `\v` (vertical tab = soft line break in Word) to preserve inline formatting within the same paragraph
- Auto-rescan if zero CCs found (covers Ctrl+Z recovery where user undid the scan itself)
- Auto-scroll to first empty field with auto-focus on Fill click when some fields are still empty
- If search filter is active and there are hidden empty fields, Fill clears the filter to reveal them
- Empty fields get `.field-empty` class (orange highlight)

### Reset behavior

- **Full reset (`confirmReset()`):** replaces CC text with `{{key}}` for all keys in `lastFilledValues`. Two-phase: batch load all CC collections (one sync), then batch replace text (one sync). Scrolls to top after reset.
- **Per-field reset (`resetField()`):** replaces CC text with `{{key}}` for a single key. Clears the form input and hides the reset button for that field.
- CCs are **never deleted** during reset. The CC wrapper persists; only its inner text changes.
- Inline confirmation dialog for full reset (not `window.confirm()`) -- shows what will happen and offers Cancel.

### Create mode state machine

Three states:
1. **Idle** -- no text selected. Selection preview shows "Highlight text in the document to begin." Name input and Convert button are disabled.
2. **Active** -- text selected (single paragraph only, multi-paragraph discarded). Selection preview shows the selected text. Name input enabled with auto-suggested name. Convert button enabled.
3. **Confirmation** -- multiple occurrences, case variants, or existing CCs detected. Inline confirmation dialog with options like "This one only", "All N matches", "Link to existing", "Use different name", "Cancel".

**Selection loss problem:** Clicking task pane buttons loses the document selection. Solved by storing `lastSelectedText` from the debounced `DocumentSelectionChanged` event; the button handler uses the stored value, not the live selection.

**Selection-based targeting:** When `confirmReplace('single')` runs, it uses `Range.compareLocationWith()` (WordApi 1.3) against the current selection to find which occurrence the user intended. If the selection has changed, it shows "Selection changed. Please select the text again."

**Occurrence counting:** Case-insensitive search across all bodies (`searchAllBodies()`) with `dedupeRanges()` to handle linked headers. Parent-CC check skips ranges inside any content control (not just DocFill CCs). Reports exact-case count vs. variant count separately.

**Key conflict handling:**
- If CCs for the chosen name already exist: "Link to existing" (adds to the same field) or "Use different name" (clears input, lets user rename)
- `promptRenamePlaceholder()` clears the name input, keeps the pending text, re-enables the Convert button

**Created So Far list:**
- Scrollable, searchable list with count badges and x delete buttons
- `loadExistingPlaceholders()` counts both DocFill CCs (by tag) and raw `{{key}}` text (for not-yet-converted patterns)
- Search filters the list in real time via `filterCreatedList()`
- Click a row to navigate to that placeholder in the document (cycles through occurrences via `chipNavIndex`)
- Count badges auto-sync if the CC count changes (e.g., user manually deleted one)

**Delete (`deleteCreatedPlaceholder()`):** Converts CC back to plain text (strips `{{braces}}`, keeps the word), then removes the CC wrapper (`cc.delete(true)`). Two-phase sync (replace text, then delete wrapper). Removes from `createdPlaceholders`, `currentFields`, and `lastFilledValues`.

### Header/footer support

All scan, fill, and create operations process the full document: body + all section headers and footers (Primary, FirstPage, EvenPages). Helper functions `getAllBodies(context)` and `searchAllBodies(context, text, options)` enumerate all non-empty Body objects. Content controls in headers/footers are found by the same `contentControls.getByTag()` call (it is document-wide). Linked headers (Link to Previous) are handled by processing bodies sequentially for mutating operations -- placeholders already consumed by a linked copy are naturally skipped.

### Selection monitoring in Create mode

- `DocumentSelectionChanged` event is always registered (in `Office.onReady`), but handler returns immediately if `activeTab !== 'create'`
- Handler debounces with 250ms timeout then calls `fetchCurrentSelection()`
- `fetchCurrentSelection()` guards against overlapping calls with `selectionFetchInProgress` flag
- Multi-paragraph selections (Word returns `\r` or `\n` between paragraphs) are discarded -- only single-line selections are usable as placeholder text
- Auto-suggests a placeholder name via `suggestPlaceholderName()` -- only overwrites if the current input matches the last suggestion (preserves manual edits)

### Range proxy lifetime

Word Range objects only live within their `Word.run` context -- they cannot be persisted across calls. The create flow stores selected text as a string and uses `body.search()` at replacement time to find and replace all matching ranges.

### Ctrl+Z state sync

When `scanDocument()` runs, it rebuilds `lastFilledValues` entirely from DocFill CCs. A CC whose text matches `{{key}}` or is empty is "unfilled." A CC with any other text is "filled" and its value is hydrated into `lastFilledValues` (with `\v` converted back to `\n`). If a user undoes a fill via Ctrl+Z, the CC's text reverts, and the next scan picks up the change. If a CC is deleted entirely (user manually removed it), the key disappears from the field list.

### `confirm()` is blocked in Office add-in webviews

`window.confirm()` silently returns `false` in the Office add-in webview on Mac (and likely Windows). **Never use native confirm/alert/prompt dialogs.** All confirmation UX uses inline HTML rendered into `#status` or `#create-status`.

### Field types and date formatting

Three field types: `text` (default), `date` (month/day/year dropdowns + format selector), `paragraph` (textarea). Auto-detected from placeholder key name via `guessFieldType()`. Type pills (segmented controls, iOS-style) are always visible on each field card.

Date input uses three `<select>` dropdowns (Month, Day, Year) plus a "Today" button inside a `.date-dropdowns` container. The container div gets `id="val-${field.key}"` so `collectValues()` can read the selected values. Year range is current year +/- 5 years (21 options), defaulting to the current year.

**Date validation:** Day options are dynamically constrained by the selected month and year via `updateDayOptions(key)`. Changing month or year triggers `onchange` which rebuilds the day `<option>` list using `daysInMonth(month, year)` (handles leap years). If the previously selected day exceeds the new max, it clamps to the last valid day. `collectValues()` also applies a safety clamp via `daysInMonth()` before formatting. Empty date dropdowns get orange validation styling.

Date format system:
- **Global default** stored in `localStorage` under `docfill:dateFormat` (default: `"long"`)
- **Per-field override** stored as `dateFormat` property in the field config object (in `localStorage` under the `template-filler:` key)
- `formatDate(isoDate, format)` handles 5 formats: `long`, `abbr`, `short-us`, `short-intl`, `iso`
- `collectValues()` resolves: per-field override -> global default -> `"long"`
- Global selector shown above the fields list when any date fields exist; per-field dropdown shown below each date field

Legacy migration: old `type: "number"` values are silently converted to `"text"` on load.

### Search and sort on Fill tab

- **Search bar** above the field list filters fields by key or label (case-insensitive substring match)
- **Sort dropdown** with two options: Document order (default) and A-Z (by label)
- `applyFieldDisplayOrder()` reorders existing DOM nodes without rebuilding them (preserves typed draft values)
- Shows "No fields match your search" when filter returns zero results
- Fill clears search filter if there are hidden empty fields (so validation highlighting is visible)

### localStorage persistence

Field labels, types, and per-field date formats are saved keyed by document fingerprint + sorted placeholder keys. The fingerprint is a djb2 hash of the first 300 non-placeholder, non-CC characters of the document body, distinguishing templates that share the same placeholder names but have different surrounding text.

Key format: `template-filler:{fingerprint}:{sorted,keys}`. Legacy keys without a fingerprint are auto-migrated on first load via `loadFieldConfigsWithMigration()`. Case-insensitive scan of existing localStorage keys handles old case-sensitive entries.

Global date format stored separately under `docfill:dateFormat`.

### Draft value preservation

When the form is re-rendered (e.g., after HF scan finds new fields, or after incremental check), user-typed values that have not been filled yet are preserved:
- `collectRawDrafts()` reads all current form inputs including date dropdown states
- `restoreRawDrafts(drafts)` writes them back after rerender
- Overrides CC-hydrated values since drafts represent the user's latest edits

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

Hosted on Cloudflare Pages with GitHub integration. Every push to `main` deploys automatically to https://docfill.smplhq.com -- no manual deploy step needed.

## Performance

- **Body scan batched:** All raw `{{key}}` searches issued in parallel (one `body.search()` per key), loaded in one sync. Parent-CC checks also batched into one sync.
- **HF scan deferred and non-blocking:** Form renders immediately from body CCs. HF scan runs in the background; user can start filling while it completes.
- **Incremental check on tab switch:** `checkForNewPlaceholders()` uses body text + CC comparison -- no HF loading. Only re-renders if changes detected.
- **Fill is CC-only:** Two syncs total (load collections + apply updates) regardless of field count. No text search involved.
- **DOM reorder preserves drafts:** Sort and filter move DOM nodes without destroying/rebuilding them.
- **Full Scan Document runs deferred HF scan every time; tab-switch incremental checks are body-only.**

## Known Limitations

- **Text boxes and shapes** are not accessible via the Word.js API. Placeholders inside text boxes, watermarks, or floating shapes cannot be scanned or filled.
- **Multi-paragraph selection** is not supported in Create mode. Only single-line text selections can be converted to placeholders.
- **Selection targeting fallback:** In Create mode, when `confirmReplace('single')` cannot match the selection to a range (e.g., selection changed between click and execution), it shows an error asking the user to re-select.

## Testing

Pure logic functions are extracted to `lib/pure.mjs` and tested with Vitest:

```bash
npm test           # run once
npm run test:watch # watch mode
```

58 tests covering: `toTitleCase`, `escapeHtml`, `escapeAttr`, `guessFieldType`, `suggestPlaceholderName`, `daysInMonth`, `formatDate`, `buildStorageKey`, `isDocFillCC`, `ccTagToKey`, `keyToCCTag`, `placeholderText`, `isPlaceholderText`, `isPlaceholderTextForKey`, `isCCUnfilled`.

CI runs on every push/PR to `main` via GitHub Actions (`.github/workflows/ci.yml`): syntax check, manifest validation, and test suite.

## Office.js Gotchas

- `body.search()` is case-sensitive when `matchCase: true` -- use consistently
- All Word API calls must be inside `Word.run(async context => { ... await context.sync() })`
- `Word.InsertLocation.replace` is the correct enum for in-place text substitution
- `Range.select()` scrolls to and highlights a range in the document (WordApi 1.1)
- `Range.compareLocationWith()` is WordApi 1.3 -- used for selection-based occurrence targeting
- `range.parentContentControlOrNullObject` returns the enclosing CC or a null object (avoids exceptions)
- `range.insertContentControl()` wraps a range in a new CC
- `window.confirm()` is silently blocked in Office add-in webviews -- never use it
- Minimum API requirement: `WordApi 1.3` (set in manifest `<Requirements>`)
- Headers/footers accessed via `section.getHeader(type)` / `section.getFooter(type)` -- returns a `Body` with same API as main body
- Three header/footer types must be checked separately: `Primary`, `FirstPage`, `EvenPages`
- `contentControls.getByTag()` is document-wide (includes body + headers/footers)
- `\v` (vertical tab) inserts a soft line break in Word, keeping text within the same paragraph

## Fonts

Uses system font stack: `"Segoe UI", -apple-system, BlinkMacSystemFont, system-ui, sans-serif`. No external font dependencies. Google Fonts was removed to eliminate a third-party dependency (security hardening for AppSource).

## AppSource Submission

**Status:** In progress. Preparing for Microsoft AppSource (Office Add-in Store) listing.

### Completed
- Privacy policy page (`privacy.html`) -- describes local-only data handling
- Support page (`support.html`) -- quick start guide, FAQ, contact email (support@smplhq.com)
- Manifest updated: `<SupportUrl>`, correct icon sizes (64/128). Passes `npx office-addin-manifest validate`.
- Icon sizes generated: 16, 32, 64, 80, 128px
- **Note:** `<LongDescription>` and `<PrivacyUrl>` are NOT valid in the XML manifest schema -- they must be entered in Partner Center during submission, not in manifest.xml.

### Remaining Before Submission
- [ ] Create Microsoft Partner Center account ($19 individual / $99 company)
- [ ] Complete identity verification and tax/payout profile
- [ ] Take screenshots (1366x768 recommended) showing scan, fill, and create workflows
- [ ] Write testing instructions for Microsoft reviewer
- [ ] Enter long description and privacy policy URL in Partner Center (not in manifest.xml -- schema rejects them there)
- [ ] Test add-in in Word for the web (reviewers often test there first)
- [ ] Submit via Partner Center > Marketplace offers > Office Add-ins

### Distribution Strategy
- **AppSource** for public discovery and free install
- **smplhq.com** for licensing/billing -- add-in checks license against SMPL HQ backend
- Direct sideload option available for enterprise users via manifest URL

## Future: AI Phase

Planned addition to the same task pane -- "Fill with AI" button:
- User pastes source text (email, brief, notes) into a textarea
- Calls Claude API (client-side or thin proxy) -> returns `{ field_key: value }` JSON
- Populates the same form; user reviews and clicks "Fill Document" as usual
- No architectural changes needed to existing fill/reset logic
