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
|  Placeholders: scrollable list      |
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
- `suppressSelectionPreview` -- boolean flag that prevents selection preview updates during programmatic chip navigation (so `cc.select()` does not trigger the selection UI)
- `selectionFetchGeneration` -- generation counter incremented on chip navigation; in-flight `fetchCurrentSelection()` calls check this to abort if a newer navigation has started
- `chipNavGeneration` -- generation counter for scroll-lock; each `navigateToChip()` call increments it so older locks stop immediately
- `scrollLockRaf / scrollLockTimers` -- handle for `requestAnimationFrame` loop and scheduled `setTimeout` callbacks that enforce scroll position during chip navigation
- `suppressionTimer` -- delayed re-enable of `suppressSelectionPreview` (500ms after navigation completes)
- `statusPreserveTimer` -- delayed cleanup of the layout-preserving status clear (releases the fixed-height invisible placeholder after 1600ms)

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
- Auto-scroll to center + auto-focus first empty field on Fill click when some fields are still empty (`scrollToFirstEmptyField()` uses `block: "center"` and a 400ms delayed `focus()`)
- If search filter is active and there are hidden empty fields, Fill clears the search input and resets `fillFilterText` before applying validation styling, so all empty fields are visible
- Empty fields get `.field-empty` class (orange highlight). Empty date dropdowns also receive orange validation styling via CSS.

### Reset behavior

- **Full reset (`confirmReset()`):** replaces CC text with `{{key}}` for all keys in `lastFilledValues`. Two-phase: batch load all CC collections (one sync), then batch replace text (one sync). Scrolls to top after reset.
- **Per-field reset (`resetField()`):** replaces CC text with `{{key}}` for a single key. Clears the form input and hides the reset button for that field.
- CCs are **never deleted** during reset. The CC wrapper persists; only its inner text changes.
- Inline confirmation dialog for full reset (not `window.confirm()`) -- shows what will happen and offers Cancel.

### Create mode state machine

Three states:
1. **Idle** -- no text selected. Selection preview shows a dashed dropzone instruction box ("Highlight text in the document to begin"). Name input and Convert button are disabled (grayed out). The preview has `selection-idle` class styling.
2. **Active** -- text selected (single paragraph only, multi-paragraph discarded). Selection preview switches to a blue `has-selection` style showing "SELECTED" label and the quoted text. Name input enabled with auto-suggested name and auto-focus. Convert button enabled.
3. **Confirmation** -- multiple occurrences, case variants, or existing CCs detected. Inline confirmation dialog whose language depends on whether the chosen name already has CCs in the document:
   - **Existing field (Link language):** Header asks to "Link to your existing {{name}} field?" Buttons: "Link this one" / "Link all N" / Cancel. "Link N exact" button only appears when `exactCount > 1`. "Use different name" button shown full-width below the main buttons. A note explains "You already have a {{name}} field. Converting will link to the same field."
   - **No existing field (Convert/Replace language):** Header asks "Replace with {{name}}?" Buttons: "This one only" / "All N matches" / Cancel. Single exact match with no variants shows just "Convert" + Cancel. "All N exact" button only appears when `exactCount > 1`.
   - Headers always specify match type: "Found 2 exact" or "Found 2 exact and 1 with different capitalization"
   - Variant-only matches (no exact) use `confirmReplace('all')` (case-insensitive search), not `'single'`

**Selection loss problem:** Clicking task pane buttons loses the document selection. Solved by storing `lastSelectedText` from the debounced `DocumentSelectionChanged` event; the button handler uses the stored value, not the live selection.

**Selection-based targeting:** When `confirmReplace('single')` runs, it uses `Range.compareLocationWith()` (WordApi 1.3) against the current selection to find which occurrence the user intended. If the selection has changed, it shows "Selection changed. Please select the text again."

**Word matching:** `matchWholeWord` is conditionally enabled for simple word/phrase selections (text matching `/^\w+(\s+\w+)*$/`). This prevents partial-word matches when the selected text is a plain word or phrase. Selections containing punctuation or symbols use substring matching.

**Occurrence counting:** Case-insensitive search across all bodies (`searchAllBodies()`) with `dedupeRanges()` to handle linked headers. Parent-CC check skips ranges inside any content control (not just DocFill CCs). Reports exact-case count vs. variant count separately.

**Key conflict handling:**
- If CCs for the chosen name already exist, the confirmation dialog switches to "Link" language (see Confirmation state above). "Use different name" button (full-width, secondary style) calls `promptRenamePlaceholder()`.
- `promptRenamePlaceholder()` clears the name input, keeps the pending text, re-enables the Convert button

**Placeholders list:**
- Scrollable, searchable list with count badges and x delete buttons
- `loadExistingPlaceholders()` counts both DocFill CCs (by tag) and raw `{{key}}` text (for not-yet-converted patterns)
- Search filters the list in real time via `filterCreatedList()`
- Click a row to navigate to that placeholder in the document (cycles through occurrences via `chipNavIndex`)
- Count badges auto-sync if the CC count changes (e.g., user manually deleted one)

**Delete (`deleteCreatedPlaceholder()`):** Inline confirmation dialog first. Converts CC back to plain text (strips `{{braces}}`, keeps the word), then removes the CC wrapper (`cc.delete(true)`). Two-phase sync (replace text, then delete wrapper). Removes from `createdPlaceholders`, `currentFields`, and `lastFilledValues`.

### Chip navigation and scroll-lock

Clicking a row in the "Placeholders" list calls `navigateToChip(name)`, which selects the corresponding CC in the document and cycles through occurrences. This involves several coordinated mechanisms to prevent the task pane UI from jumping:

**Scroll-lock:** `cc.select()` can cause the task pane to scroll (Office WebView side-effect). To prevent this:
1. `captureTaskPaneScroll()` snapshots scroll positions of `window`, `document.scrollingElement`, `document.body`, `main`, and the created-list scroll container
2. `startTaskPaneScrollLock(snapshot, generation, durationMs)` runs a `requestAnimationFrame` loop for 1.5s that continuously restores the snapshot, plus scheduled `setTimeout` callbacks at key intervals (0, 50, 150, 300, 600, 1000, 1500ms) as a belt-and-suspenders approach
3. The scroll-lock re-fires right before and after `cc.select()` + `context.sync()`
4. Generation tokens (`chipNavGeneration`) ensure only the latest navigation holds the lock

**Selection preview suppression:** `cc.select()` fires `DocumentSelectionChanged`, which would normally update the selection preview and show the CC's placeholder text as a "selection." This is suppressed via:
1. `suppressSelectionPreview = true` before navigation begins
2. `selectionFetchGeneration++` to cancel any in-flight `fetchCurrentSelection()` calls
3. `onSelectionChanged()` and `fetchCurrentSelection()` both early-return when `suppressSelectionPreview` is true
4. A 500ms `suppressionTimer` re-enables `suppressSelectionPreview` after navigation completes

**Layout-preserving status clear:** When navigating, any existing create-status content (e.g., a confirmation dialog) must be cleared without causing layout shift. `navigateToChip()` sets the status element's `height` to its current `offsetHeight`, sets `visibility: hidden`, and clears `innerHTML`. A 1600ms timer later releases the fixed height. If a newer status message appears before the timer fires, `prepareCreateStatus()` cancels the timer and resets the height/visibility.

**Floating toast:** `showChipToast(msg)` displays a floating snackbar at the bottom of the task pane (CSS class `chip-toast`, positioned with `position: fixed`). Shows the placeholder name and occurrence index (e.g., "{{name}} (2 of 3)"). Auto-dismisses with opacity fade after 2s. Does not affect layout.

### Match navigation in Create mode

When `createPlaceholder()` finds multiple matches (2+), the confirmation dialog includes prev/next arrows (`< 1 of 4 >`) so users can cycle through matches and preview each one in the document before deciding. Wrap-around cycling (3 of 3 -> next -> 1 of 3).

**State:** `pendingMatchCount`, `pendingMatchIndex`, `pendingMatchCase`, `pendingMatchWholeWord`, `matchNavInFlight`. All cleared in `cancelCreateAction()`, `confirmReplace()`, and `navigateToChip()`.

**`navigateMatch(delta)`:** Re-searches with `searchAllBodies` + `dedupeRanges`, filters to non-DocFill-CC ranges, selects the match at `pendingMatchIndex`. Uses `chipNavGeneration` to invalidate stale navigations. `matchNavInFlight` flag prevents overlapping navigations. All dialog buttons except Cancel disabled during async work.

**Initial match index:** Set to the user's selected occurrence via `compareLocationWith` (includes boundary relations: ContainsStart, ContainsEnd, InsideStart, InsideEnd) so "This one only" defaults to the match the user originally selected.

**`confirmReplace('single')` with match nav:** When `pendingMatchCount > 0`, uses `pendingMatchIndex` with case-insensitive `searchAllBodies` instead of selection-based case-sensitive targeting. This correctly handles capitalization variants.

**Parent-CC filtering:** Create mode filters only skip ranges inside DocFill CCs (tag starts with `docfill:`), not all CCs. This prevents false "already inside a placeholder" errors from Word's built-in content controls. Both `tag` and `isNullObject` are explicitly loaded to avoid stale proxy data.

### Fill tab clickable labels

Field labels in the Fill tab are clickable buttons that call `navigateToChip(key)` to jump to the corresponding placeholder in the document. Pencil icon next to label for editing the display name (tooltip: "Edit label"). Reset button positioned inside the value input area (top-right, tooltip: "Clear value").

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

- **Search bar** (`#fill-search`) above the field list filters fields by key or label (case-insensitive substring match via `filterFillFields()`)
- **Sort dropdown** with two options: Document order (default, `fillSortMode = "doc"`) and A-Z by label (`fillSortMode = "az"`). Dropdown menu opened/closed by `toggleSortMenu()` with outside-click dismissal.
- `applyFieldDisplayOrder()` reorders existing DOM nodes without rebuilding them (preserves typed draft values). For each key in the ordered list, it sets `row.style.display` based on filter match, then appends the row to the list (DOM reorder without destroy/recreate).
- Shows "No fields match your search" when filter returns zero results (dynamically created `.fill-no-results` div)
- Fill clears search filter if there are hidden empty fields (so validation highlighting is visible)
- State variables: `fillSortMode` (`"doc"` | `"az"`) and `fillFilterText` (lowercase query string)

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
Generated sizes: 16, 32, 64, 80, 128px. AppSource requires `<IconUrl>` = 32x32 and `<HighResolutionIconUrl>` = 64x64. Ribbon icons use 16, 32, and 80px.

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
- **HF scan deferred and non-blocking:** Form renders immediately from body CCs. HF scan runs in the background; user can start filling while it completes. Scan status banner shows an indeterminate progress bar at its bottom edge.
- **Incremental check on tab switch:** `checkForNewPlaceholders()` is body-only (no HF loading). Loads body text + all CCs in one sync. Compares current CC keys against `currentFields` to detect additions/deletions. Only re-renders if changes detected.
- **Fill is CC-only:** Two syncs total (load collections + apply updates) regardless of field count. No text search involved.
- **DOM reorder preserves drafts:** `applyFieldDisplayOrder()` moves existing DOM nodes via `fieldsList.appendChild(row)` without destroying/rebuilding them. Sort and filter both use this function.
- **Document order via pairwise scoring:** `sortKeysByDocumentOrder()` uses `Range.compareLocationWith()` (WordApi 1.3) with a pairwise scoring approach. For multi-occurrence keys, it first finds the earliest CC per key using within-key pairwise comparisons (one sync), then compares representative CCs across keys (one sync). Each comparison increments a score for the "before" element; keys are sorted by descending score.
- **Full Scan Document runs deferred HF scan every time; tab-switch incremental checks are body-only.**

## Visual / CSS Notes

- **Tab bar:** iOS-style segmented controls for Create/Fill tabs and for field type pills (Text, Date, Long text)
- **Create tab idle state:** Dashed dropzone instruction box with icon and muted text
- **System font stack:** `"Segoe UI", -apple-system, BlinkMacSystemFont, system-ui, sans-serif` -- no external font dependencies (Google Fonts removed for security hardening)
- **Reset All Fields button:** Muted red text with hover underline (`.btn-clear` class)
- **Undo button:** Darker icon color (`#9ca3af`) and thicker SVG stroke (`stroke-width="2"`) for visibility
- **Input fields:** Explicit `cursor: text` set on `.field-value-input` and `.field-value-textarea` for Office WebView compatibility (WebView does not always default to text cursor on input elements)
- **Scan banners:** Include an indeterminate progress bar at the bottom edge (`.scan-banner-progress` with a sliding `.scan-banner-progress-bar` animation)
- **Empty field validation:** `.field-empty` class adds orange left border to the field row and orange border to inputs/textareas/date-selects
- **Placeholders list:** Monospace font for placeholder names, scrollable container, count badges, x delete buttons
- **Chip toast:** Fixed-position floating snackbar at the bottom of the task pane, auto-fades after 2s

## Known Limitations

- **Text boxes and shapes** are not accessible via the Word.js API. Placeholders inside text boxes, watermarks, or floating shapes cannot be scanned or filled.
- **Multi-paragraph selection** is not supported in Create mode. Only single-line text selections can be converted to placeholders.
- **Selection targeting fallback:** In Create mode, when `confirmReplace('single')` cannot match the selection to a range (e.g., selection changed between click and execution), it shows an error asking the user to re-select.

### Import

Users can import field values from a CSV file or pasted text instead of typing each value manually. Import populates the form only -- user still reviews and clicks "Fill Document" to apply to the document.

**UI:** Import button (down-arrow icon) in the fill toolbar next to the sort button. Opens an inline panel between the toolbar and the field list with two sections: file upload (.csv) and paste textarea. Button is disabled during HF scan.

**Parsing:** Two paths -- `parseCSV()` (RFC 4180 state-machine, handles quoted fields) and `parsePastedText()` (auto-detects tab vs comma via `detectDelimiter()`, which is paste-only). Both return `{rows: [{key, value}], skippedEmpty: number}`. Header rows are skipped only when BOTH columns are header-like words (via `isHeaderRow(col1, col2)`). Rows with empty values are skipped and counted.

**Key matching:** `normalizeImportKey(rawKey)` strips `{{braces}}`, replaces spaces/hyphens with underscores, lowercases, and removes non-word characters. `matchImportKeys(rows, fields)` normalizes both sides and matches. Deduplicates with last-wins; tracks duplicates with original imported keys for the summary.

**Date parsing:** `parseDateValue(value)` supports ISO, US slash, day-first slash (when first number >12), month names (long and abbreviated), and dash formats. No `new Date()` fallback -- returns null if no explicit format matches. `setFieldValue()` dynamically adds out-of-range year options to the dropdown.

**Summary buckets:** After import, the panel shows results in distinct categories with color coding:
1. Filled (green) -- fields successfully populated
2. Not recognized (red) -- imported keys that didn't match any placeholder
3. Could not parse date (red) -- key matched a date field but value couldn't be parsed
4. Empty values skipped (orange) -- rows with blank values
5. Duplicate keys (orange) -- same key appeared multiple times, shows original imported keys

Panel auto-closes after 3s only if completely clean (no warnings/errors). "Import different data" button resets to input state.

**File guard:** Files >5MB rejected before FileReader runs.

**No XLSX:** Deferred to avoid third-party runtime dependency (SheetJS). CSV + paste covers the core workflow since users can export or copy from Excel.

**Horizontal (multi-column) import:** When imported data has 3+ columns and the first row looks like headers (all cells contain letters, consistent column width across data rows, at least 1 data row), the import auto-detects horizontal format and shows a **row picker** -- a scrollable table with radio buttons where the user selects which row to import. The selected row's values are paired with column headers and converted to vertical `{key, value}` pairs, then flow through the existing `matchImportKeys -> setFieldValue -> showImportSummary` pipeline. 2-column data is always treated as vertical (no ambiguity). An "Use as two-column format" escape hatch on the row picker bypasses horizontal detection for edge cases.

**Row picker rendering:** Uses DOM APIs (`createElement`, `textContent`, `title` property) instead of innerHTML for safe handling of untrusted spreadsheet data. Rows are clickable (full `<tr>` click target). Capped at 100 displayed rows.

**Pure functions:** `normalizeImportKey`, `isHeaderRow`, `detectDelimiter`, `parseCSVRaw`, `parsePastedRaw`, `parseCSV`, `parsePastedText`, `parseDateValue`, `matchImportKeys`, `isHorizontalHeaderRow`, `detectImportFormat`, `extractHorizontalData`, `horizontalRowToVertical` -- all in both `lib/pure.mjs` (tested) and `taskpane.js` (inline duplicate).

## Testing

Pure logic functions are extracted to `lib/pure.mjs` and tested with Vitest:

```bash
npm test           # run once
npm run test:watch # watch mode
```

172 tests covering: `toTitleCase`, `escapeHtml`, `escapeAttr`, `guessFieldType`, `suggestPlaceholderName`, `daysInMonth`, `formatDate`, `buildStorageKey`, `isDocFillCC`, `ccTagToKey`, `keyToCCTag`, `placeholderText`, `isPlaceholderText`, `isPlaceholderTextForKey`, `isCCUnfilled`, `normalizeImportKey`, `isHeaderRow`, `detectDelimiter`, `parseCSVRaw`, `parsePastedRaw`, `parseCSV`, `parsePastedText`, `parseDateValue`, `matchImportKeys`, `isHorizontalHeaderRow`, `detectImportFormat`, `extractHorizontalData`, `horizontalRowToVertical`.

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

**Status:** Submitted April 19, 2026. Awaiting Microsoft review (typically 3-5 business days).

### Submission Details
- **Partner Center account:** Created under gardead23@gmail.com
- **Offer type:** Office Add-in (Microsoft 365 and Copilot program)
- **Categories:** Productivity, Content management, Utilities + tools
- **EULA:** Microsoft Standard Contract
- **Privacy policy:** https://docfill.smplhq.com/privacy
- **Support page:** https://docfill.smplhq.com/support
- **Markets:** All 242 markets
- **Additional purchases:** None (free)
- **Manifest icon sizes:** IconUrl = 32x32, HighResolutionIconUrl = 64x64

### Assets Submitted
- 5 screenshots (Mac, showing fill, import, row picker, create, and reset workflows)
- Testing instructions for Microsoft reviewer (create template, scan, fill, import, create mode)
- Long description entered in Partner Center (not in manifest.xml -- schema rejects it there)
- Privacy policy URL and support URL entered in Partner Center

### Distribution Strategy
- **AppSource** for public discovery and free install
- **smplhq.com** for licensing/billing (future -- premium features like import)
- Direct sideload option available for enterprise users via manifest URL

## Future: AI Phase

Planned addition to the same task pane -- "Fill with AI" button:
- User pastes source text (email, brief, notes) into a textarea
- Calls Claude API (client-side or thin proxy) -> returns `{ field_key: value }` JSON
- Populates the same form; user reviews and clicks "Fill Document" as usual
- No architectural changes needed to existing fill/reset logic
