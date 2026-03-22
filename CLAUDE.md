# DocFill — Word Add-In

Microsoft Word task pane add-in. Pure static files, no build step. Hosted on Cloudflare Pages.

**Live URL:** https://docfill.pages.dev
**Hosting:** Cloudflare Pages — auto-deploys on push to `main` (repo: `gardead23/docfill`)

## Architecture

- `taskpane.html` — task pane shell, loads Office.js from CDN
- `taskpane.js` — all add-in logic (scan, render, fill, reset, create, localStorage)
- `taskpane.css` — styles
- `manifest.xml` — Office add-in descriptor pointing to Cloudflare Pages URL
- `commands.html` — required empty shell for Office command surface

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
- `currentFields` — ordered array of `{ key, label, type }` from last scan
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

### Two-phase fill (re-fill without clearing)
`fillDocument()` uses a two-phase approach to support updating already-filled fields:
1. **Phase 1 (restore):** For each key in `toFill` that has a `lastFilledValues` entry, search the doc for the old value and replace it back with `{{key}}`
2. **Phase 2 (fill):** Search for `{{key}}` and replace with the new value

This avoids "placeholder already consumed" errors when re-filling without a full document reset.

**Duplicate value guard:** Before running either phase, `fillDocument()` checks if any two filled fields share the same value. If so, fill is blocked entirely with an error naming the conflicting fields. Same values would cause Phase 1 to confuse which occurrence to restore, producing incorrect output.

### Selection monitoring in Create mode
- `DocumentSelectionChanged` event fires on every cursor move — always registered, but handler returns immediately if `activeTab !== 'create'`
- Handler debounces with 250ms timeout then calls `fetchCurrentSelection()`
- `fetchCurrentSelection()` guards against overlapping calls with `selectionFetchInProgress` flag
- Multi-paragraph selections (Word returns `\r` between paragraphs) are discarded — only single-line selections are usable as placeholder text
- **Selection loss problem:** Clicking task pane buttons loses the document selection. Solved by storing `lastSelectedText` from the debounced event; the button handler uses the stored value, not the live selection

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
- **This word only** — replaces `results.items[0]` only
- **All N occurrences** — replaces all items
- **Cancel** — clears pending state

`window.confirm()` is blocked in Office add-in webviews on Mac (silently returns `false`). All confirmation UX must use inline HTML rendered into a `#status` or `#create-status` div.

### `confirm()` is blocked in Office add-in webviews
`window.confirm()` silently returns `false` in the Office add-in webview on Mac (and likely Windows). **Never use native confirm/alert/prompt dialogs.** All confirmation UX must use inline HTML rendered into `#status` or `#create-status`.

### Field order preservation on rescan
After a fill, rescanning finds fewer placeholders (consumed ones are gone). To preserve order:
1. Merge `docKeys` (found in doc) with `Object.keys(lastFilledValues)` (already filled)
2. Reconstruct order using `orderedExisting` (filter `currentFields` against the merged set) + `brandNewKeys` (truly new keys appended at the end)

### Ctrl+Z state sync
When `scanDocument()` runs, for any key that appears as `{{placeholder}}` in the document, that key is deleted from `lastFilledValues`. This handles cases where the user undid a fill via Ctrl+Z, keeping DocFill's state consistent with the document.

### localStorage persistence
Field labels and types are saved keyed by the sorted+joined set of placeholder keys. Same template shape → same configs on next open. Prefix: `template-filler:` (kept as-is to avoid breaking existing data).

## Deployment

Hosted on Cloudflare Pages with GitHub integration. Every push to `main` deploys automatically to https://docfill.pages.dev — no manual deploy step needed.

## Office.js Gotchas

- `body.search()` is case-sensitive when `matchCase: true` — use consistently
- All Word API calls must be inside `Word.run(async context => { ... await context.sync() })`
- `body.getOoxml()` returns a proxy object; read `.value` after `context.sync()`
- `Word.InsertLocation.replace` is the correct enum for in-place text substitution
- `Range.select()` scrolls to and highlights a range in the document (WordApi 1.1)
- `window.confirm()` is silently blocked in Office add-in webviews — never use it
- Minimum API requirement: `WordApi 1.3` (set in manifest `<Requirements>`)

## Future: AI Phase

Planned addition to the same task pane — "Fill with AI" button:
- User pastes source text (email, brief, notes) into a textarea
- Calls Claude API (client-side or thin proxy) → returns `{ field_key: value }` JSON
- Populates the same form; user reviews and clicks "Fill Document" as usual
- No architectural changes needed to existing fill/reset logic
