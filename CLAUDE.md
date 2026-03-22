# DocFill — Word Add-In

Microsoft Word task pane add-in. Pure static files, no build step. Hosted on Netlify.

**Live URL:** https://boisterous-dolphin-c2960c.netlify.app
**Netlify site ID:** `036b0e4d-dd56-4f26-92b5-52d6ccfb2192`

## Architecture

- `taskpane.html` — task pane shell, loads Office.js from CDN
- `taskpane.js` — all add-in logic (scan, render, fill, reset, localStorage)
- `taskpane.css` — styles
- `manifest.xml` — Office add-in descriptor pointing to Netlify URL
- `commands.html` — required empty shell for Office command surface

## Key Architectural Decisions

### State model
Five top-level state variables:
- `currentFields` — ordered array of `{ key, label, type }` from last scan
- `currentStorageKey` — localStorage key for field config persistence
- `originalOoxml` — OOXML snapshot of the document captured at first scan (pre-fill); used for full document restore
- `hasFilled` — whether any fills have been applied in this session
- `lastFilledValues` — `Record<string, string>` mapping field keys to their last-filled values

### OOXML snapshot (`originalOoxml`)
- Captured in `scanDocument()` via `body.getOoxml()`, but **only when `lastFilledValues` is empty** — i.e., frozen at the true original template state, never overwritten after fills happen
- Used by `confirmReset()` to restore the full document via `body.insertOoxml(..., Word.InsertLocation.replace)`
- Stays in JS memory only; never sent to any server; clears when task pane closes

### Two-phase fill (re-fill without clearing)
`fillDocument()` uses a two-phase approach to support updating already-filled fields:
1. **Phase 1 (restore):** For each key in `toFill` that has a `lastFilledValues` entry, search the doc for the old value and replace it back with `{{key}}`
2. **Phase 2 (fill):** Search for `{{key}}` and replace with the new value

This avoids "placeholder already consumed" errors when re-filling without a full document reset.

### `confirm()` is blocked in Office add-in webviews
`window.confirm()` silently returns `false` in the Office add-in webview on Mac (and likely Windows). **Never use native confirm/alert/prompt dialogs.** All confirmation UX must use inline HTML rendered into the `#status` div.

### Field order preservation on rescan
After a fill, rescanning finds fewer placeholders (consumed ones are gone). To preserve order:
1. Merge `docKeys` (found in doc) with `Object.keys(lastFilledValues)` (already filled)
2. Reconstruct order using `orderedExisting` (filter `currentFields` against the merged set) + `brandNewKeys` (truly new keys appended at the end)

### localStorage persistence
Field labels and types are saved keyed by the sorted+joined set of placeholder keys. Same template shape → same configs on next open. Prefix: `template-filler:`.

## Deployment

Deploy via Netlify REST API (zip upload):
```bash
cd template-filler
zip -r deploy.zip taskpane.html taskpane.js taskpane.css manifest.xml commands.html
curl -X POST "https://api.netlify.com/api/v1/sites/036b0e4d-dd56-4f26-92b5-52d6ccfb2192/deploys" \
  -H "Authorization: Bearer $NETLIFY_TOKEN" \
  -H "Content-Type: application/zip" \
  --data-binary @deploy.zip
```

## Office.js Gotchas

- `body.search()` is case-sensitive when `matchCase: true` — use consistently
- All Word API calls must be inside `Word.run(async context => { ... await context.sync() })`
- `body.getOoxml()` returns a proxy object; read `.value` after `context.sync()`
- `Word.InsertLocation.replace` is the correct enum for in-place text substitution
- Minimum API requirement: `WordApi 1.3` (set in manifest `<Requirements>`)

## Future: AI Phase

Planned addition to the same task pane — "Fill with AI" button:
- User pastes source text (email, brief, notes) into a textarea
- Calls Claude API (client-side or thin proxy) → returns `{ field_key: value }` JSON
- Populates the same form; user reviews and clicks "Fill Document" as usual
- No architectural changes needed to existing fill/reset logic
