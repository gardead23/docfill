# DocFill — Word Add-In

A Microsoft Word task pane add-in that detects `{{placeholders}}` in your document templates and fills them with a clean sidebar form. No apps to switch to, no server required — everything happens inside Word.

**Live add-in:** https://boisterous-dolphin-c2960c.netlify.app
**GitHub:** https://github.com/gardead23/docfill

---

## How It Works

1. Add `{{placeholder_name}}` markers anywhere in your Word document
2. Open the DocFill task pane (Home ribbon → **DocFill** button)
3. Click **Scan Document** — the add-in detects all placeholders
4. Customize labels and field types (Text, Date, Number, Long text) if needed
5. Fill in the values
6. Click **Fill Document** — all placeholders are replaced instantly

Labels and field types are remembered the next time you open the same template.

---

## Placeholder Format

Use `{{snake_case_name}}` syntax. Only letters, numbers, and underscores.

```
{{client_name}}       → auto-labeled "Client Name"
{{start_date}}        → auto-labeled "Start Date"
{{total_fee}}         → auto-labeled "Total Fee"
{{project_scope}}     → auto-labeled "Project Scope"
```

---

## Sideloading the Add-In

The add-in is already hosted and live — you just need to sideload `manifest.xml` into Word once.

### Mac (Word for Mac)

1. Open Word
2. Go to **Insert → Add-ins → My Add-ins**
3. Click **"..."** → **Upload My Add-in**
4. Select `manifest.xml` from this repo
5. The **DocFill** button appears in your Home ribbon

### Windows (Word for Windows)

**Option A — Upload directly (Microsoft 365):**
1. Open Word → **Insert → Get Add-ins → My Add-ins → Upload My Add-in**
2. Select `manifest.xml`

**Option B — Shared folder catalog:**
1. Put `manifest.xml` in a shared network folder
2. Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
3. Add the folder path, check "Show in Menu", restart Word
4. **Insert → My Add-ins** → select DocFill

---

## Team Deployment

For a small team:
1. Each member downloads `manifest.xml` from this repo and sideloads it once
2. Everyone uses the same hosted add-in — no local files needed after that

For Microsoft 365 organizations, an admin can deploy centrally:
- **Microsoft 365 Admin Center → Settings → Integrated apps → Upload custom app**
- All users get the add-in automatically, no sideloading required

---

## Field Types

| Type | Input | Inserted as |
|---|---|---|
| Text | Single-line input | Value as-is |
| Date | Date picker | "March 20, 2026" |
| Number | Number input | Value as-is |
| Long text | Multi-line textarea | Value as-is |

---

## File Structure

```
docfill/
├── manifest.xml      ← Office add-in descriptor (points to Netlify URL)
├── taskpane.html     ← task pane UI
├── taskpane.css      ← styles
├── taskpane.js       ← all add-in logic (scan, fill, reset, localStorage)
├── commands.html     ← required Office command surface shell
├── icon-16.png       ← ribbon icons
├── icon-32.png
├── icon-80.png
├── CLAUDE.md         ← architectural notes and dev conventions
└── README.md
```

---

## Development & Deployment

The add-in is hosted on Netlify. To deploy changes:

```bash
cd docfill
zip -r deploy.zip taskpane.html taskpane.js taskpane.css manifest.xml commands.html icon-16.png icon-32.png icon-80.png
curl -X POST "https://api.netlify.com/api/v1/sites/036b0e4d-dd56-4f26-92b5-52d6ccfb2192/deploys" \
  -H "Authorization: Bearer <NETLIFY_TOKEN>" \
  -H "Content-Type: application/zip" \
  --data-binary @deploy.zip && rm deploy.zip
```

For local testing, serve the files over localhost:
```bash
npx serve .
# → http://localhost:3000
```
Then update `<SourceLocation>` in `manifest.xml` to `http://localhost:3000/taskpane.html` and re-sideload.

---

## Tips

- **Re-fill:** Change a value and click Fill Document again — it updates the document without needing to clear first
- **Per-field undo:** Click the ↺ icon on any filled field to restore just that placeholder
- **Full reset:** Click **Clear all fields** → **Reset Document** to restore the original template
- **Rescan:** Click **↺ Rescan** to pick up any new placeholders added to the document
- **Labels are saved** per template shape — your customizations persist across sessions

---

## Roadmap

- [ ] **AI-assisted filling** — paste an email or brief and Claude extracts field values automatically, populating the form for review before filling
