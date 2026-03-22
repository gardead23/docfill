# DocFill — Word Add-In

A Microsoft Word task pane add-in for creating and filling document templates. Build templates from existing docs, then fill them — all inside Word with no server required.

**Live add-in:** https://docfill.pages.dev
**GitHub:** https://github.com/gardead23/docfill

---

## How It Works

### Filling a template

1. Add `{{placeholder_name}}` markers anywhere in your Word document
2. Open the DocFill task pane (Home ribbon → **DocFill** button)
3. Go to the **Fill** tab and click **Scan Document** — the add-in detects all placeholders
4. Customize labels and field types (Text, Date, Number, Long text) if needed
5. Fill in the values
6. Click **Fill Document** — all placeholders are replaced instantly

Labels and field types are remembered the next time you open the same template.

### Creating a template from an existing document

1. Open the DocFill task pane and go to the **Create** tab
2. Select any text in your document — the task pane shows a live preview of your selection
3. Type a placeholder name (auto-suggested from the selected text) and click **Replace with Placeholder**
4. Repeat for every piece of text you want to turn into a field
5. Click **Done — Fill This Template** to switch to the Fill tab and fill it immediately

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
├── manifest.xml      ← Office add-in descriptor (points to Cloudflare Pages URL)
├── taskpane.html     ← task pane UI
├── taskpane.css      ← styles
├── taskpane.js       ← all add-in logic (scan, fill, reset, create, localStorage)
├── commands.html     ← required Office command surface shell
├── icon-16.png       ← ribbon icons
├── icon-32.png
├── icon-80.png
├── CLAUDE.md         ← architectural notes and dev conventions
└── README.md
```

---

## Development & Deployment

The add-in is hosted on Cloudflare Pages, connected to the `gardead23/docfill` GitHub repo. Deployments are automatic — every push to `main` deploys to https://docfill.pages.dev.

For local testing, serve the files over localhost:
```bash
npx serve .
# → http://localhost:3000
```
Then update `<SourceLocation>` in `manifest.xml` to `http://localhost:3000/taskpane.html` and re-sideload.

---

## Tips

- **Create tab:** Select text in your doc, give it a name, and DocFill replaces it with a `{{placeholder}}` — builds your template without typing brackets manually
- **Multiple occurrences:** When you replace text that appears more than once, you can replace just the first instance or all of them at once
- **Navigate placeholders:** In Create mode, click any chip in "Created so far" to highlight that placeholder in the document; click again to cycle through multiple occurrences
- **Re-fill:** Change a value and click Fill Document again — it updates the document without needing to clear first
- **Per-field undo:** Click the ↺ icon on any filled field to restore just that placeholder
- **Full reset:** Click **Clear all fields** → **Reset Document** to restore the original template (including any text you added around placeholders after scanning)
- **Rescan:** Click **↺ Rescan** to pick up any new placeholders added to the document
- **Labels are saved** per template shape — your customizations persist across sessions
- **Duplicate values blocked:** If two fields have the same value, Fill Document will stop and tell you which fields conflict

---

## Roadmap

- [ ] **AI-assisted filling** — paste an email or brief and Claude extracts field values automatically, populating the form for review before filling
