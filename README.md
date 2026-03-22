# Template Filler — Word Add-In

A simple Word task pane add-in that detects `{{placeholders}}` in your document and fills them with a clean form. No apps to switch to, no server required — everything happens inside Word.

---

## How It Works

1. Add `{{placeholder_name}}` markers anywhere in your Word document
2. Open the Template Filler task pane
3. Click **Scan Document** — the add-in detects all placeholders
4. Customize labels and field types (Text, Date, Number, Paragraph) if needed
5. Fill in the values
6. Click **Fill Document** — all placeholders are replaced instantly

Labels and field types are remembered the next time you open the same template.

---

## Placeholder Format

Use `{{snake_case_name}}` syntax. Only letters, numbers, and underscores.

Examples:
```
{{client_name}}       → auto-labeled "Client Name"
{{start_date}}        → auto-labeled "Start Date"
{{total_fee}}         → auto-labeled "Total Fee"
{{project_scope}}     → auto-labeled "Project Scope"
```

---

## Setup

### Step 1: Host the add-in files

The add-in files (`taskpane.html`, `taskpane.css`, `taskpane.js`) must be served over HTTPS (or localhost for testing).

**Option A — Local testing (localhost)**
```bash
cd template-filler
npx serve .
# → Serving at http://localhost:3000
```
The `manifest.xml` already points to `http://localhost:3000/taskpane.html` for local testing.

**Option B — GitHub Pages (recommended for team sharing)**
1. Push the `template-filler` folder to a GitHub repository
2. Go to Settings → Pages → Deploy from branch (`main`, `/root`)
3. Note your GitHub Pages URL (e.g. `https://yourusername.github.io/template-filler/`)
4. Edit `manifest.xml` and update the `<SourceLocation>` URL:
   ```xml
   <SourceLocation DefaultValue="https://yourusername.github.io/template-filler/taskpane.html"/>
   ```
5. Share the updated `manifest.xml` with your team

---

### Step 2: Sideload the add-in into Word

#### Mac (Word for Mac)

1. Open Word
2. Go to **Insert → Add-ins → My Add-ins**
3. Click **"..."** (three dots) → **Upload My Add-in**
4. Select `manifest.xml`
5. The "Template Filler" button appears in your ribbon under **Home**

#### Windows (Word for Windows)

**Option A — Upload directly (Microsoft 365 subscribers):**
1. Open Word
2. Go to **Insert → Get Add-ins → My Add-ins → Upload My Add-in**
3. Select `manifest.xml`

**Option B — Shared folder catalog:**
1. Put `manifest.xml` in a shared network folder (e.g. `\\server\addins\`)
2. In Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
3. Add the folder path, check "Show in Menu"
4. Restart Word
5. **Insert → My Add-ins** → find Template Filler in the catalog

---

### Step 3: Open the task pane

After sideloading, a **Template Filler** button will appear in the Home ribbon. Click it to open the task pane.

---

## Team Deployment

For a small team, the simplest setup is:
1. Host on GitHub Pages (free HTTPS)
2. Share `manifest.xml` with each team member — they sideload once
3. Everyone accesses the same hosted add-in files

For Microsoft 365 organizations, an admin can deploy the add-in centrally:
- Microsoft 365 Admin Center → Settings → Integrated apps → Upload custom app
- All users get the add-in automatically — no sideloading needed

---

## File Structure

```
template-filler/
├── manifest.xml      ← Office add-in descriptor (update SourceLocation URL)
├── taskpane.html     ← task pane UI
├── taskpane.css      ← styles
├── taskpane.js       ← all add-in logic
└── README.md
```

---

## Tips

- **Undo**: If you fill and want to start over, use **Ctrl+Z** (Mac: **Cmd+Z**) to undo all replacements
- **Rescan**: If you add new placeholders to the document, click **↺ Rescan** to detect them
- **Field types**:
  - **Text** — plain text input
  - **Date** — date picker, inserts as "March 20, 2026"
  - **Number** — number input
  - **Paragraph** — multi-line text area
- **Labels are saved** per template (based on its placeholder set) — your customizations are remembered

---

## Coming Soon

- AI-assisted filling: paste an email or brief and Claude extracts the field values automatically
