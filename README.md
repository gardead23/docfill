# DocFill -- Word Add-In

A Microsoft Word task pane add-in for creating and filling document templates. Build templates from existing docs, then fill them -- all inside Word with no server required.

**Live add-in:** https://docfill.smplhq.com
**GitHub:** https://github.com/gardead23/docfill

---

## How It Works

### Filling a template

1. Add `{{placeholder_name}}` markers to your Word document (body, headers, and footers)
2. Open the DocFill task pane (Home ribbon -> **DocFill** button)
3. Go to the **Fill** tab and click **Scan Document** -- the add-in detects all placeholders and converts them into fillable fields
4. Customize labels and field types (Text, Date, Long text) if needed
5. Fill in the values
6. Click **Fill Document** -- all fields are updated instantly

Labels and field types are remembered the next time you open the same template. Fields persist in the document as content controls, so re-scanning picks up existing fields without losing data.

### Creating a template from an existing document

1. Open the DocFill task pane and go to the **Create** tab
2. Select any text in your document -- the task pane shows a live preview of your selection
3. Type a placeholder name (auto-suggested from the selected text) and click **Convert to Placeholder**
4. If the text appears multiple times, choose to replace just this one or all occurrences
5. The "Created so far" list shows all your fields -- click any to jump to it in the document
6. Click **Done -- Fill This Template** to switch to the Fill tab and fill it immediately

---

## Placeholder Format

Use `{{snake_case_name}}` syntax. Only letters, numbers, and underscores. Case does not matter -- `{{Client_Name}}` and `{{client_name}}` are treated as the same field.

```
{{client_name}}       -> auto-labeled "Client Name" (Text)
{{start_date}}        -> auto-labeled "Start Date" (Date)
{{total_fee}}         -> auto-labeled "Total Fee" (Text)
{{project_scope}}     -> auto-labeled "Project Scope" (Long text)
```

---

## Sideloading the Add-In

The add-in is already hosted and live -- you just need to sideload `manifest.xml` into Word once.

### Mac (Word for Mac)

1. Open Word
2. Go to **Insert -> Add-ins -> My Add-ins**
3. Click **"..."** -> **Upload My Add-in**
4. Select `manifest.xml` from this repo
5. The **DocFill** button appears in your Home ribbon

### Windows (Word for Windows)

**Option A -- Upload directly (Microsoft 365):**
1. Open Word -> **Insert -> Get Add-ins -> My Add-ins -> Upload My Add-in**
2. Select `manifest.xml`

**Option B -- Shared folder catalog:**
1. Put `manifest.xml` in a shared network folder
2. Word: **File -> Options -> Trust Center -> Trust Center Settings -> Trusted Add-in Catalogs**
3. Add the folder path, check "Show in Menu", restart Word
4. **Insert -> My Add-ins** -> select DocFill

---

## Team Deployment

For a small team:
1. Each member downloads `manifest.xml` from this repo and sideloads it once
2. Everyone uses the same hosted add-in -- no local files needed after that

For Microsoft 365 organizations, an admin can deploy centrally:
- **Microsoft 365 Admin Center -> Settings -> Integrated apps -> Upload custom app**
- All users get the add-in automatically, no sideloading required

---

## Field Types

| Type | Input | Inserted as |
|---|---|---|
| Text | Single-line input | Value as-is |
| Date | Month / Day / Year dropdowns + format selector | Formatted date (see below) |
| Long text | Multi-line textarea | Value with soft line breaks (preserves inline formatting) |

### Date Formats

Date fields support multiple output formats. Set a **global default** that applies to all date fields, or override the format on individual fields.

| Format | Example |
|---|---|
| Long (default) | March 22, 2026 |
| Abbreviated | Mar 22, 2026 |
| US short | 03/22/2026 |
| International | 22/03/2026 |
| ISO | 2026-03-22 |

---

## File Structure

```
docfill/
+-- manifest.xml           <- Office add-in descriptor (points to Cloudflare Pages URL)
+-- taskpane.html          <- task pane UI
+-- taskpane.css           <- styles
+-- taskpane.js            <- all add-in logic (scan, fill, reset, create, localStorage)
+-- lib/
|   +-- pure.mjs           <- pure helper functions (extracted for testability)
+-- tests/
|   +-- pure.test.js       <- Vitest unit tests (58 tests)
+-- commands.html          <- required Office command surface shell
+-- privacy.html           <- privacy policy (required for AppSource)
+-- support.html           <- support/help page (required for AppSource)
+-- icon-16.png            <- ribbon icons (16, 32, 64, 80, 128px)
+-- icon-32.png
+-- icon-64.png
+-- icon-80.png
+-- icon-128.png
+-- .github/workflows/
|   +-- ci.yml             <- GitHub Actions: syntax check, manifest validation, tests
+-- CLAUDE.md              <- architectural notes and dev conventions
+-- README.md
```

---

## Development & Deployment

The add-in is hosted on Cloudflare Pages, connected to the `gardead23/docfill` GitHub repo. Deployments are automatic -- every push to `main` deploys to https://docfill.smplhq.com.

For local testing, serve the files over localhost:
```bash
npx serve .
# -> http://localhost:3000
```
Then update `<SourceLocation>` in `manifest.xml` to `http://localhost:3000/taskpane.html` and re-sideload.

### Running Tests

```bash
npm test           # run once
npm run test:watch # watch mode
```

---

## Tips

- **Create tab:** Select text in your doc, give it a name, and DocFill converts it into a placeholder field -- builds your template without typing brackets manually
- **Multiple occurrences:** When you convert text that appears more than once, you can convert just this occurrence, all exact matches, or all matches including different capitalizations
- **Existing field linking:** If a placeholder name already exists, you can link to the existing field or choose a different name
- **Navigate placeholders:** In Create mode, click any item in "Created so far" to jump to that placeholder in the document; click again to cycle through multiple occurrences
- **Delete a placeholder:** Click the x button next to any created placeholder to convert it back to plain text
- **Re-fill:** Change a value and click Fill Document again -- it updates the document without needing to clear first
- **Per-field undo:** Click the undo icon on any filled field to restore just that placeholder
- **Full reset:** Click **Reset All Fields** to restore all filled values back to their `{{placeholder}}` markers. Other edits you made to the document are preserved
- **Search and sort:** Use the search bar to filter fields by name, and the sort dropdown to switch between document order and A-Z
- **Ctrl+Z recovery:** If you undo changes in Word, DocFill picks up the reverted state the next time it scans
- **Labels are saved** per template -- your customizations persist across sessions
- **Headers and footers supported:** DocFill scans and fills placeholders in the document body, headers, and footers. Text boxes and floating shapes are not supported

---

## Roadmap

- [ ] **AI-assisted filling** -- paste an email or brief and Claude extracts field values automatically, populating the form for review before filling
