# Import Feature -- Manual Test Checklist

Template: `test-large-template.docx` (116 fields)

## Setup
1. Start local server: `cd /Users/danny/AI/docfill && npx serve .`
2. Open Word with `test-large-template.docx`
3. Scan the document (Fill tab > Scan Document)
4. Click the Import button in the section bar

---

## A. Vertical CSV File Tests

### 01 - Basic Match (`01-basic-match.csv`)
- [ ] Upload the file
- [ ] Summary: "7 fields filled"
- [ ] Verify: client_name, company_name, contract_number, prepared_by, total_fee, project_manager, technical_lead all populated
- [ ] Header row ("Key,Value") is NOT treated as data
- [ ] Click Done to close panel

### 02 - Key Normalization (`02-key-normalization.csv`)
- [ ] Upload the file
- [ ] Summary: "7 fields filled" (same 7 fields)
- [ ] All match despite spaces, ALLCAPS, hyphens, {{braces}}, {{ spaced braces }}

### 03 - Date Parsing (`03-dates.csv`)
- [ ] Upload the file
- [ ] Summary: "12 fields filled"
- [ ] start_date: March 15, 2026 (US slash)
- [ ] end_date: September 30, 2026 (ISO)
- [ ] effective_date: March 15, 2026 (long month)
- [ ] execution_date: April 1, 2026 (abbreviated month)
- [ ] All milestone and signer dates populated correctly

### 04 - Quoted Values (`04-quoted-values.csv`)
- [ ] Upload the file
- [ ] client_name: Smith, Jones & Associates (comma preserved)
- [ ] total_fee: $125,000.00
- [ ] project_scope: 4 lines in textarea (newlines preserved)
- [ ] force_majeure_events: escaped quotes unescaped correctly

### 05 - Unmatched Keys (`05-unmatched-keys.csv`)
- [ ] Summary: "1 fields filled" (client_name only)
- [ ] Red bucket: "5 not recognized: favorite_color, zodiac_sign, blood_type, shoe_size, middle_name"
- [ ] Panel stays open

### 06 - Empty Values (`06-empty-values.csv`)
- [ ] Summary: "3 fields filled" (client_name, prepared_by, technical_lead)
- [ ] Orange bucket: "4 rows with empty values skipped"
- [ ] Empty fields NOT cleared if they had draft values

### 07 - Duplicate Keys (`07-duplicate-keys.csv`)
- [ ] Summary: "3 fields filled"
- [ ] client_name = "Third Value" (last wins)
- [ ] company_name = "Acme Corporation" (last wins)
- [ ] total_fee = "$150000" (last wins)
- [ ] Orange duplicate warnings showing original key forms

### 08 - Bad Dates (`08-bad-dates.csv`)
- [ ] Summary: single red box "None of this data matched your document fields. Please check your spelling or formatting and try again."
- [ ] No individual date error lines shown
- [ ] Date dropdowns remain empty

### 09 - No Header Row (`09-no-header.csv`)
- [ ] Summary: "3 fields filled"
- [ ] First row treated as data (not skipped)

### 10 - Empty File (`10-empty-file.csv`)
- [ ] Error: "We couldn't read this data..."

### 11 - Single Column (`11-single-column.csv`)
- [ ] Error: "We couldn't read this data..."

### 12 - Kitchen Sink (`12-kitchen-sink.csv`)
- [ ] High filled count (50+ fields)
- [ ] Text, date, and long text fields all populated
- [ ] Multi-line project_scope shows 4 lines
- [ ] Quoted commas preserved
- [ ] Click "Fill Document" -- values appear in Word document
- [ ] Click "Reset All Fields" -- all revert to {{placeholders}}

---

## B. Vertical Paste Tests

### P1 - Tab-separated
```
client_name	Acme Corporation
company_name	Smith & Wesson Legal Group
start_date	03/15/2026
project_manager	Sarah Chen
total_fee	$125000
```
- [ ] Summary: "5 fields filled"

### P2 - Comma-separated
```
client_name,Acme Corporation
company_name,Smith & Wesson Legal Group
contract_number,CTR-2026-0042
```
- [ ] Summary: "3 fields filled"

### P3 - With header row
```
Field	Value
client_name	Acme Corporation
prepared_by	Danny Garcia
technical_lead	Mike Rivera
```
- [ ] Summary: "3 fields filled" (header skipped)

### P4 - Header-like first key preserved
```
name	Danny Garcia
client_name	Acme Corporation
```
- [ ] "name" row NOT skipped (col2 is not a header word)
- [ ] name is unmatched, client_name is filled

### P5 - Empty paste
- [ ] Click "Import pasted data" with empty textarea
- [ ] Error: "Paste some data first..."

### P6 - All bad dates (zero-fill vertical)
```
start_date	not a date
end_date	yesterday
effective_date	Q4 2026
```
- [ ] Single red box: "None of this data matched... Please check your spelling or formatting and try again."
- [ ] No individual date error lines

---

## C. Horizontal CSV File Tests

### 13 - Basic Horizontal (`13-horizontal-basic.csv`)
- [ ] Row picker appears with 4 rows and 6 columns
- [ ] Headers: Effective Date, Company Name, Client Name, Contract Number, Total Fee, Project Manager
- [ ] First row pre-selected (highlighted blue)
- [ ] Click row 2 -- highlight moves, radio changes
- [ ] Click "Import Row" -- fields populated with row 2 values
- [ ] Summary shows filled count

### 14 - 30 Rows (`14-horizontal-30rows.csv`)
- [ ] Row picker appears with all 30 rows
- [ ] Table scrolls vertically (max-height 250px)
- [ ] Headers stay sticky at top while scrolling
- [ ] Radio column stays sticky on left while scrolling horizontally
- [ ] Select row 15, click Import Row -- correct values filled

---

## D. Horizontal Paste Tests

### HP1 - Tab-separated horizontal
```
Effective Date	Company Name	Client Name	Contract Number
03/15/2026	Acme Corporation	Danny Garcia	CTR-2026-0042
03/22/2026	Apple Inc	Elmo Rodriguez	CTR-2026-0043
03/23/2026	Google LLC	Edgar Martinez	CTR-2026-0044
```
- [ ] Row picker appears (3 data rows)
- [ ] Select row 1, Import Row -- effective_date=March 15 2026, company_name=Acme, client_name=Danny Garcia, contract_number=CTR-2026-0042

### HP2 - Horizontal zero-match (wrong column headers)
```
Favorite Color	Zodiac Sign	Blood Type
Blue	Capricorn	O+
Red	Aries	A+
```
- [ ] Row picker appears
- [ ] Select row 1, Import Row
- [ ] Single red box: "None of this data matched... If you pasted a simple list, try importing again and clicking 'Import as Field Name and Value pairs instead'."
- [ ] Message mentions escape hatch (horizontal-specific)

### HP3 - Escape hatch (on genuinely horizontal data)
```
Effective Date	Company Name	Client Name	Contract Number
03/15/2026	Acme Corporation	Danny Garcia	CTR-2026-0042
03/22/2026	Apple Inc	Elmo Rodriguez	CTR-2026-0043
```
- [ ] Row picker appears
- [ ] Click "Import as Field Name and Value pairs instead"
- [ ] Shows "None of this data matched..." (expected -- escape hatch forces vertical parsing on horizontal data, producing nonsensical key-value pairs)
- [ ] This confirms the escape hatch is NOT useful for genuinely horizontal data -- it's only for when the heuristic incorrectly classifies vertical data as horizontal

### HP4 - Escape hatch (correct use case)
```
client_name	Acme Corporation	Legal Division
company_name	Google LLC	Cloud Division
```
- [ ] Row picker appears (misdetected as horizontal since 3 cols, all letters, uniform width)
- [ ] Click "Import as Field Name and Value pairs instead"
- [ ] client_name filled with "Acme Corporation	Legal Division" (vertical parser rejoins extra columns)
- [ ] company_name filled with "Google LLC	Cloud Division"
- [ ] This is the correct use of the escape hatch -- data was vertical but misdetected

---

## E. Row Picker UI Tests

### Clickable rows
- [ ] Click anywhere on a row -- radio selects and row highlights blue
- [ ] Click a different row -- previous deselects, new one selects

### Keyboard navigation
- [ ] Tab to radio buttons -- focusable
- [ ] Note: Arrow keys may scroll the task pane instead of changing selection (Office WebView limitation). Use Tab + Space to select a different row, or click.

### Back button
- [ ] Click "Back" -- returns to file/paste input view
- [ ] Import again -- works normally (no stale state)

### Pagination (100+ rows)
- [ ] Import a CSV with 100+ rows -- pagination appears below table
- [ ] Shows "1-100 of N" label with Previous (disabled) and Next buttons
- [ ] Click Next -- shows rows 101-N, Previous now enabled
- [ ] Click Previous -- back to first page
- [ ] Select a row on page 2, click Import Row -- correct row imported

---

## F. UI/UX Tests

### Import button states
- [ ] Import button disabled during Rescan
- [ ] Import button disabled during Fill Document
- [ ] Import button re-enables after scan/fill completes

### Panel behavior
- [ ] Import button toggles panel open/closed
- [ ] X button closes panel (on input view)
- [ ] Done button closes panel (on summary view)
- [ ] "Import different data" resets to input view

### Overwrite behavior
- [ ] Type "MANUAL VALUE" in client_name
- [ ] Import CSV with client_name=Acme Corporation
- [ ] Verify "Acme Corporation" overwrites "MANUAL VALUE"
- [ ] Orange empty-field highlight cleared on imported fields

### Fill after import
- [ ] Import data, click "Fill Document"
- [ ] Values appear in the Word document
- [ ] Click "Reset All Fields" -- all revert to {{placeholders}}

### Placeholder text
- [ ] Paste textarea says "Paste your data here..."
- [ ] Hint says "Two columns or a full spreadsheet"

---

## G. Error Message Tests

### Empty file
- [ ] "We couldn't read this data. Try a different format or check your file."

### Wrong file type
- [ ] Upload a file with a non-.csv extension (e.g., .txt or .docx) -- "Please select a .csv file."

### File too large
- [ ] Upload >5MB file -- "File is too large (max 5MB)..."

### Zero-match vertical
- [ ] Import garbage vertical data -- "None of this data matched... Please check your spelling or formatting and try again."

### Zero-match horizontal
- [ ] Select a row with no matching headers -- "None of this data matched... If you pasted a simple list, try importing again and clicking 'Import as Field Name and Value pairs instead'."

### All-empty vertical
- [ ] Import CSV where every value is blank -- "All rows had empty values. No data to import."

### All-empty horizontal rows
- [ ] Note: All-empty rows are filtered before the picker, so this case cannot be reached via the row picker UI. No test needed.

---

## H. Edge Cases

### 2-column data always vertical
- [ ] Paste or upload a 2-column spreadsheet with headers -- no row picker, vertical import
- [ ] Even if both columns have letter headers like "Name, Email"

### Inconsistent column widths stay vertical
- [ ] Paste data with 3 cols on first line, 2 cols on second -- vertical import, no row picker

### Import then Ctrl+Z in document
- [ ] Import values, fill the document, then Ctrl+Z in Word
- [ ] Click Rescan -- DocFill picks up the reverted state

### Clickable field labels
- [ ] Click a field label in Fill tab -- jumps to that placeholder in the document
- [ ] Click again -- cycles through occurrences
- [ ] Label turns blue on hover

### Pencil edit
- [ ] Click pencil icon -- label becomes editable input
- [ ] Type new label, press Enter -- label saved
- [ ] Press Escape -- reverts to original label
- [ ] Tooltip says "Edit label"

### Reset button
- [ ] Reset button appears inside the input after filling
- [ ] Positioned top-right of the input/textarea
- [ ] Tooltip says "Clear value"
- [ ] Click -- clears that field only
