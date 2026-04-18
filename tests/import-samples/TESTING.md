# Import Feature -- Manual Test Checklist

Template: `test-large-template.docx` (116 fields)

## Setup
1. Start local server: `cd /Users/danny/AI/docfill && npx serve .`
2. Open Word with `test-large-template.docx`
3. Scan the document (Fill tab > Scan Document)
4. Click the import button (down-arrow icon in toolbar)

---

## CSV File Tests

### 01 - Basic Match (`01-basic-match.csv`)
- [ ] Upload the file
- [ ] Summary: "7 fields filled"
- [ ] Verify: client_name=Acme Corporation, company_name=Smith & Wesson Legal Group, contract_number=CTR-2026-0042, prepared_by=Danny Garcia, total_fee=$125000, project_manager=Sarah Chen, technical_lead=Mike Rivera
- [ ] Header row ("Key,Value") is NOT treated as data
- [ ] Click Done to close panel

### 02 - Key Normalization (`02-key-normalization.csv`)
- [ ] Upload the file
- [ ] Summary: "7 fields filled" (same 7 fields as test 01)
- [ ] All match despite: "Client Name" (spaces), "COMPANY_NAME" (caps), "CONTRACT-NUMBER" (hyphen), "{{prepared_by}}" (braces), "{{ Total Fee }}" (spaced braces), "Project Manager" (spaces), "technical-lead" (hyphen)

### 03 - Date Parsing (`03-dates.csv`)
- [ ] Upload the file
- [ ] Summary: "12 fields filled"
- [ ] start_date dropdowns: March 15, 2026 (US slash)
- [ ] end_date dropdowns: September 30, 2026 (ISO)
- [ ] effective_date dropdowns: March 15, 2026 (long month name)
- [ ] execution_date dropdowns: April 1, 2026 (abbreviated month)
- [ ] deadline dropdowns: December 31, 2026 (US slash)
- [ ] review_date dropdowns: June 15, 2026 (US dash)
- [ ] All milestone and signer dates populated correctly

### 04 - Quoted Values (`04-quoted-values.csv`)
- [ ] Upload the file
- [ ] client_name shows: Smith, Jones & Associates (comma preserved)
- [ ] client_address shows full address with commas
- [ ] total_fee shows: $125,000.00
- [ ] project_scope textarea shows 4 lines (newlines preserved)
- [ ] payment_terms shows full text with commas
- [ ] force_majeure_events shows: ..."acts of God" (escaped quotes unescaped)

### 05 - Unmatched Keys (`05-unmatched-keys.csv`)
- [ ] Summary: "1 fields filled" (client_name only)
- [ ] Red text: "5 not recognized: favorite_color, zodiac_sign, blood_type, shoe_size, middle_name"
- [ ] Panel stays open (not auto-closed)

### 06 - Empty Values (`06-empty-values.csv`)
- [ ] Summary: "3 fields filled" (client_name, prepared_by, technical_lead)
- [ ] Orange text: "4 rows with empty values skipped"
- [ ] company_name, contract_number, total_fee, project_manager NOT cleared

### 07 - Duplicate Keys (`07-duplicate-keys.csv`)
- [ ] Summary: "3 fields filled" (client_name, company_name, total_fee)
- [ ] client_name = "Third Value" (last of 3 wins)
- [ ] company_name = "Acme Corporation" (last of 2 wins)
- [ ] total_fee = "$150000" (last of 3 wins)
- [ ] Orange duplicate warnings showing original key forms

### 08 - Bad Dates (`08-bad-dates.csv`)
- [ ] Red: "Could not parse date for start_date: "not a date""
- [ ] Red: "Could not parse date for end_date: "15th of March""
- [ ] Red: "Could not parse date for effective_date: "02/31/2026"" (Feb 31 invalid)
- [ ] Red: "Could not parse date for execution_date: "tomorrow""
- [ ] Red: "Could not parse date for deadline: "Q4 2026""
- [ ] Red: "Could not parse date for milestone_1_date: "02/29/2025"" (not a leap year)
- [ ] Date dropdowns remain empty/unchanged

### 09 - No Header Row (`09-no-header.csv`)
- [ ] Summary: "3 fields filled"
- [ ] First row "client_name,Acme Corporation" is treated as data (not skipped)

### 10 - Empty File (`10-empty-file.csv`)
- [ ] Error: "No data found to import..."

### 11 - Single Column (`11-single-column.csv`)
- [ ] Error: "No data found to import..." (rows with <2 columns skipped)

### 12 - Kitchen Sink (`12-kitchen-sink.csv`)
- [ ] Summary: high filled count (50+ fields)
- [ ] Text fields: client_name, company_name, contract_number, etc. all populated
- [ ] Date fields: start_date, end_date, milestone dates all have correct dropdowns
- [ ] Long text: project_scope shows 4 lines, project_description populated
- [ ] Quoted commas preserved in addresses, fees, payment_terms
- [ ] Click "Fill Document" after import -- values appear in the Word document
- [ ] Click "Reset All Fields" -- all revert to {{placeholders}}

---

## Paste Tests

Click import button, paste text into textarea, click "Import pasted data".

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
- [ ] "name" row NOT skipped (col2 "Danny Garcia" is not a header word)
- [ ] Both rows processed

### P5 - Empty paste
- [ ] Click "Import pasted data" with empty textarea
- [ ] Error: "Paste some data first..."

### P6 - Large paste with dates and long text
```
client_name	Acme Corporation
company_name	Smith & Wesson Legal Group
start_date	March 15, 2026
end_date	September 30, 2026
effective_date	2026-04-01
total_fee	$125,000
project_scope	Full platform redesign including mobile apps and web dashboard
payment_terms	Net 30 from invoice date
agreement_subject	Platform Redesign Services
governing_law	State of New York
jurisdiction	New York County
milestone_1_name	Discovery Complete
milestone_1_date	January 15, 2026
milestone_2_name	Design Approved
milestone_2_date	April 30, 2026
signer_1_name	Danny Garcia
signer_1_title	CEO
signer_2_name	John Smith
signer_2_title	VP Operations
```
- [ ] Summary: "19 fields filled"
- [ ] All dates parsed correctly
- [ ] All text values populated

---

## UI/UX Tests

### Panel behavior
- [ ] Import button toggles panel open/closed
- [ ] X button closes panel
- [ ] Done button closes panel (shown after successful import)
- [ ] "Import different data" resets to input state
- [ ] Panel stays open until dismissed

### Button states
- [ ] Import button disabled during Rescan
- [ ] Import button disabled during Fill Document
- [ ] Import button re-enables after scan/fill completes

### Overwrite behavior
- [ ] Type "MANUAL VALUE" in client_name
- [ ] Import CSV with client_name=Acme Corporation
- [ ] Verify "Acme Corporation" overwrites "MANUAL VALUE"
- [ ] Orange empty-field highlight cleared on imported fields

### Fill after import
- [ ] Import kitchen sink CSV
- [ ] Click "Fill Document"
- [ ] Scroll through Word document -- all imported values visible
- [ ] Click "Reset All Fields" -- all revert to {{placeholders}}
