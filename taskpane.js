/* global Office, Word */

"use strict";

// ── State ─────────────────────────────────────────────────────────────────────

/** @type {{ key: string, label: string, type: string, dateFormat?: string }[]} */
let currentFields = [];
let currentStorageKey = "";
let hasFilled = false;
/** @type {Record<string, string>} */
let lastFilledValues = {};

// ── Date Formats ─────────────────────────────────────────────────────────────

const DATE_FORMATS = [
  { value: "long",      label: "March 22, 2026" },
  { value: "abbr",      label: "Mar 22, 2026" },
  { value: "short-us",  label: "03/22/2026" },
  { value: "short-intl", label: "22/03/2026" },
  { value: "iso",       label: "2026-03-22" },
];

const DATE_FORMAT_LS_KEY = "docfill:dateFormat";

function getGlobalDateFormat() {
  try { return localStorage.getItem(DATE_FORMAT_LS_KEY) || "long"; }
  catch { return "long"; }
}

function setGlobalDateFormat(format) {
  try { localStorage.setItem(DATE_FORMAT_LS_KEY, format); } catch { /* ignore */ }
  // Update all per-field dropdowns to reflect new default label
  document.querySelectorAll(".date-format-select").forEach((sel) => {
    const defaultOpt = sel.querySelector('option[value=""]');
    if (defaultOpt) defaultOpt.textContent = `Default (${formatDatePreview(format)})`;
  });
}

function formatDatePreview(format) {
  return DATE_FORMATS.find((f) => f.value === format)?.label || "March 22, 2026";
}

// ── Create Mode State ──────────────────────────────────────────────────────────

let activeTab = "fill";
let lastSelectedText = "";
let lastSuggestedName = "";
/** @type {{ name: string, count: number }[]} */
let createdPlaceholders = [];
let pendingCreateText = "";
let pendingCreateName = "";
let pendingCreateIndex = -1;
/** @type {Record<string, number>} tracks which occurrence to navigate to next per placeholder */
const chipNavIndex = {};
let selectionDebounceTimer = null;
let selectionFetchInProgress = false;
/** Index of the selected occurrence among all matches (for "This occurrence" targeting). -1 = unknown. */
let lastSelectedOccurrenceIndex = -1;

// ── Document Range Helpers ────────────────────────────────────────────────────

/** Header/footer types to process. Lazily initialized after Office.js loads. */
let HF_TYPES = null;
function getHfTypes() {
  if (!HF_TYPES) {
    HF_TYPES = [Word.HeaderFooterType.primary, Word.HeaderFooterType.firstPage, Word.HeaderFooterType.evenPages];
  }
  return HF_TYPES;
}

/**
 * Collect all searchable Body objects in the document (body + all header/footer bodies).
 * Must be called inside Word.run. Returns an array of Body objects after sync.
 */
async function getAllBodies(context) {
  const bodies = [context.document.body];
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  const hfBodies = [];
  for (const section of sections.items) {
    for (const hfType of getHfTypes()) {
      hfBodies.push(section.getHeader(hfType));
      hfBodies.push(section.getFooter(hfType));
    }
  }
  // Load text to check which are non-empty
  for (const b of hfBodies) b.load("text");
  await context.sync();

  // Include all non-empty header/footer bodies (no dedup by text --
  // two unlinked regions can legitimately have identical text).
  for (const b of hfBodies) {
    if (b.text && b.text.trim()) bodies.push(b);
  }
  return bodies;
}

/**
 * Search for text across all document bodies (body + headers/footers).
 * Returns a flat array of Range items. Must be called inside Word.run.
 */
async function searchAllBodies(context, searchText, options) {
  const bodies = await getAllBodies(context);
  const allResults = [];
  for (const body of bodies) {
    const results = body.search(searchText, options);
    results.load("items");
    allResults.push(results);
  }
  await context.sync();
  return allResults.flatMap((r) => r.items);
}

/**
 * Deduplicate ranges that point to the same location (e.g., from linked headers).
 * Batches all Range.compareLocationWith() calls into a single sync for performance.
 * Must be called inside Word.run.
 */
async function dedupeRanges(context, ranges) {
  if (ranges.length <= 1) return ranges;

  // Queue all pairwise comparisons in one batch
  const comparisons = [];
  for (let i = 1; i < ranges.length; i++) {
    for (let j = 0; j < i; j++) {
      comparisons.push({ i, j, result: ranges[i].compareLocationWith(ranges[j]) });
    }
  }
  await context.sync(); // single sync for all comparisons

  // Mark duplicates
  const duplicates = new Set();
  for (const { i, result } of comparisons) {
    if (duplicates.has(i)) continue;
    const v = result.value;
    if (v === "Equal" || v === "Inside" || v === "Contains" ||
        v === Word.LocationRelation.equal ||
        v === Word.LocationRelation.inside ||
        v === Word.LocationRelation.contains) {
      duplicates.add(i);
    }
  }

  return ranges.filter((_, idx) => !duplicates.has(idx));
}

// ── Office Initialization ──────────────────────────────────────────────────────

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChanged
    );
  }
});

// ── Scan Document ──────────────────────────────────────────────────────────────

async function scanDocument() {
  showStatus("Scanning document...", "info");
  setScanButtonLoading(true);

  try {
    // Compute fingerprint before scanning so storage key is document-scoped
    await computeDocumentFingerprint();

    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      // Also load content controls to detect filled fields
      const allCCs = context.document.contentControls;
      allCCs.load("items,tag,text");

      // Load header/footer text for placeholder scanning
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      const hfBodies = [];
      for (const section of sections.items) {
        for (const hfType of getHfTypes()) {
          const h = section.getHeader(hfType);
          const f = section.getFooter(hfType);
          h.load("text");
          f.load("text");
          hfBodies.push(h, f);
        }
      }
      await context.sync();

      // Combine body + header/footer text for placeholder detection
      // (duplicate keys from linked headers are harmless -- deduped via Set below)
      let raw = body.text || "";
      for (const hf of hfBodies) {
        if (hf.text && hf.text.trim()) raw += " " + hf.text;
      }
      const matches = raw.match(/\{\{(\w+)\}\}/g) || [];
      const docKeys = [...new Set(matches)].map((m) => m.replace(/\{\{|\}\}/g, ""));

      // Build a map of currently filled fields from content controls
      const ccFilledKeys = new Set();
      const ccFilledValues = {};
      for (const cc of allCCs.items) {
        if (cc.tag) {
          ccFilledKeys.add(cc.tag);
          // Store the first value found for each tag (for hydration on reopen)
          if (!ccFilledValues[cc.tag]) ccFilledValues[cc.tag] = cc.text;
        }
      }

      // Hydrate lastFilledValues from content controls for keys not yet in memory
      // (handles reopen: content controls persist but in-memory state is empty)
      for (const key of ccFilledKeys) {
        if (!lastFilledValues[key] && !docKeys.includes(key)) {
          lastFilledValues[key] = ccFilledValues[key] || "";
        }
      }

      // Sync state with doc: if a key appears as a {{placeholder}} it wasn't filled
      // (or was undone via Ctrl+Z, which also removes the content control).
      for (const key of docKeys) {
        delete lastFilledValues[key];
      }
      // Also remove any keys that no longer have content controls
      for (const key of Object.keys(lastFilledValues)) {
        if (!ccFilledKeys.has(key)) {
          delete lastFilledValues[key];
        }
      }
      if (Object.keys(lastFilledValues).length === 0) hasFilled = false;
      if (ccFilledKeys.size > 0 && Object.keys(lastFilledValues).length > 0) hasFilled = true;

      // Merge: unfilled placeholders + filled fields (from memory and content controls)
      const allKeys = new Set([...docKeys, ...Object.keys(lastFilledValues)]);

      // Preserve original field order; append any brand-new keys at the end
      const orderedExisting = currentFields.map((f) => f.key).filter((k) => allKeys.has(k));
      const brandNewKeys = [...allKeys].filter((k) => !currentFields.some((f) => f.key === k));
      const keys = [...orderedExisting, ...brandNewKeys];

      if (keys.length === 0) {
        // Clear stale form state and show empty state with guidance
        currentFields = [];
        currentStorageKey = "";
        document.getElementById("fields-section").style.display = "none";
        document.getElementById("actions").style.display = "none";
        document.getElementById("empty-state").style.display = "block";
        document.querySelector(".empty-desc").innerHTML =
          'No <code>{{placeholders}}</code> found. Add fields like <code>{{client_name}}</code> to your document, then scan again.';
        setScanButtonLoading(false);
        return;
      }

      currentStorageKey = buildStorageKey(keys);
      const saved = loadFieldConfigsWithMigration(currentStorageKey, keys);

      currentFields = keys.map((key) => {
        const savedType = saved[key]?.type === "number" ? "text" : saved[key]?.type; // migrate old "number"
        return {
          key,
          label: saved[key]?.label || toTitleCase(key),
          type: savedType || guessFieldType(key),
          dateFormat: saved[key]?.dateFormat,
        };
      });

      saveFieldConfigs(currentStorageKey, currentFields);
      renderForm(currentFields);
      if (Object.keys(lastFilledValues).length > 0) hasFilled = true; // may still be true for partial fills
      hideStatus();
    });
  } catch (err) {
    showStatus("Error reading document: " + err.message, "error");
  }

  setScanButtonLoading(false);
}

// ── Render Form ────────────────────────────────────────────────────────────────

function renderForm(fields) {
  document.getElementById("empty-state").style.display = "none";
  document.getElementById("fields-section").style.display = "block";
  document.getElementById("actions").style.display = "flex";

  const n = fields.length;
  document.getElementById("field-count").textContent = n === 1 ? "1 field" : `${n} fields`;

  const fieldsList = document.getElementById("fields-list");
  fieldsList.innerHTML = "";

  fields.forEach((field) => {
    const row = document.createElement("div");
    row.className = "field-row";
    row.dataset.key = field.key;

    const fieldType = field.type === "number" ? "text" : field.type; // migrate old "number" type
    if (field.type === "number") field.type = "text";

    row.innerHTML = `
      <div class="field-top">
        <input
          class="field-label-input"
          type="text"
          value="${escapeAttr(field.label)}"
          placeholder="Label"
          onchange="onLabelChange('${escapeAttr(field.key)}', this.value)"
        />
        <button
          class="field-reset-btn"
          id="reset-btn-${escapeAttr(field.key)}"
          title="Reset this field"
          onclick="resetField('${escapeAttr(field.key)}')"
          style="display:none"
        >
          <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
            <path d="M2 5.5A4 4 0 1 1 3.5 9" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/>
            <path d="M2 3v2.5h2.5" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </button>
      </div>
      <div class="field-type-pills">
        ${[["text", "Text"], ["date", "Date"], ["paragraph", "Long text"]].map(([t, label]) => `
          <button
            class="type-pill ${fieldType === t ? "active" : ""}"
            data-type="${t}"
            onclick="setFieldType('${escapeAttr(field.key)}', '${t}')"
          >${label}</button>
        `).join("")}
      </div>
      ${buildValueInput(field)}
    `;

    fieldsList.appendChild(row);

    // Restore filled value and show reset button if this field has been filled
    if (lastFilledValues[field.key]) {
      if (field.type !== "date") {
        const input = row.querySelector(".field-value-input, .field-value-textarea");
        if (input) input.value = lastFilledValues[field.key];
      }
      const resetBtn = document.getElementById(`reset-btn-${field.key}`);
      if (resetBtn) resetBtn.style.display = "inline-flex";
    }
  });

  // Show global date format selector if any date fields exist
  renderGlobalDateFormat(fields);
}


function renderGlobalDateFormat(fields) {
  const container = document.getElementById("global-date-format");
  if (!container) return;
  const hasDateFields = fields.some((f) => f.type === "date");
  if (!hasDateFields) {
    container.style.display = "none";
    return;
  }
  const current = getGlobalDateFormat();
  container.style.display = "flex";
  container.innerHTML = `
    <label class="global-date-label" for="global-date-select">Default date format</label>
    <select id="global-date-select" class="date-format-select" onchange="onGlobalDateFormatChange(this.value)">
      ${DATE_FORMATS.map((f) => `<option value="${f.value}" ${f.value === current ? "selected" : ""}>${f.label}</option>`).join("")}
    </select>
  `;
}

function onGlobalDateFormatChange(format) {
  setGlobalDateFormat(format);
}

/** Build the right input element based on field type. */
function buildValueInput(field) {
  const id = `val-${field.key}`;
  if (field.type === "paragraph") {
    return `<textarea
      id="${id}"
      class="field-value-textarea"
      placeholder="Enter ${escapeHtml(field.label).toLowerCase()}..."
      rows="3"
    ></textarea>`;
  }
  if (field.type === "date") {
    const globalFmt = getGlobalDateFormat();
    const fieldFmt = field.dateFormat || "";
    const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    const monthOpts = months.map((m, i) => `<option value="${i + 1}">${m}</option>`).join("");
    const dayOpts = Array.from({length: 31}, (_, i) => `<option value="${i + 1}">${i + 1}</option>`).join("");
    const curYear = new Date().getFullYear();
    const yearOpts = Array.from({length: 21}, (_, i) => {
      const y = curYear - 5 + i;
      return `<option value="${y}">${y}</option>`;
    }).join("");
    return `<div class="date-dropdowns" id="${id}">
      <select class="date-select date-month" title="Month" onchange="updateDayOptions('${escapeAttr(field.key)}')"><option value="">Month</option>${monthOpts}</select>
      <select class="date-select date-day" title="Day"><option value="">Day</option>${dayOpts}</select>
      <select class="date-select date-year" title="Year" onchange="updateDayOptions('${escapeAttr(field.key)}')" ><option value="" selected>Year</option>${yearOpts}</select>
      <button type="button" class="date-today-btn" onclick="setDateToday('${escapeAttr(field.key)}')" title="Set to today">Today</button>
    </div>
    <div class="date-format-row">
      <span class="date-format-label">Format:</span>
      <select
        class="date-format-select"
        id="datefmt-${field.key}"
        onchange="setFieldDateFormat('${escapeAttr(field.key)}', this.value)"
        title="Date output format"
      >
        <option value="" ${!fieldFmt ? "selected" : ""}>Default (${formatDatePreview(globalFmt)})</option>
        ${DATE_FORMATS.map((f) => `<option value="${f.value}" ${fieldFmt === f.value ? "selected" : ""}>${f.label}</option>`).join("")}
      </select>
    </div>`;
  }
  return `<input
    id="${id}"
    class="field-value-input"
    type="text"
    placeholder="Enter ${escapeHtml(field.label).toLowerCase()}..."
  />`;
}

// ── Field Edit Handlers ────────────────────────────────────────────────────────

function setFieldType(key, newType) {
  const field = currentFields.find((f) => f.key === key);
  if (!field || field.type === newType) return;

  const oldValue = document.getElementById(`val-${key}`)?.value || "";
  field.type = newType;
  if (newType !== "date") delete field.dateFormat;
  saveFieldConfigs(currentStorageKey, currentFields);

  // Rebuild the value input
  const row = document.querySelector(`.field-row[data-key="${key}"]`);
  if (!row) return;
  // Remove old input + date dropdowns + date format row
  row.querySelectorAll(".field-value-input, .field-value-textarea, .date-dropdowns, .date-format-row").forEach((el) => el.remove());
  row.insertAdjacentHTML("beforeend", buildValueInput(field));
  if (newType !== "date") {
    const newInput = row.querySelector(".field-value-input, .field-value-textarea");
    if (newInput) newInput.value = oldValue;
  }

  // Update pill active states
  row.querySelectorAll(".type-pill").forEach((pill) => {
    pill.classList.toggle("active", pill.dataset.type === newType);
  });
}

function setDateToday(key) {
  const container = document.getElementById(`val-${key}`);
  if (!container) return;
  const now = new Date();
  container.querySelector(".date-month").value = now.getMonth() + 1;
  container.querySelector(".date-year").value = now.getFullYear();
  updateDayOptions(key);
  container.querySelector(".date-day").value = now.getDate();
}

/** Return how many days a given month/year has (handles leap years). */
function daysInMonth(month, year) {
  if (!month) return 31;
  if (!year) year = new Date().getFullYear();
  return new Date(year, month, 0).getDate();
}

/** Re-build day <option>s for a date field based on the selected month/year. */
function updateDayOptions(key) {
  const container = document.getElementById(`val-${key}`);
  if (!container) return;
  const monthSel = container.querySelector(".date-month");
  const daySel = container.querySelector(".date-day");
  const yearSel = container.querySelector(".date-year");
  const month = parseInt(monthSel.value, 10) || 0;
  const year = parseInt(yearSel.value, 10) || 0;
  const maxDay = daysInMonth(month, year);
  const currentDay = parseInt(daySel.value, 10) || 0;

  // Rebuild options
  let html = '<option value="">Day</option>';
  for (let d = 1; d <= maxDay; d++) {
    html += `<option value="${d}">${d}</option>`;
  }
  daySel.innerHTML = html;

  // Preserve selection, clamping if needed
  if (currentDay > 0) {
    daySel.value = currentDay > maxDay ? maxDay : currentDay;
  }
}

function setFieldDateFormat(key, format) {
  const field = currentFields.find((f) => f.key === key);
  if (!field) return;
  field.dateFormat = format || undefined;
  saveFieldConfigs(currentStorageKey, currentFields);
}

function onLabelChange(key, newLabel) {
  const field = currentFields.find((f) => f.key === key);
  if (field) field.label = newLabel;
  saveFieldConfigs(currentStorageKey, currentFields);
}

/** Auto-guess field type from key name on first scan. */
function guessFieldType(key) {
  const k = key.toLowerCase();
  if (/date|day|month|year|when|start|end|deadline|due|expir|signed|effective/.test(k)) return "date";
  if (/description|notes?|bio|summary|detail|scope|address|comments?|message|body|terms/.test(k)) return "paragraph";
  return "text";
}

// ── Fill Document ──────────────────────────────────────────────────────────────

async function fillDocument() {
  const btn = document.getElementById("fill-btn");
  const allValues = collectValues();

  // Separate filled vs empty
  const toFill = Object.fromEntries(Object.entries(allValues).filter(([, v]) => v.trim()));
  const emptyKeys = Object.keys(allValues).filter((k) => !allValues[k].trim());

  // Clear previous highlights
  document.querySelectorAll(".field-row.field-empty").forEach((r) => r.classList.remove("field-empty"));

  if (Object.keys(toFill).length === 0) {
    showStatus("Fill in at least one field to continue.", "error");
    return;
  }

  // Highlight skipped fields
  emptyKeys.forEach((key) => {
    document.querySelector(`.field-row[data-key="${key}"]`)?.classList.add("field-empty");
  });

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Filling...';
  hideStatus();

  try {
    let totalReplaced = 0;

    await Word.run(async (context) => {
      // Load all bodies once (not per-key)
      const bodies = await getAllBodies(context);

      for (const [key, value] of Object.entries(toFill)) {
        // Check if content controls already exist for this field (refill case)
        const existing = context.document.contentControls.getByTag(key);
        existing.load("items");
        await context.sync();

        if (existing.items.length > 0) {
          // Refill: update text inside existing content controls
          for (const cc of existing.items) {
            cc.insertText(value, Word.InsertLocation.replace);
          }
          totalReplaced += existing.items.length;
        } else {
          // First fill: search each body sequentially so linked headers
          // naturally dedupe (placeholder is gone by the second pass)
          for (const b of bodies) {
            try {
              const results = b.search(`{{${key}}}`, { matchCase: true });
              results.load("items");
              await context.sync();
              if (results.items.length === 0) continue;
              for (const range of results.items) {
                // Wrap in content control FIRST, then replace text inside it.
                // This ensures multi-paragraph text stays fully inside the CC.
                const cc = range.insertContentControl();
                cc.tag = key;
                cc.title = key;
                cc.appearance = Word.ContentControlAppearance.hidden;
                cc.insertText(value, Word.InsertLocation.replace);
              }
              totalReplaced += results.items.length;
              await context.sync();
            } catch (bodyErr) {
              // Linked headers throw GeneralException when their content was
              // already modified via the linked copy. Log other errors.
              if (bodyErr.code !== "GeneralException") {
                console.warn(`DocFill: skipped a region for {{${key}}}:`, bodyErr.message || bodyErr);
              }
            }
          }
        }
        await context.sync();
      }
    });

    if (totalReplaced === 0) {
      showStatus("No placeholders found — the document may already be filled.", "error");
    } else {
      hasFilled = true;
      Object.assign(lastFilledValues, toFill);
      for (const key of Object.keys(toFill)) {
        const resetBtn = document.getElementById(`reset-btn-${key}`);
        if (resetBtn) resetBtn.style.display = "inline-flex";
      }
      if (emptyKeys.length > 0) {
        const skipped = emptyKeys
          .map((k) => currentFields.find((f) => f.key === k)?.label || k)
          .join(", ");
        showStatus(`✓ Done. Highlighted fields were skipped: ${skipped}`, "info");
        // Scroll to the first skipped field so the user sees it
        const firstEmpty = document.querySelector(".field-row.field-empty");
        if (firstEmpty) firstEmpty.scrollIntoView({ behavior: "smooth", block: "nearest" });
      } else {
        showStatus("✓ All fields filled successfully.", "success");
      }
    }
  } catch (err) {
    showStatus("Error: " + err.message, "error");
  }

  btn.disabled = false;
  btn.innerHTML = "Fill Document";
}

function collectValues() {
  const values = {};
  const globalFmt = getGlobalDateFormat();
  currentFields.forEach((field) => {
    if (field.type === "date") {
      const container = document.getElementById(`val-${field.key}`);
      if (container) {
        const m = container.querySelector(".date-month")?.value;
        const d = container.querySelector(".date-day")?.value;
        const y = container.querySelector(".date-year")?.value;
        if (m && d && y) {
          const mi = parseInt(m, 10);
          const di = parseInt(d, 10);
          const yi = parseInt(y, 10);
          const maxDay = daysInMonth(mi, yi);
          const pad = (n) => String(n).padStart(2, "0");
          const safeDay = di > maxDay ? maxDay : di;
          const isoDate = `${yi}-${pad(mi)}-${pad(safeDay)}`;
          const fmt = field.dateFormat || globalFmt;
          values[field.key] = formatDate(isoDate, fmt);
        } else {
          values[field.key] = "";
        }
      } else {
        values[field.key] = "";
      }
    } else {
      const el = document.getElementById(`val-${field.key}`);
      values[field.key] = el ? el.value.trim() : "";
    }
  });
  return values;
}

function formatDate(isoDate, format) {
  try {
    const [year, month, day] = isoDate.split("-").map(Number);
    const d = new Date(year, month - 1, day);
    const pad = (n) => String(n).padStart(2, "0");
    switch (format) {
      case "abbr":
        return d.toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric" });
      case "short-us":
        return `${pad(month)}/${pad(day)}/${year}`;
      case "short-intl":
        return `${pad(day)}/${pad(month)}/${year}`;
      case "iso":
        return isoDate;
      case "long":
      default:
        return d.toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
    }
  } catch {
    return isoDate;
  }
}

// ── Clear Form ─────────────────────────────────────────────────────────────────

async function clearForm() {
  if (hasFilled && Object.keys(lastFilledValues).length > 0) {
    showClearConfirm();
    return;
  }
  doFormClear();
}

function showClearConfirm() {
  const el = document.getElementById("status");
  el.innerHTML = `
    <div style="margin-bottom:6px;font-weight:600">Reset all filled fields?</div>
    <div style="margin-bottom:10px;font-size:12px;color:#64748b">All filled values will be replaced with their original {{placeholders}}. Other edits you made to the document will be preserved.</div>
    <div style="display:flex;gap:8px">
      <button onclick="confirmReset()" style="flex:1;padding:7px 0;background:#dc2626;color:#fff;border:none;border-radius:7px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer">Reset All Fields</button>
      <button onclick="hideStatus()" style="padding:7px 12px;background:none;border:1.5px solid #bfdbfe;border-radius:7px;font-family:inherit;font-size:12px;color:#1d4ed8;cursor:pointer">Cancel</button>
    </div>
  `;
  el.className = "info";
  el.style.display = "block";
  el.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

async function confirmReset() {
  hideStatus();
  const clearBtn = document.querySelector(".btn-clear");
  if (clearBtn) { clearBtn.disabled = true; clearBtn.textContent = "Resetting..."; }

  try {
    await Word.run(async (context) => {
      // Phase 1: replace filled text with placeholders inside content controls
      for (const key of Object.keys(lastFilledValues)) {
        const ccs = context.document.contentControls.getByTag(key);
        ccs.load("items");
        await context.sync();
        for (const cc of ccs.items) {
          cc.insertText(`{{${key}}}`, Word.InsertLocation.replace);
        }
        await context.sync();
      }

      // Phase 2: remove content control wrappers (must be a separate sync)
      for (const key of Object.keys(lastFilledValues)) {
        const ccs = context.document.contentControls.getByTag(key);
        ccs.load("items");
        await context.sync();
        for (const cc of ccs.items) {
          cc.delete(true); // keep the placeholder text, remove the control wrapper
        }
        await context.sync();
      }
    });
    hasFilled = false;
  } catch (err) {
    showStatus("Failed to reset document: " + err.message, "error");
    if (clearBtn) { clearBtn.disabled = false; clearBtn.textContent = "Reset All Fields"; }
    return;
  }

  if (clearBtn) { clearBtn.disabled = false; clearBtn.textContent = "Reset All Fields"; }
  doFormClear();
}

async function resetField(key) {
  const filledValue = lastFilledValues[key];
  if (!filledValue) return;

  const resetBtn = document.getElementById(`reset-btn-${key}`);
  if (resetBtn) resetBtn.disabled = true;

  try {
    let found = false;
    await Word.run(async (context) => {
      // Find content controls tagged with this placeholder key
      const ccs = context.document.contentControls.getByTag(key);
      ccs.load("items");
      await context.sync();

      if (ccs.items.length > 0) {
        found = true;
        // Phase 1: replace text
        for (const cc of ccs.items) {
          cc.insertText(`{{${key}}}`, Word.InsertLocation.replace);
        }
        await context.sync();
        // Phase 2: remove control wrappers (separate sync)
        const ccs2 = context.document.contentControls.getByTag(key);
        ccs2.load("items");
        await context.sync();
        for (const cc of ccs2.items) {
          cc.delete(true);
        }
        await context.sync();
      }
    });

    if (found) {
      delete lastFilledValues[key];
      const input = document.getElementById(`val-${key}`);
      if (input) {
        // Check if this is a date field (container with date-select children)
        const dateSelects = input.querySelectorAll?.(".date-select");
        if (dateSelects && dateSelects.length > 0) {
          dateSelects.forEach((sel) => { sel.value = ""; });
        } else {
          input.value = "";
        }
      }
      if (resetBtn) { resetBtn.style.display = "none"; resetBtn.disabled = false; }
      if (Object.keys(lastFilledValues).length === 0) hasFilled = false;
    } else {
      if (resetBtn) resetBtn.disabled = false;
      showStatus("Could not find this field in the document — it may have been edited directly.", "error");
    }
  } catch (err) {
    if (resetBtn) resetBtn.disabled = false;
    showStatus("Error resetting field: " + err.message, "error");
  }
}

function doFormClear() {
  document.querySelectorAll(".field-value-input, .field-value-textarea").forEach((el) => {
    el.value = "";
  });
  // Clear date dropdowns back to placeholder
  document.querySelectorAll(".date-dropdowns").forEach((container) => {
    container.querySelectorAll(".date-select").forEach((sel) => { sel.value = ""; });
  });
  document.querySelectorAll(".field-row.field-empty").forEach((r) => r.classList.remove("field-empty"));
  document.querySelectorAll(".field-reset-btn").forEach((btn) => { btn.style.display = "none"; });
  lastFilledValues = {};
  hideStatus();
}

// ── localStorage ───────────────────────────────────────────────────────────────

const LS_PREFIX = "template-filler:";

/** A short fingerprint of the document to scope persistence. */
let documentFingerprint = "";

/**
 * Build a stable fingerprint from the document's template text.
 * Includes body + all header/footer text.
 * Strips both {{placeholder}} patterns and content-control text (filled values)
 * so the fingerprint is the same before and after filling.
 */
async function computeDocumentFingerprint() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");

      // Load content controls to identify filled regions
      const allCCs = context.document.contentControls;
      allCCs.load("items,text");

      // Load header/footer text
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      const hfBodies = [];
      for (const section of sections.items) {
        for (const hfType of getHfTypes()) {
          const h = section.getHeader(hfType);
          const f = section.getFooter(hfType);
          h.load("text");
          f.load("text");
          hfBodies.push(h, f);
        }
      }
      await context.sync();

      // Combine all text sources
      let raw = body.text || "";
      for (const hf of hfBodies) {
        if (hf.text && hf.text.trim()) raw += " " + hf.text;
      }

      // Remove filled values (content control text) so fingerprint is stable after fill
      for (const cc of allCCs.items) {
        if (cc.text) raw = raw.replace(cc.text, "");
      }

      // Strip {{placeholders}} and whitespace, take first 300 chars
      const stripped = raw.replace(/\{\{\w+\}\}/g, "").replace(/\s+/g, " ").trim().substring(0, 300);

      // djb2 hash
      let hash = 5381;
      for (let i = 0; i < stripped.length; i++) {
        hash = ((hash << 5) + hash + stripped.charCodeAt(i)) >>> 0;
      }
      documentFingerprint = hash.toString(36);
    });
  } catch {
    documentFingerprint = "";
  }
}

function buildStorageKey(keys) {
  const base = [...keys].sort().join(",");
  return documentFingerprint
    ? LS_PREFIX + documentFingerprint + ":" + base
    : LS_PREFIX + base;
}

/**
 * Try to load configs from the fingerprinted key first.
 * If nothing found and a fingerprint is set, fall back to the old un-fingerprinted key
 * and migrate the data forward.
 */
function loadFieldConfigsWithMigration(fingerprintedKey, keys) {
  let data = loadFieldConfigs(fingerprintedKey);
  if (Object.keys(data).length > 0) return data;

  // Fall back to legacy key (no fingerprint)
  if (documentFingerprint) {
    const legacyKey = LS_PREFIX + [...keys].sort().join(",");
    data = loadFieldConfigs(legacyKey);
    if (Object.keys(data).length > 0) {
      // Migrate to fingerprinted key
      try { localStorage.setItem(fingerprintedKey, JSON.stringify(data)); } catch { /* ignore */ }
      return data;
    }
  }
  return {};
}

function loadFieldConfigs(storageKey) {
  try { return JSON.parse(localStorage.getItem(storageKey) || "{}"); }
  catch { return {}; }
}

function saveFieldConfigs(storageKey, fields) {
  const data = {};
  fields.forEach((f) => {
    const entry = { label: f.label, type: f.type };
    if (f.dateFormat) entry.dateFormat = f.dateFormat;
    data[f.key] = entry;
  });
  try { localStorage.setItem(storageKey, JSON.stringify(data)); } catch { /* ignore */ }
}

// ── UI Helpers ─────────────────────────────────────────────────────────────────

function showStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = type;
  el.style.display = "block";
  el.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function hideStatus() {
  document.getElementById("status").style.display = "none";
}

function setScanButtonLoading(loading) {
  const btn = document.getElementById("scan-btn-empty");
  if (!btn) return;
  btn.disabled = loading;
  btn.innerHTML = loading ? '<span class="spinner dark"></span> Scanning...' : "Scan Document";
}

// ── String Utilities ───────────────────────────────────────────────────────────

function toTitleCase(str) {
  return str
    .replace(/_/g, " ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function escapeAttr(str) {
  return String(str).replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

// ── Tab Switching ──────────────────────────────────────────────────────────────

function switchTab(tab) {
  if (tab === activeTab) return;
  activeTab = tab;

  document.getElementById("tab-btn-create").classList.toggle("active", tab === "create");
  document.getElementById("tab-btn-fill").classList.toggle("active", tab === "fill");

  document.getElementById("panel-create").style.display = tab === "create" ? "flex" : "none";
  document.getElementById("panel-fill").style.display = tab === "fill" ? "block" : "none";

  if (tab === "create") {
    document.getElementById("actions").style.display = "none";
    fetchCurrentSelection();
  } else if (tab === "fill" && currentFields.length > 0) {
    document.getElementById("actions").style.display = "flex";
  }
}

// ── Selection Monitoring ───────────────────────────────────────────────────────

function onSelectionChanged() {
  if (activeTab !== "create") return;
  clearTimeout(selectionDebounceTimer);
  selectionDebounceTimer = setTimeout(fetchCurrentSelection, 250);
}

async function fetchCurrentSelection() {
  if (activeTab !== "create" || selectionFetchInProgress) return;
  selectionFetchInProgress = true;
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      const text = sel.text.trim();
      // Ignore multi-paragraph selections (Word returns \r between paragraphs)
      lastSelectedText = text.includes("\r") || text.includes("\n") ? "" : text;
      lastSelectedOccurrenceIndex = -1;

      // If text exists and might have multiple occurrences, find which one is selected
      // Search all bodies (body + headers/footers) to match the replacement domain
      if (lastSelectedText && lastSelectedText.length > 0) {
        const rawItems = await searchAllBodies(context, lastSelectedText, { matchCase: true });
        const items = await dedupeRanges(context, rawItems);

        if (items.length > 1) {
          // Compare each result with the selection to find the matching one
          for (let i = 0; i < items.length; i++) {
            const loc = sel.compareLocationWith(items[i]);
            await context.sync();
            if (loc.value === Word.LocationRelation.equal ||
                loc.value === Word.LocationRelation.contains ||
                loc.value === Word.LocationRelation.inside ||
                loc.value === "Equal" || loc.value === "Contains" || loc.value === "Inside") {
              lastSelectedOccurrenceIndex = i;
              break;
            }
          }
        }
      }

      updateSelectionPreview(lastSelectedText);
    });
  } catch {
    // ignore — selection may change during async fetch
  } finally {
    selectionFetchInProgress = false;
  }
}

function updateSelectionPreview(text) {
  const preview = document.getElementById("selection-preview");
  const nameInput = document.getElementById("placeholder-name-input");
  if (!preview || !nameInput) return;

  if (!text) {
    preview.className = "selection-preview";
    preview.innerHTML = '<span class="selection-hint-text">Select text in your document to get started</span>';
    return;
  }

  const display = text.length > 60 ? text.substring(0, 60) + "…" : text;
  preview.className = "selection-preview has-selection";
  preview.innerHTML = `<span class="selection-label">Selected</span><span class="selection-text">"${escapeHtml(display)}"</span>`;

  const suggested = suggestPlaceholderName(text);
  if (!nameInput.value || nameInput.value === lastSuggestedName) {
    nameInput.value = suggested;
    lastSuggestedName = suggested;
  }
}

function suggestPlaceholderName(text) {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/_+$/, "")
    .substring(0, 40);
}

// ── Create Placeholder ─────────────────────────────────────────────────────────

async function createPlaceholder() {
  const text = lastSelectedText;
  const nameInput = document.getElementById("placeholder-name-input");
  const name = nameInput.value.trim();

  if (!text) {
    showCreateStatus("Select some text in the document first.", "error");
    return;
  }
  if (!name) {
    showCreateStatus("Enter a placeholder name.", "error");
    nameInput.focus();
    return;
  }
  if (!/^\w+$/.test(name)) {
    showCreateStatus("Use only letters, numbers, and underscores.", "error");
    return;
  }

  const btn = document.getElementById("create-replace-btn");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Replacing...';
  hideCreateStatus();

  let shouldProceed = true;
  let occurrenceCount = 0;

  try {
    await Word.run(async (context) => {
      const rawItems = await searchAllBodies(context, text, { matchCase: true });
      const items = await dedupeRanges(context, rawItems);

      occurrenceCount = items.length;

      if (occurrenceCount === 0) {
        shouldProceed = false;
        return;
      }

      if (occurrenceCount > 1) {
        shouldProceed = false;
        pendingCreateText = text;
        pendingCreateName = name;
        pendingCreateIndex = lastSelectedOccurrenceIndex;
        showReplaceAllConfirm(occurrenceCount, name, lastSelectedOccurrenceIndex);
        return;
      }

      items[0].insertText(`{{${name}}}`, Word.InsertLocation.replace);
      await context.sync();
    });

    if (!shouldProceed && occurrenceCount === 0) {
      showCreateStatus("Could not find that text in the document — try selecting it again.", "error");
    } else if (shouldProceed) {
      onPlaceholderCreated(name, occurrenceCount);
    }
  } catch (err) {
    showCreateStatus("Error: " + err.message, "error");
  }

  btn.disabled = false;
  btn.innerHTML = "Replace with Placeholder";
}

function showReplaceAllConfirm(count, name, selectedIndex) {
  const el = document.getElementById("create-status");
  const singleLabel = selectedIndex >= 0
    ? `This occurrence (#${selectedIndex + 1})`
    : "First occurrence only";
  el.innerHTML = `
    <div style="margin-bottom:8px">Found <strong>${count} occurrences</strong> of this text. Replace with <code>{{${escapeHtml(name)}}}</code>?</div>
    <div style="display:flex;gap:6px;flex-wrap:wrap">
      <button onclick="confirmReplace(false)" style="flex:1;padding:6px 0;background:#2563eb;color:#fff;border:none;border-radius:6px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;min-width:80px">${singleLabel}</button>
      <button onclick="confirmReplace(true)" style="flex:1;padding:6px 0;background:#2563eb;color:#fff;border:none;border-radius:6px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;min-width:80px">All ${count} occurrences</button>
      <button onclick="hideCreateStatus()" style="padding:6px 10px;background:none;border:1.5px solid #bfdbfe;border-radius:6px;font-family:inherit;font-size:12px;color:#1d4ed8;cursor:pointer">Cancel</button>
    </div>
  `;
  el.className = "info";
  el.style.display = "block";
}

async function confirmReplace(replaceAll) {
  hideCreateStatus();

  const text = pendingCreateText;
  const name = pendingCreateName;
  const targetIndex = pendingCreateIndex;
  pendingCreateText = "";
  pendingCreateName = "";
  pendingCreateIndex = -1;
  if (!text || !name) return;

  const btn = document.getElementById("create-replace-btn");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Replacing...';

  try {
    let count = 0;
    await Word.run(async (context) => {
      if (replaceAll) {
        // Process each body sequentially so linked headers don't double-replace
        const bodies = await getAllBodies(context);
        for (const b of bodies) {
          const results = b.search(text, { matchCase: true });
          results.load("items");
          await context.sync();
          if (results.items.length > 0) {
            count += results.items.length;
            results.items.forEach((item) => item.insertText(`{{${name}}}`, Word.InsertLocation.replace));
            await context.sync();
          }
        }
      } else {
        // Single occurrence: dedupe linked ranges then replace one
        const rawItems = await searchAllBodies(context, text, { matchCase: true });
        const items = await dedupeRanges(context, rawItems);
        if (items.length === 0) return;
        const idx = (targetIndex >= 0 && targetIndex < items.length) ? targetIndex : 0;
        count = 1;
        items[idx].insertText(`{{${name}}}`, Word.InsertLocation.replace);
        await context.sync();
      }
    });
    if (count > 0) onPlaceholderCreated(name, count);
  } catch (err) {
    showCreateStatus("Error: " + err.message, "error");
  }

  btn.disabled = false;
  btn.innerHTML = "Replace with Placeholder";
}

function onPlaceholderCreated(name, count) {
  const nameInput = document.getElementById("placeholder-name-input");
  nameInput.value = "";
  lastSuggestedName = "";
  lastSelectedText = "";
  updateSelectionPreview("");

  // Merge into existing entry if the same placeholder was created again (e.g. after Ctrl+Z)
  const existing = createdPlaceholders.find((e) => e.name === name);
  if (existing) {
    existing.count += count;
  } else {
    createdPlaceholders.push({ name, count });
  }

  renderCreatedList();
  showCreateStatus(
    `✓ Created {{${name}}}${count > 1 ? ` — replaced ${count} occurrences` : ""}.`,
    "success"
  );
}

function renderCreatedList() {
  const section = document.getElementById("created-list-section");
  const list = document.getElementById("created-list");
  const doneBtn = document.getElementById("done-fill-btn");

  if (createdPlaceholders.length === 0) {
    section.style.display = "none";
    doneBtn.style.display = "none";
    return;
  }

  section.style.display = "block";
  doneBtn.style.display = "block";
  list.innerHTML = createdPlaceholders
    .map((e) => `<span class="created-chip" onclick="navigateToChip('${escapeAttr(e.name)}')" title="Click to highlight in document">{{${escapeHtml(e.name)}}}${e.count > 1 ? `<span class="chip-count">×${e.count}</span>` : ""}</span>`)
    .join("");
}

async function navigateToChip(name) {
  const idx = chipNavIndex[name] || 0;
  try {
    await Word.run(async (context) => {
      const rawItems = await searchAllBodies(context, `{{${name}}}`, { matchCase: true });
      const items = await dedupeRanges(context, rawItems);

      if (items.length === 0) {
        showCreateStatus(`{{${name}}} not found — it may have been filled or removed.`, "error");
        return;
      }

      // Sync count if doc has more/fewer occurrences than tracked (e.g. user added one manually)
      const entry = createdPlaceholders.find((e) => e.name === name);
      if (entry && entry.count !== items.length) {
        entry.count = items.length;
        renderCreatedList();
      }

      const targetIdx = idx % items.length;
      items[targetIdx].select();
      await context.sync();

      chipNavIndex[name] = (targetIdx + 1) % items.length;

      if (items.length > 1) {
        showCreateStatus(`{{${name}}} — occurrence ${targetIdx + 1} of ${items.length}`, "info");
      } else {
        hideCreateStatus();
      }
    });
  } catch (err) {
    showCreateStatus("Error: " + err.message, "error");
  }
}

function switchToFill() {
  switchTab("fill");
  scanDocument();
}

// ── Create Status ──────────────────────────────────────────────────────────────

function showCreateStatus(msg, type) {
  const el = document.getElementById("create-status");
  el.textContent = msg;
  el.className = type;
  el.style.display = "block";
}

function hideCreateStatus() {
  const el = document.getElementById("create-status");
  if (el) el.style.display = "none";
}
