/* global Office, Word */

"use strict";

// ── DocFill Tag Convention ───────────────────────────────────────────────────
// Every DocFill content control has tag = "docfill:{key}".
// This distinguishes DocFill fields from other CCs in the document.

const DOCFILL_TAG_PREFIX = "docfill:";

function isDocFillCC(cc) {
  return !!(cc && cc.tag && cc.tag.startsWith(DOCFILL_TAG_PREFIX));
}
function ccTagToKey(tag) {
  return tag.startsWith(DOCFILL_TAG_PREFIX) ? tag.slice(DOCFILL_TAG_PREFIX.length) : tag;
}
function keyToCCTag(key) {
  return DOCFILL_TAG_PREFIX + key;
}

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
const chipNavIndex = {};
let selectionDebounceTimer = null;
let selectionFetchInProgress = false;
let lastSelectedOccurrenceIndex = -1;

// ── Document Range Helpers ────────────────────────────────────────────────────

let HF_TYPES = null;
function getHfTypes() {
  if (!HF_TYPES) {
    HF_TYPES = [Word.HeaderFooterType.primary, Word.HeaderFooterType.firstPage, Word.HeaderFooterType.evenPages];
  }
  return HF_TYPES;
}

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
  for (const b of hfBodies) b.load("text");
  await context.sync();

  for (const b of hfBodies) {
    if (b.text && b.text.trim()) bodies.push(b);
  }
  return bodies;
}

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

async function dedupeRanges(context, ranges) {
  if (ranges.length <= 1) return ranges;
  const comparisons = [];
  for (let i = 1; i < ranges.length; i++) {
    for (let j = 0; j < i; j++) {
      comparisons.push({ i, j, result: ranges[i].compareLocationWith(ranges[j]) });
    }
  }
  await context.sync();
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

// ── DocFill CC Helpers ───────────────────────────────────────────────────────

/**
 * Get all DocFill content controls, grouped by key.
 * Returns { key: { items: [cc, ...], text: string } }
 */
async function getAllDocFillCCs(context) {
  const allCCs = context.document.contentControls;
  allCCs.load("items,tag,text");
  await context.sync();

  const result = {};
  for (const cc of allCCs.items) {
    if (!isDocFillCC(cc)) continue;
    const key = ccTagToKey(cc.tag);
    if (!result[key]) result[key] = { items: [], text: cc.text };
    result[key].items.push(cc);
  }
  return result;
}

/**
 * Migrate old-style CCs (tag = raw key, no prefix) to docfill: prefix.
 * Only migrates if the tag looks like a placeholder key and has no prefix.
 */
function migrateOldCC(cc) {
  if (cc.tag && !cc.tag.startsWith(DOCFILL_TAG_PREFIX) && /^\w+$/.test(cc.tag)) {
    cc.tag = keyToCCTag(cc.tag);
    return true;
  }
  return false;
}

/**
 * Convert a raw {{key}} text range into a DocFill content control.
 * The CC wraps the range, with placeholder text shown inside.
 */
function convertRangeToCC(range, key) {
  const cc = range.insertContentControl();
  cc.tag = keyToCCTag(key);
  cc.title = toTitleCase(key);
  cc.appearance = Word.ContentControlAppearance.boundingBox;
  cc.placeholderText = `{{${key}}}`;
  return cc;
}

// ── Scan Document ──────────────────────────────────────────────────────────────

async function scanDocument() {
  showStatus("Scanning document...", "info");
  setScanButtonLoading(true);

  try {
    await computeDocumentFingerprint();

    await Word.run(async (context) => {
      // ── Phase A: Discover and migrate existing CCs ──
      const allCCs = context.document.contentControls;
      allCCs.load("items,tag,text");
      await context.sync();

      // Migrate old-style CCs (no docfill: prefix)
      let migrated = false;
      for (const cc of allCCs.items) {
        if (migrateOldCC(cc)) migrated = true;
      }
      if (migrated) await context.sync();

      // Build map of DocFill CCs by key
      const ccsByKey = {};
      for (const cc of allCCs.items) {
        if (!isDocFillCC(cc)) continue;
        const key = ccTagToKey(cc.tag);
        if (!ccsByKey[key]) ccsByKey[key] = [];
        ccsByKey[key].push(cc);
      }

      // ── Phase B: Discover raw {{key}} text and convert to CCs ──
      // Collect all body text to find placeholder keys via JS regex
      const bodies = await getAllBodies(context);
      let allText = "";
      for (const b of bodies) b.load("text");
      await context.sync();
      for (const b of bodies) allText += " " + (b.text || "");

      const rawMatches = allText.match(/\{\{(\w+)\}\}/g) || [];
      // Search ALL keys found in text (not just new ones -- user may have added
      // another {{client_name}} after an earlier scan)
      const keysToSearch = [...new Set(rawMatches.map((m) => m.replace(/\{\{|\}\}/g, "")))];
      let convertedAny = false;

      if (keysToSearch.length > 0) {
        // Search main body, then header/footer bodies sequentially.
        // Sequential processing handles linked headers: once a placeholder is
        // converted in one body, the linked copy no longer has raw text.
        // After each body that finds results, sync so conversions take effect.
        const mainBody = bodies[0]; // getAllBodies always returns document.body first
        const hfBodies = bodies.slice(1).filter((b) => {
          const t = b.text || "";
          return keysToSearch.some((k) => t.includes(`{{${k}}}`));
        });

        for (const b of [mainBody, ...hfBodies]) {
          try {
            const searches = {};
            for (const key of keysToSearch) {
              searches[key] = b.search(`{{${key}}}`, { matchCase: true });
              searches[key].load("items");
            }
            await context.sync();

            // Collect all ranges and batch-check parent CCs in one sync
            const rangeEntries = []; // { key, range, parentCC }
            for (const key of keysToSearch) {
              for (const range of searches[key].items) {
                const parentCC = range.parentContentControlOrNullObject;
                parentCC.load("tag");
                rangeEntries.push({ key, range, parentCC });
              }
            }
            if (rangeEntries.length === 0) continue;
            await context.sync(); // single sync for all parent-CC checks

            let foundInBody = false;
            for (const { key, range, parentCC } of rangeEntries) {
              const insideExisting = !parentCC.isNullObject &&
                parentCC.tag && parentCC.tag.startsWith(DOCFILL_TAG_PREFIX);
              if (!insideExisting) {
                if (!ccsByKey[key]) ccsByKey[key] = [];
                convertRangeToCC(range, key);
                foundInBody = true;
                convertedAny = true;
              }
            }
            // Sync after each body so linked copies see the conversion
            if (foundInBody) await context.sync();
          } catch (bodyErr) {
            if (bodyErr.code !== "GeneralException") {
              console.warn("DocFill: error scanning a region:", bodyErr.message || bodyErr);
            }
          }
        }
      }

      // Reload CCs after conversion
      allCCs.load("items,tag,text");
      await context.sync();

      // Rebuild CC map
      const ccMap = {};
      for (const cc of allCCs.items) {
        if (!isDocFillCC(cc)) continue;
        const key = ccTagToKey(cc.tag);
        if (!ccMap[key]) ccMap[key] = { items: [], text: cc.text };
        ccMap[key].items.push(cc);
      }

      // ── Phase C: Build field list and hydrate state ──
      const allKeys = Object.keys(ccMap);

      if (allKeys.length === 0) {
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

      // Hydrate lastFilledValues from CCs
      lastFilledValues = {};
      for (const [key, data] of Object.entries(ccMap)) {
        const text = data.text.trim();
        // A CC is "filled" if its text is not the placeholder pattern and not empty
        if (text && text !== `{{${key}}}`) {
          lastFilledValues[key] = data.text;
        }
      }
      hasFilled = Object.keys(lastFilledValues).length > 0;

      // Preserve field order; append new keys at end
      const orderedExisting = currentFields.map((f) => f.key).filter((k) => allKeys.includes(k));
      const brandNewKeys = allKeys.filter((k) => !currentFields.some((f) => f.key === k));
      const keys = [...orderedExisting, ...brandNewKeys];

      currentStorageKey = buildStorageKey(keys);
      const saved = loadFieldConfigsWithMigration(currentStorageKey, keys);

      currentFields = keys.map((key) => {
        const savedType = saved[key]?.type === "number" ? "text" : saved[key]?.type;
        return {
          key,
          label: saved[key]?.label || toTitleCase(key),
          type: savedType || guessFieldType(key),
          dateFormat: saved[key]?.dateFormat,
        };
      });

      saveFieldConfigs(currentStorageKey, currentFields);
      renderForm(currentFields);
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

    const fieldType = field.type === "number" ? "text" : field.type;
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

    // Restore filled value and show reset button
    if (lastFilledValues[field.key]) {
      if (field.type !== "date") {
        const input = row.querySelector(".field-value-input, .field-value-textarea");
        if (input) input.value = lastFilledValues[field.key];
      }
      const resetBtn = document.getElementById(`reset-btn-${field.key}`);
      if (resetBtn) resetBtn.style.display = "inline-flex";
    }
  });

  renderGlobalDateFormat(fields);
}

function renderGlobalDateFormat(fields) {
  const container = document.getElementById("global-date-format");
  if (!container) return;
  const hasDateFields = fields.some((f) => f.type === "date");
  if (!hasDateFields) { container.style.display = "none"; return; }
  const current = getGlobalDateFormat();
  container.style.display = "flex";
  container.innerHTML = `
    <label class="global-date-label" for="global-date-select">Default date format</label>
    <select id="global-date-select" class="date-format-select" onchange="onGlobalDateFormatChange(this.value)">
      ${DATE_FORMATS.map((f) => `<option value="${f.value}" ${f.value === current ? "selected" : ""}>${f.label}</option>`).join("")}
    </select>
  `;
}

function onGlobalDateFormatChange(format) { setGlobalDateFormat(format); }

function buildValueInput(field) {
  const id = `val-${field.key}`;
  if (field.type === "paragraph") {
    return `<textarea id="${id}" class="field-value-textarea" placeholder="Enter ${escapeHtml(field.label).toLowerCase()}..." rows="3"></textarea>`;
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
      <select class="date-format-select" id="datefmt-${field.key}" onchange="setFieldDateFormat('${escapeAttr(field.key)}', this.value)" title="Date output format">
        <option value="" ${!fieldFmt ? "selected" : ""}>Default (${formatDatePreview(globalFmt)})</option>
        ${DATE_FORMATS.map((f) => `<option value="${f.value}" ${fieldFmt === f.value ? "selected" : ""}>${f.label}</option>`).join("")}
      </select>
    </div>`;
  }
  return `<input id="${id}" class="field-value-input" type="text" placeholder="Enter ${escapeHtml(field.label).toLowerCase()}..." />`;
}

// ── Field Edit Handlers ────────────────────────────────────────────────────────

function setFieldType(key, newType) {
  const field = currentFields.find((f) => f.key === key);
  if (!field || field.type === newType) return;
  const oldValue = document.getElementById(`val-${key}`)?.value || "";
  field.type = newType;
  if (newType !== "date") delete field.dateFormat;
  saveFieldConfigs(currentStorageKey, currentFields);
  const row = document.querySelector(`.field-row[data-key="${key}"]`);
  if (!row) return;
  row.querySelectorAll(".field-value-input, .field-value-textarea, .date-dropdowns, .date-format-row").forEach((el) => el.remove());
  row.insertAdjacentHTML("beforeend", buildValueInput(field));
  if (newType !== "date") {
    const newInput = row.querySelector(".field-value-input, .field-value-textarea");
    if (newInput) newInput.value = oldValue;
  }
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

function daysInMonth(month, year) {
  if (!month) return 31;
  if (!year) year = new Date().getFullYear();
  return new Date(year, month, 0).getDate();
}

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
  let html = '<option value="">Day</option>';
  for (let d = 1; d <= maxDay; d++) html += `<option value="${d}">${d}</option>`;
  daySel.innerHTML = html;
  if (currentDay > 0) daySel.value = currentDay > maxDay ? maxDay : currentDay;
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

function guessFieldType(key) {
  const k = key.toLowerCase();
  if (/date|day|month|year|when|start|end|deadline|due|expir|signed|effective/.test(k)) return "date";
  if (/description|notes?|bio|summary|detail|scope|address|comments?|message|body|terms/.test(k)) return "paragraph";
  return "text";
}

// ── Fill Document ──────────────────────────────────────────────────────────────
// All fields are DocFill CCs after scan. Fill simply updates CC text.

async function fillDocument() {
  const btn = document.getElementById("fill-btn");
  const allValues = collectValues();

  const toFill = Object.fromEntries(Object.entries(allValues).filter(([, v]) => v.trim()));
  const emptyKeys = Object.keys(allValues).filter((k) => !allValues[k].trim());

  document.querySelectorAll(".field-row.field-empty").forEach((r) => r.classList.remove("field-empty"));

  if (Object.keys(toFill).length === 0) {
    showStatus("Fill in at least one field to continue.", "error");
    return;
  }

  emptyKeys.forEach((key) => {
    document.querySelector(`.field-row[data-key="${key}"]`)?.classList.add("field-empty");
  });

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Filling...';
  hideStatus();

  try {
    let totalReplaced = 0;

    await Word.run(async (context) => {
      const keys = Object.keys(toFill);

      // Batch-load all CC collections for all keys in one sync
      const ccCollections = {};
      for (const key of keys) {
        ccCollections[key] = context.document.contentControls.getByTag(keyToCCTag(key));
        ccCollections[key].load("items");
      }
      await context.sync();

      // Update all CCs in one batch
      for (const key of keys) {
        for (const cc of ccCollections[key].items) {
          cc.insertText(toFill[key], Word.InsertLocation.replace);
        }
        totalReplaced += ccCollections[key].items.length;
      }

      if (totalReplaced > 0) await context.sync();
    });

    if (totalReplaced === 0) {
      showStatus("No fields found. Try scanning the document first.", "error");
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
        showStatus(`Done. Highlighted fields were skipped: ${skipped}`, "info");
        const firstEmpty = document.querySelector(".field-row.field-empty");
        if (firstEmpty) firstEmpty.scrollIntoView({ behavior: "smooth", block: "nearest" });
      } else {
        showStatus("All fields filled successfully.", "success");
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
          const mi = parseInt(m, 10), di = parseInt(d, 10), yi = parseInt(y, 10);
          const maxDay = daysInMonth(mi, yi);
          const pad = (n) => String(n).padStart(2, "0");
          const safeDay = di > maxDay ? maxDay : di;
          const isoDate = `${yi}-${pad(mi)}-${pad(safeDay)}`;
          values[field.key] = formatDate(isoDate, field.dateFormat || globalFmt);
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

// ── Reset ─────────────────────────────────────────────────────────────────────
// Reset replaces CC text with {{key}} placeholder. CCs are NEVER deleted.

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
      // Batch-load all CC collections
      const keysToReset = Object.keys(lastFilledValues);
      const ccCollections = {};
      for (const key of keysToReset) {
        ccCollections[key] = context.document.contentControls.getByTag(keyToCCTag(key));
        ccCollections[key].load("items");
      }
      await context.sync();

      // Replace all CC text with placeholder text in one batch
      for (const key of keysToReset) {
        for (const cc of ccCollections[key].items) {
          cc.insertText(`{{${key}}}`, Word.InsertLocation.replace);
        }
      }
      await context.sync();
    });
    hasFilled = false;
  } catch (err) {
    showStatus("Failed to reset: " + err.message, "error");
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
      const ccs = context.document.contentControls.getByTag(keyToCCTag(key));
      ccs.load("items");
      await context.sync();

      if (ccs.items.length > 0) {
        found = true;
        for (const cc of ccs.items) {
          cc.insertText(`{{${key}}}`, Word.InsertLocation.replace);
        }
        await context.sync();
      }
    });

    if (found) {
      delete lastFilledValues[key];
      const input = document.getElementById(`val-${key}`);
      if (input) {
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
      showStatus("Could not find this field in the document.", "error");
    }
  } catch (err) {
    if (resetBtn) resetBtn.disabled = false;
    showStatus("Error resetting field: " + err.message, "error");
  }
}

function doFormClear() {
  document.querySelectorAll(".field-value-input, .field-value-textarea").forEach((el) => { el.value = ""; });
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
let documentFingerprint = "";

async function computeDocumentFingerprint() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      const allCCs = context.document.contentControls;
      allCCs.load("items,tag,text");
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

      let raw = body.text || "";
      for (const hf of hfBodies) {
        if (hf.text && hf.text.trim()) raw += " " + hf.text;
      }

      // Strip DocFill CC text and placeholder patterns for stability
      for (const cc of allCCs.items) {
        if (isDocFillCC(cc) && cc.text) raw = raw.replace(cc.text, "");
      }
      const stripped = raw.replace(/\{\{\w+\}\}/g, "").replace(/\s+/g, " ").trim().substring(0, 300);

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
  return documentFingerprint ? LS_PREFIX + documentFingerprint + ":" + base : LS_PREFIX + base;
}

function loadFieldConfigsWithMigration(fingerprintedKey, keys) {
  let data = loadFieldConfigs(fingerprintedKey);
  if (Object.keys(data).length > 0) return data;
  if (documentFingerprint) {
    const legacyKey = LS_PREFIX + [...keys].sort().join(",");
    data = loadFieldConfigs(legacyKey);
    if (Object.keys(data).length > 0) {
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
      lastSelectedText = text.includes("\r") || text.includes("\n") ? "" : text;
      lastSelectedOccurrenceIndex = -1;

      if (lastSelectedText && lastSelectedText.length > 0) {
        const rawItems = await searchAllBodies(context, lastSelectedText, { matchCase: true });
        const items = await dedupeRanges(context, rawItems);
        if (items.length > 1) {
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
  } catch { /* ignore */ }
  finally { selectionFetchInProgress = false; }
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
  const display = text.length > 60 ? text.substring(0, 60) + "\u2026" : text;
  preview.className = "selection-preview has-selection";
  preview.innerHTML = `<span class="selection-label">Selected</span><span class="selection-text">"${escapeHtml(display)}"</span>`;
  const suggested = suggestPlaceholderName(text);
  if (!nameInput.value || nameInput.value === lastSuggestedName) {
    nameInput.value = suggested;
    lastSuggestedName = suggested;
  }
}

function suggestPlaceholderName(text) {
  return text.toLowerCase().replace(/[^a-z0-9\s]/g, "").trim()
    .replace(/\s+/g, "_").replace(/_+$/, "").substring(0, 40);
}

// ── Create Placeholder ─────────────────────────────────────────────────────────
// Create mode now inserts DocFill CCs directly instead of raw {{text}}.

async function createPlaceholder() {
  const text = lastSelectedText;
  const nameInput = document.getElementById("placeholder-name-input");
  const name = nameInput.value.trim();

  if (!text) { showCreateStatus("Select some text in the document first.", "error"); return; }
  if (!name) { showCreateStatus("Enter a placeholder name.", "error"); nameInput.focus(); return; }
  if (!/^\w+$/.test(name)) { showCreateStatus("Use only letters, numbers, and underscores.", "error"); return; }

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

      if (occurrenceCount === 0) { shouldProceed = false; return; }

      if (occurrenceCount > 1) {
        shouldProceed = false;
        pendingCreateText = text;
        pendingCreateName = name;
        pendingCreateIndex = lastSelectedOccurrenceIndex;
        showReplaceAllConfirm(occurrenceCount, name, lastSelectedOccurrenceIndex);
        return;
      }

      // Single occurrence: replace text and wrap in DocFill CC
      const cc = items[0].insertContentControl();
      cc.tag = keyToCCTag(name);
      cc.title = toTitleCase(name);
      cc.appearance = Word.ContentControlAppearance.boundingBox;
      cc.placeholderText = `{{${name}}}`;
      cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
      await context.sync();
    });

    if (!shouldProceed && occurrenceCount === 0) {
      showCreateStatus("Could not find that text in the document -- try selecting it again.", "error");
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
    let createSkipped = 0;
    await Word.run(async (context) => {
      if (replaceAll) {
        const bodies = await getAllBodies(context);
        for (const b of bodies) {
          try {
            const results = b.search(text, { matchCase: true });
            results.load("items");
            await context.sync();
            if (results.items.length > 0) {
              for (const range of results.items) {
                const cc = range.insertContentControl();
                cc.tag = keyToCCTag(name);
                cc.title = toTitleCase(name);
                cc.appearance = Word.ContentControlAppearance.boundingBox;
                cc.placeholderText = `{{${name}}}`;
                cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
              }
              count += results.items.length;
              await context.sync();
            }
          } catch (bodyErr) {
            if (bodyErr.code === "GeneralException") {
              // Linked header already modified -- expected
            } else {
              createSkipped++;
              console.warn("DocFill: error in create:", bodyErr.message || bodyErr);
            }
          }
        }
      } else {
        const rawItems = await searchAllBodies(context, text, { matchCase: true });
        const items = await dedupeRanges(context, rawItems);
        if (items.length === 0) return;
        const idx = (targetIndex >= 0 && targetIndex < items.length) ? targetIndex : 0;
        const cc = items[idx].insertContentControl();
        cc.tag = keyToCCTag(name);
        cc.title = toTitleCase(name);
        cc.appearance = Word.ContentControlAppearance.boundingBox;
        cc.placeholderText = `{{${name}}}`;
        cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
        count = 1;
        await context.sync();
      }
    });
    if (count > 0) {
      onPlaceholderCreated(name, count);
      if (createSkipped > 0) {
        showCreateStatus(
          `Created {{${name}}} but ${createSkipped} region${createSkipped > 1 ? "s were" : " was"} skipped.`,
          "info"
        );
      }
    }
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
  const existing = createdPlaceholders.find((e) => e.name === name);
  if (existing) { existing.count += count; } else { createdPlaceholders.push({ name, count }); }
  renderCreatedList();
  showCreateStatus(`Created {{${name}}}${count > 1 ? ` -- replaced ${count} occurrences` : ""}.`, "success");
}

function renderCreatedList() {
  const section = document.getElementById("created-list-section");
  const list = document.getElementById("created-list");
  const doneBtn = document.getElementById("done-fill-btn");
  if (createdPlaceholders.length === 0) { section.style.display = "none"; doneBtn.style.display = "none"; return; }
  section.style.display = "block";
  doneBtn.style.display = "block";
  list.innerHTML = createdPlaceholders
    .map((e) => `<span class="created-chip" onclick="navigateToChip('${escapeAttr(e.name)}')" title="Click to highlight in document">{{${escapeHtml(e.name)}}}${e.count > 1 ? `<span class="chip-count">\u00d7${e.count}</span>` : ""}</span>`)
    .join("");
}

async function navigateToChip(name) {
  const idx = chipNavIndex[name] || 0;
  try {
    await Word.run(async (context) => {
      // Navigate via DocFill CCs instead of text search
      const ccs = context.document.contentControls.getByTag(keyToCCTag(name));
      ccs.load("items");
      await context.sync();

      if (ccs.items.length === 0) {
        showCreateStatus(`{{${name}}} not found.`, "error");
        return;
      }

      const entry = createdPlaceholders.find((e) => e.name === name);
      if (entry && entry.count !== ccs.items.length) {
        entry.count = ccs.items.length;
        renderCreatedList();
      }

      const targetIdx = idx % ccs.items.length;
      ccs.items[targetIdx].select();
      await context.sync();
      chipNavIndex[name] = (targetIdx + 1) % ccs.items.length;

      if (ccs.items.length > 1) {
        showCreateStatus(`{{${name}}} -- occurrence ${targetIdx + 1} of ${ccs.items.length}`, "info");
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
