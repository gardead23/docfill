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
  const raw = tag.startsWith(DOCFILL_TAG_PREFIX) ? tag.slice(DOCFILL_TAG_PREFIX.length) : tag;
  return raw.toLowerCase();
}
function keyToCCTag(key) {
  return DOCFILL_TAG_PREFIX + key.toLowerCase();
}

/** Check if text looks like a placeholder pattern (any casing). */
function isPlaceholderText(text) {
  return /^\{\{\w+\}\}$/.test(text.trim());
}

/** Check if text is a placeholder for a specific key (case-insensitive). */
function isPlaceholderTextForKey(text, key) {
  const m = text.trim().match(/^\{\{(\w+)\}\}$/);
  return m !== null && m[1].toLowerCase() === key.toLowerCase();
}

// ── State ─────────────────────────────────────────────────────────────────────

/** @type {{ key: string, label: string, type: string, dateFormat?: string }[]} */
let currentFields = [];
let currentStorageKey = "";
let hasFilled = false;
/** True after at least one full scan (body + HF) has completed. */
let hasScannedOnce = false;
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
/** True while background HF scan is running */
let hfScanInProgress = false;

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

let scanInProgress = false;

async function scanDocument() {
  if (scanInProgress || hfScanInProgress) return;
  scanInProgress = true;
  setScanButtonLoading(true);

  let hadExistingCCs = false;

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
      let ccCount = 0;
      for (const cc of allCCs.items) {
        if (!isDocFillCC(cc)) continue;
        ccCount++;
        const key = ccTagToKey(cc.tag);
        if (!ccsByKey[key]) ccsByKey[key] = [];
        ccsByKey[key].push(cc);
      }
      hadExistingCCs = ccCount > 0;

      // ── Phase B: Discover raw {{key}} text and convert to CCs ──
      // Quick check: does the body text we already loaded contain any raw {{}} patterns?
      // Body text was loaded in Phase A (allCCs sync also loaded body).
      // Load main body text cheaply to check.
      const mainBody = context.document.body;
      mainBody.load("text");
      await context.sync();

      const bodyText = mainBody.text || "";
      const bodyMatches = bodyText.match(/\{\{(\w+)\}\}/g) || [];
      // Canonicalize to unique lowercase keys, search once per key with matchCase:false
      const keysInBody = [...new Set(bodyMatches.map((m) => m.replace(/\{\{|\}\}/g, "").toLowerCase()))];

      let convertedAny = false;

      if (keysInBody.length > 0) {
        const searches = {};
        for (const key of keysInBody) {
          searches[key] = mainBody.search(`{{${key}}}`, { matchCase: false });
          searches[key].load("items");
        }
        await context.sync();

        // Batch parent-CC checks
        const rangeEntries = [];
        for (const key of keysInBody) {
          for (const range of searches[key].items) {
            const parentCC = range.parentContentControlOrNullObject;
            parentCC.load("tag");
            rangeEntries.push({ key, range, parentCC });
          }
        }
        if (rangeEntries.length > 0) {
          await context.sync();
          for (const { key, range, parentCC } of rangeEntries) {
            const insideExisting = !parentCC.isNullObject &&
              parentCC.tag && parentCC.tag.startsWith(DOCFILL_TAG_PREFIX);
            if (!insideExisting) {
              if (!ccsByKey[key]) ccsByKey[key] = [];
              convertRangeToCC(range, key);
              convertedAny = true;
            }
          }
          if (convertedAny) await context.sync();
        }
      }

      // Reload CCs after body conversion
      if (convertedAny) {
        allCCs.load("items,tag,text");
        await context.sync();
      }

      // Rebuild CC map from current state
      const ccMap = {};
      for (const cc of allCCs.items) {
        if (!isDocFillCC(cc)) continue;
        const key = ccTagToKey(cc.tag);
        if (!ccMap[key]) ccMap[key] = { items: [], text: cc.text };
        ccMap[key].items.push(cc);
      }

      // ── Phase C: Build field list and hydrate state ──
      // Render form immediately from body + existing CCs (fast path)
      const allKeys = Object.keys(ccMap);

      if (allKeys.length === 0 && keysInBody.length === 0) {
        currentFields = [];
        currentStorageKey = "";
        document.getElementById("fields-section").style.display = "none";
        document.getElementById("actions").style.display = "none";
        document.getElementById("empty-state").style.display = "block";
        document.querySelector(".empty-desc").innerHTML =
          'No <code>{{placeholders}}</code> found. Add fields like <code>{{client_name}}</code> to your document, then scan again.';
        setScanButtonLoading(false);
        return; // scanInProgress cleared in outer finally
      }

      // Hydrate lastFilledValues from CCs
      lastFilledValues = {};
      for (const [key, data] of Object.entries(ccMap)) {
        const text = data.text.trim();
        // A CC is "filled" if its text is not its own placeholder pattern
        if (text && !isPlaceholderTextForKey(text, key)) {
          lastFilledValues[key] = data.text;
        }
      }
      hasFilled = Object.keys(lastFilledValues).length > 0;

      // Use document order: allKeys comes from ccMap which is built from
      // contentControls in document order (first occurrence of each key)
      const keys = allKeys;

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

      // Scroll to very top of fields section (above HF status and date format)
      const fieldsSection = document.getElementById("fields-section");
      if (fieldsSection) {
        fieldsSection.scrollIntoView({ behavior: "instant", block: "start" });
      }
    });
  } catch (err) {
    showStatus("Error reading document: " + err.message, "error");
  }

  setScanButtonLoading(false);

  // ── Deferred HF scan: always check for new HF placeholders ──
  // The text-load is cheap; only bodies with raw {{}} are processed further.
  try {
    const hfStatusEl = document.getElementById("hf-status");
    const fillBtn = document.getElementById("fill-btn");
    if (hfStatusEl) {
      hfStatusEl.innerHTML = `
        <div class="scan-banner-content">
          <div class="scan-banner-header">You can enter values now.</div>
          <div class="scan-banner-body">We're finalizing the document setup in the background. This can take up to 30 seconds for large templates, but you can fill out the fields below while you wait. The final "Fill Document" button will unlock when setup finishes.</div>
        </div>
        <div class="scan-banner-progress"><div class="scan-banner-progress-bar"></div></div>`;
      hfStatusEl.style.display = "block";
    }
    if (fillBtn) { fillBtn.disabled = true; fillBtn.textContent = "Finishing setup..."; }
    scanHeaderFooters().then(() => {
      if (hfStatusEl) hfStatusEl.style.display = "none";
      if (fillBtn) { fillBtn.disabled = false; fillBtn.innerHTML = "Fill Document"; }
      hasScannedOnce = true;
    });
  } finally {
    scanInProgress = false;
  }
}

/** Scan headers/footers for raw {{key}} text and convert to CCs. Runs after main scan. */
async function scanHeaderFooters() {
  hfScanInProgress = true;
  try {
    let foundNew = false;
    await Word.run(async (context) => {
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

      const relevantHfBodies = hfBodies.filter((b) => {
        const t = b.text && b.text.trim();
        return t && /\{\{\w+\}\}/.test(t);
      });

      if (relevantHfBodies.length === 0) return;

      for (const b of relevantHfBodies) {
        try {
          const hfText = b.text || "";
          const hfMatches = hfText.match(/\{\{(\w+)\}\}/g) || [];
          const hfKeys = [...new Set(hfMatches.map((m) => m.replace(/\{\{|\}\}/g, "").toLowerCase()))];
          if (hfKeys.length === 0) continue;

          const hfSearches = {};
          for (const key of hfKeys) {
            hfSearches[key] = b.search(`{{${key}}}`, { matchCase: false });
            hfSearches[key].load("items");
          }
          await context.sync();

          const rangeEntries = [];
          for (const key of hfKeys) {
            for (const range of hfSearches[key].items) {
              const parentCC = range.parentContentControlOrNullObject;
              parentCC.load("tag");
              rangeEntries.push({ key, range, parentCC });
            }
          }
          if (rangeEntries.length === 0) continue;
          await context.sync();

          let converted = false;
          for (const { key, range, parentCC } of rangeEntries) {
            const insideExisting = !parentCC.isNullObject &&
              parentCC.tag && parentCC.tag.startsWith(DOCFILL_TAG_PREFIX);
            if (!insideExisting) {
              convertRangeToCC(range, key);
              converted = true;
              foundNew = true;
            }
          }
          if (converted) await context.sync();
        } catch (bodyErr) {
          if (bodyErr.code !== "GeneralException") {
            console.warn("DocFill: error scanning a header/footer:", bodyErr.message || bodyErr);
          }
        }
      }
    });

    if (foundNew) {
      // Preserve raw draft values (including date select states) before rerendering
      const draftValues = collectRawDrafts();

      // Reload CCs to pick up newly converted HF fields
      await Word.run(async (context) => {
        const allCCs = context.document.contentControls;
        allCCs.load("items,tag,text");
        await context.sync();

        const ccMap = {};
        for (const cc of allCCs.items) {
          if (!isDocFillCC(cc)) continue;
          const key = ccTagToKey(cc.tag);
          if (!ccMap[key]) ccMap[key] = { items: [], text: cc.text };
          ccMap[key].items.push(cc);
        }

        const allKeys = Object.keys(ccMap);
        // Hydrate filled values
        for (const [key, data] of Object.entries(ccMap)) {
          const text = data.text.trim();
          if (text && !isPlaceholderTextForKey(text, key)) {
            lastFilledValues[key] = data.text;
          }
        }

        hasFilled = Object.keys(lastFilledValues).length > 0;
        const keys = allKeys;

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

        // Restore draft values (user-typed but not yet filled) into form inputs only
        restoreRawDrafts(draftValues);
      });
    }
  } catch {
    // HF scan is best-effort
  } finally {
    hfScanInProgress = false;
  }
}

// ── Render Form ────────────────────────────────────────────────────────────────

function renderForm(fields) {
  document.getElementById("empty-state").style.display = "none";
  document.getElementById("fields-section").style.display = "block";
  // Only show Fill/Reset footer when Fill tab is active
  if (activeTab === "fill") {
    document.getElementById("actions").style.display = "flex";
  }

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

      // Update all CCs in one batch.
      // Replace \n with \v (vertical tab = soft line break in Word) so
      // multi-line text stays within the same paragraph and doesn't push
      // surrounding inline text onto separate lines.
      for (const key of keys) {
        const value = toFill[key].replace(/\n/g, "\v");
        for (const cc of ccCollections[key].items) {
          cc.insertText(value, Word.InsertLocation.replace);
        }
        totalReplaced += ccCollections[key].items.length;
      }

      if (totalReplaced > 0) await context.sync();
    });

    if (totalReplaced === 0) {
      // CCs may have been removed by Ctrl+Z or manual deletion. Auto-rescan.
      btn.disabled = false;
      btn.innerHTML = "Fill Document";
      showStatus("Fields not found in document. Rescanning...", "info");
      await scanDocument();
      return;
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

/** Collect raw draft values from ALL visible form fields, including already-filled ones the user may have edited. */
function collectRawDrafts() {
  const drafts = {};
  currentFields.forEach((field) => {
    if (field.type === "date") {
      const container = document.getElementById(`val-${field.key}`);
      if (container) {
        const m = container.querySelector(".date-month")?.value || "";
        const d = container.querySelector(".date-day")?.value || "";
        const y = container.querySelector(".date-year")?.value || "";
        if (m || d || y) drafts[field.key] = { type: "date", month: m, day: d, year: y };
      }
    } else {
      const el = document.getElementById(`val-${field.key}`);
      const val = el ? el.value : "";
      if (val.trim()) drafts[field.key] = { type: "text", value: val };
    }
  });
  return drafts;
}

/** Restore raw draft values into form inputs after a rerender. Overrides CC-hydrated values since drafts represent user edits. */
function restoreRawDrafts(drafts) {
  for (const [key, draft] of Object.entries(drafts)) {
    if (draft.type === "date") {
      const container = document.getElementById(`val-${key}`);
      if (container) {
        const monthSel = container.querySelector(".date-month");
        const daySel = container.querySelector(".date-day");
        const yearSel = container.querySelector(".date-year");
        if (monthSel) monthSel.value = draft.month;
        if (yearSel) yearSel.value = draft.year;
        if (draft.month || draft.year) updateDayOptions(key);
        if (daySel) daySel.value = draft.day;
      }
    } else {
      const el = document.getElementById(`val-${key}`);
      if (el) el.value = draft.value;
    }
  }
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
  // Scroll back to top of form
  const fieldsSection = document.getElementById("fields-section");
  if (fieldsSection) {
    fieldsSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }
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
      await context.sync();

      let raw = body.text || "";

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

  // Try legacy key (no fingerprint), both exact and lowercase
  const legacyKey = LS_PREFIX + [...keys].sort().join(",");
  data = loadFieldConfigs(legacyKey);
  if (Object.keys(data).length > 0) {
    try { localStorage.setItem(fingerprintedKey, JSON.stringify(data)); } catch { /* ignore */ }
    return data;
  }

  // Case-insensitive scan: check both fingerprinted and legacy targets
  try {
    const targets = [fingerprintedKey.toLowerCase(), legacyKey.toLowerCase()];
    for (let i = 0; i < localStorage.length; i++) {
      const lsKey = localStorage.key(i);
      if (lsKey && lsKey.startsWith(LS_PREFIX) && targets.includes(lsKey.toLowerCase())) {
        data = loadFieldConfigs(lsKey);
        if (Object.keys(data).length > 0) {
          try { localStorage.setItem(fingerprintedKey, JSON.stringify(data)); } catch { /* ignore */ }
          return data;
        }
      }
    }
  } catch { /* ignore */ }

  return {};
}

function loadFieldConfigs(storageKey) {
  try {
    const raw = JSON.parse(localStorage.getItem(storageKey) || "{}");
    // Normalize keys to lowercase for case-insensitive migration
    const normalized = {};
    for (const [k, v] of Object.entries(raw)) {
      normalized[k.toLowerCase()] = v;
    }
    return normalized;
  }
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
    document.getElementById("create-actions").style.display = "flex";
    initCreateTab();
  } else if (tab === "fill") {
    document.getElementById("create-actions").style.display = "none";
    if (currentFields.length > 0) {
      document.getElementById("actions").style.display = "flex";
    }
    if (!hasScannedOnce) {
      // First visit: full scan
      scanDocument();
    } else {
      // Quick check: any new raw placeholders in body not yet converted?
      checkForNewPlaceholders();
    }
  }
}

/**
 * Quick check for new placeholders. If found, does a fast body-only
 * scan (converts new raw text, rebuilds form from CCs). No HF loading.
 */
async function checkForNewPlaceholders() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      const allCCs = context.document.contentControls;
      allCCs.load("items,tag,text");
      await context.sync();

      // Build CC key set
      const ccKeys = new Set();
      for (const cc of allCCs.items) {
        if (isDocFillCC(cc)) ccKeys.add(ccTagToKey(cc.tag));
      }

      // Check if anything changed since last render
      const currentKeys = new Set(currentFields.map((f) => f.key));
      let needsUpdate = false;

      // Check 1: CCs not in current field list (created via Create tab)
      for (const key of ccKeys) {
        if (!currentKeys.has(key)) { needsUpdate = true; break; }
      }

      // Check 2: search raw body patterns and convert any NOT inside CCs.
      // Canonicalize to lowercase keys, search once per key with matchCase:false.
      const bodyText = body.text || "";
      const rawMatches = bodyText.match(/\{\{(\w+)\}\}/g) || [];
      const rawKeys = [...new Set(rawMatches.map((m) => m.replace(/\{\{|\}\}/g, "").toLowerCase()))];

      if (rawKeys.length > 0) {
        const searches = {};
        for (const key of rawKeys) {
          searches[key] = body.search(`{{${key}}}`, { matchCase: false });
          searches[key].load("items");
        }
        await context.sync();

        // Batch parent-CC checks
        const rangeEntries = [];
        for (const key of rawKeys) {
          for (const range of searches[key].items) {
            const parentCC = range.parentContentControlOrNullObject;
            parentCC.load("tag");
            rangeEntries.push({ key, range, parentCC });
          }
        }
        if (rangeEntries.length > 0) {
          await context.sync();
          for (const { key, range, parentCC } of rangeEntries) {
            const insideExisting = !parentCC.isNullObject &&
              parentCC.tag && parentCC.tag.startsWith(DOCFILL_TAG_PREFIX);
            if (!insideExisting) {
              convertRangeToCC(range, key);
              needsUpdate = true; // only rerender if we actually converted something
            }
          }
          if (needsUpdate) await context.sync();
        }

        if (needsUpdate) {
          // Reload CCs after conversion
          allCCs.load("items,tag,text");
          await context.sync();
        }
      }

      if (!needsUpdate) return;

      // Preserve draft values before rerendering
      const draftValues = collectRawDrafts();

      // Rebuild form from all CCs (same as Phase C of scanDocument)
      const ccMap = {};
      for (const cc of allCCs.items) {
        if (!isDocFillCC(cc)) continue;
        const key = ccTagToKey(cc.tag);
        if (!ccMap[key]) ccMap[key] = { items: [], text: cc.text };
        ccMap[key].items.push(cc);
      }

      const allKeys = Object.keys(ccMap);
      if (allKeys.length === 0) return;

      lastFilledValues = {};
      for (const [key, data] of Object.entries(ccMap)) {
        const text = data.text.trim();
        if (text && !isPlaceholderTextForKey(text, key)) lastFilledValues[key] = data.text;
      }
      hasFilled = Object.keys(lastFilledValues).length > 0;

      const keys = allKeys;
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
      restoreRawDrafts(draftValues);

      const fieldsSection = document.getElementById("fields-section");
      if (fieldsSection) fieldsSection.scrollIntoView({ behavior: "instant", block: "start" });
    });
  } catch {
    // Best effort
  }
}

/** Initialize Create tab: show status, load placeholders, then start selection monitoring. */
async function initCreateTab() {
  const statusEl = document.getElementById("create-loading-status");

  // Show placeholders immediately from body text (fast)
  await loadExistingPlaceholders();

  // Kick off full document scan only if not already done.
  const needsScan = !hasScannedOnce && !scanInProgress && !hfScanInProgress;
  const hfRunning = hfScanInProgress;

  if (needsScan || hfRunning) {
    if (statusEl) {
      statusEl.innerHTML = `
        <div class="scan-banner-content">
          <div class="scan-banner-header">Syncing document fields...</div>
          <div class="scan-banner-body">We're scanning your template to set up your fields. For large templates, this background process can take up to 30 seconds.</div>
        </div>
        <div class="scan-banner-progress"><div class="scan-banner-progress-bar"></div></div>`;
      statusEl.style.display = "block";
    }

    // Start scan if needed (fires HF scan in background)
    if (needsScan) scanDocument();

    // Wait for everything to finish (scan + HF scan), then reload and clear status
    const waitForCompletion = async () => {
      while (scanInProgress || hfScanInProgress) {
        await new Promise((r) => setTimeout(r, 200));
      }
      await loadExistingPlaceholders();
      if (statusEl) statusEl.style.display = "none";
    };
    waitForCompletion();
  }

  // Start selection monitoring
  fetchCurrentSelection();
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
  const name = nameInput.value.trim().toLowerCase();

  if (!text) { showCreateStatus("Select some text in the document first.", "error"); return; }
  if (!name) { showCreateStatus("Enter a placeholder name.", "error"); nameInput.focus(); return; }
  if (!/^\w+$/.test(name)) { showCreateStatus("Use only letters, numbers, and underscores.", "error"); return; }

  const btn = document.getElementById("create-replace-btn");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Replacing...';
  hideCreateStatus();

  let shouldProceed = true;
  let exactCount = 0;
  let allCount = 0;

  try {
    await Word.run(async (context) => {
      // Exact-case search
      const exactRaw = await searchAllBodies(context, text, { matchCase: true });
      const exactItems = await dedupeRanges(context, exactRaw);
      exactCount = exactItems.length;

      // Case-insensitive search for variants
      const allRaw = await searchAllBodies(context, text, { matchCase: false });
      const allItems = await dedupeRanges(context, allRaw);
      allCount = allItems.length;

      const variantCount = allCount - exactCount;

      if (allCount === 0) { shouldProceed = false; return; }

      if (exactCount === 1 && variantCount === 0) {
        // Single exact match, no variants: replace immediately
        const cc = exactItems[0].insertContentControl();
        cc.tag = keyToCCTag(name);
        cc.title = toTitleCase(name);
        cc.appearance = Word.ContentControlAppearance.boundingBox;
        cc.placeholderText = `{{${name}}}`;
        cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
        await context.sync();
      } else {
        // Multiple matches or variants: show confirmation
        shouldProceed = false;
        pendingCreateText = text;
        pendingCreateName = name;
        pendingCreateIndex = lastSelectedOccurrenceIndex;
        showReplaceAllConfirm(exactCount, allCount, name, lastSelectedOccurrenceIndex);
      }
    });

    if (!shouldProceed && allCount === 0) {
      showCreateStatus("Could not find that text in the document -- try selecting it again.", "error");
    } else if (shouldProceed) {
      onPlaceholderCreated(name, 1);
    }
  } catch (err) {
    showCreateStatus("Error: " + err.message, "error");
  }

  btn.disabled = false;
  btn.innerHTML = "Replace with Placeholder";
}

function showReplaceAllConfirm(exactCount, allCount, name, selectedIndex) {
  const el = document.getElementById("create-status");
  const variantCount = allCount - exactCount;
  const btnStyle = 'flex:1;padding:6px 0;background:#2563eb;color:#fff;border:none;border-radius:6px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;min-width:80px';
  const cancelStyle = 'padding:6px 10px;background:none;border:1.5px solid #bfdbfe;border-radius:6px;font-family:inherit;font-size:12px;color:#1d4ed8;cursor:pointer';

  let description;
  let buttons;

  if (variantCount === 0) {
    // Only exact matches
    description = `Found <strong>${exactCount} matches</strong>. Replace with <code>{{${escapeHtml(name)}}}</code>?`;
    const singleLabel = selectedIndex >= 0 ? `This one (#${selectedIndex + 1})` : "This one only";
    buttons = `
      <button onclick="confirmReplace('single')" style="${btnStyle}">${singleLabel}</button>
      <button onclick="confirmReplace('exact')" style="${btnStyle}">All ${exactCount} matches</button>
      <button onclick="hideCreateStatus()" style="${cancelStyle}">Cancel</button>`;
  } else if (exactCount === 0) {
    // Only variant matches (selected text not found exact, but variants exist)
    description = `Found <strong>${allCount} match${allCount > 1 ? "es" : ""}</strong> with different capitalization. Replace with <code>{{${escapeHtml(name)}}}</code>?`;
    buttons = `
      <button onclick="confirmReplace('all')" style="${btnStyle}">All ${allCount} match${allCount > 1 ? "es" : ""}</button>
      <button onclick="hideCreateStatus()" style="${cancelStyle}">Cancel</button>`;
  } else if (exactCount === 1 && variantCount > 0) {
    // 1 exact + variants
    description = `Found <strong>1</strong> "${escapeHtml(pendingCreateText)}" and <strong>${variantCount} more</strong> with different capitalization. Replace with <code>{{${escapeHtml(name)}}}</code>?`;
    buttons = `
      <button onclick="confirmReplace('single')" style="${btnStyle}">This one only</button>
      <button onclick="confirmReplace('all')" style="${btnStyle}">Both matches</button>
      <button onclick="hideCreateStatus()" style="${cancelStyle}">Cancel</button>`;
    if (allCount > 2) {
      buttons = `
        <button onclick="confirmReplace('single')" style="${btnStyle}">This one only</button>
        <button onclick="confirmReplace('all')" style="${btnStyle}">All ${allCount} matches</button>
        <button onclick="hideCreateStatus()" style="${cancelStyle}">Cancel</button>`;
    }
  } else {
    // Multiple exact + variants
    description = `Found <strong>${exactCount}</strong> same-capitalization matches and <strong>${variantCount} more</strong> with different capitalization. Replace with <code>{{${escapeHtml(name)}}}</code>?`;
    buttons = `
      <button onclick="confirmReplace('single')" style="${btnStyle}">This one only</button>
      <button onclick="confirmReplace('exact')" style="${btnStyle}">All ${exactCount} same-capitalization</button>
      <button onclick="confirmReplace('all')" style="${btnStyle}">All ${allCount} matches</button>
      <button onclick="hideCreateStatus()" style="${cancelStyle}">Cancel</button>`;
  }

  el.innerHTML = `
    <div style="margin-bottom:8px">${description}</div>
    <div style="display:flex;gap:6px;flex-wrap:wrap">${buttons}</div>`;
  el.className = "info";
  el.style.display = "block";
}

/** @param {'single'|'exact'|'all'} mode */
async function confirmReplace(mode) {
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
      if (mode === "single") {
        // Single occurrence: use exact-case search, pick the targeted index
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
      } else {
        // 'exact' or 'all': replace multiple matches
        const rawItems = await searchAllBodies(context, text, { matchCase: mode !== "all" });
        const items = await dedupeRanges(context, rawItems);

        for (const range of items) {
          const cc = range.insertContentControl();
          cc.tag = keyToCCTag(name);
          cc.title = toTitleCase(name);
          cc.appearance = Word.ContentControlAppearance.boundingBox;
          cc.placeholderText = `{{${name}}}`;
          cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
        }
        count += items.length;
        if (count > 0) await context.sync();
      }
    });
    if (count > 0) {
      onPlaceholderCreated(name, count);
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

/** Load all existing placeholders (DocFill CCs + raw {{text}}) and populate the list. */
async function loadExistingPlaceholders() {
  try {
    await Word.run(async (context) => {
      const allCCs = context.document.contentControls;
      allCCs.load("items,tag");
      const body = context.document.body;
      body.load("text");
      await context.sync();

      // Count from DocFill CCs
      const counts = {};
      const ccKeys = new Set();
      for (const cc of allCCs.items) {
        if (!isDocFillCC(cc)) continue;
        const key = ccTagToKey(cc.tag);
        ccKeys.add(key);
        counts[key] = (counts[key] || 0) + 1;
      }

      // Count raw {{key}} patterns not already covered by CCs
      const bodyText = body.text || "";
      const rawMatches = bodyText.match(/\{\{(\w+)\}\}/g) || [];
      for (const m of rawMatches) {
        const key = m.replace(/\{\{|\}\}/g, "").toLowerCase();
        if (!ccKeys.has(key)) {
          counts[key] = (counts[key] || 0) + 1;
        }
      }

      createdPlaceholders = Object.entries(counts).map(([name, count]) => ({ name, count }));
      renderCreatedList();
    });
  } catch {
    // Best effort
  }
}

function renderCreatedList(filter) {
  const section = document.getElementById("created-list-section");
  const list = document.getElementById("created-list");
  const countEl = document.getElementById("created-list-count");
  if (createdPlaceholders.length === 0) { section.style.display = "none"; return; }
  section.style.display = "block";
  if (countEl) countEl.textContent = `(${createdPlaceholders.length})`;

  const filterLower = (filter || "").toLowerCase();
  const filtered = filterLower
    ? createdPlaceholders.filter((e) => e.name.includes(filterLower))
    : createdPlaceholders;

  list.innerHTML = filtered.map((e) => `
    <div class="created-row" onclick="navigateToChip('${escapeAttr(e.name)}')" title="Click to highlight in document">
      <span class="created-row-name">{{${escapeHtml(e.name)}}}</span>
      <span class="created-row-badge">${e.count}</span>
      <span class="created-row-actions">
        <button class="created-row-action delete" onclick="event.stopPropagation(); deleteCreatedPlaceholder('${escapeAttr(e.name)}')" title="Delete placeholder">&#128465;</button>
      </span>
    </div>`).join("");
}

function filterCreatedList(query) {
  renderCreatedList(query);
}

/** Delete a placeholder: remove all its CCs from the document and refresh the list. */
async function deleteCreatedPlaceholder(name) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(keyToCCTag(name));
      ccs.load("items");
      await context.sync();

      // Replace CC text with the original selected text, then remove CC
      for (const cc of ccs.items) {
        cc.insertText(cc.text || name, Word.InsertLocation.replace);
        cc.delete(true);
      }
      await context.sync();
    });

    // Remove from list and re-render
    createdPlaceholders = createdPlaceholders.filter((e) => e.name !== name);
    renderCreatedList(document.getElementById("created-list-search")?.value);
  } catch (err) {
    showCreateStatus("Error deleting placeholder: " + err.message, "error");
  }
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
  // switchTab("fill") handles refresh via checkForNewPlaceholders or first-scan
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
