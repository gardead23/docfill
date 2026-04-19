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
let pendingMatchCount = 0;
let pendingMatchIndex = 0;
let pendingMatchCase = true;
let pendingMatchWholeWord = true;
let matchNavInFlight = false;
const chipNavIndex = {};
let selectionDebounceTimer = null;
let selectionFetchInProgress = false;
/** Suppresses selection preview updates during programmatic navigation. */
let suppressSelectionPreview = false;
/** Generation token for selection fetch -- incremented on chip navigation to cancel in-flight fetches. */
let selectionFetchGeneration = 0;
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
 * Sort CC map keys by document order.
 * Flattens all CCs into {key, cc} pairs, sorts them by range position,
 * then returns keys in order of their earliest occurrence.
 * Uses two syncs: one for within-key earliest detection, one for cross-key ordering.
 * Must be called inside Word.run.
 */
async function sortKeysByDocumentOrder(context, ccMap) {
  const keys = Object.keys(ccMap);
  if (keys.length <= 1) return keys;

  // For each key, pick one representative CC.
  // For single-CC keys, use that CC. For multi-CC keys, find the earliest.
  const representative = {};
  const multiKeys = [];

  for (const key of keys) {
    const items = ccMap[key].items;
    if (items.length === 1) {
      representative[key] = items[0];
    } else {
      representative[key] = items[0]; // default, may be updated
      multiKeys.push(key);
    }
  }

  // Find earliest CC per multi-occurrence key using pairwise comparison (one sync)
  if (multiKeys.length > 0) {
    const withinComps = [];
    for (const key of multiKeys) {
      const items = ccMap[key].items;
      for (let i = 0; i < items.length; i++) {
        for (let j = i + 1; j < items.length; j++) {
          withinComps.push({
            key, i, j,
            result: items[i].getRange().compareLocationWith(items[j].getRange())
          });
        }
      }
    }
    await context.sync();

    // Score each occurrence: higher = comes before more siblings
    const scores = {};
    for (const key of multiKeys) {
      scores[key] = new Array(ccMap[key].items.length).fill(0);
    }
    for (const { key, i, j, result } of withinComps) {
      const v = result.value;
      if (v === "Before" || v === "AdjacentBefore" ||
          v === Word.LocationRelation.before || v === Word.LocationRelation.adjacentBefore) {
        scores[key][i]++;
      } else if (v === "After" || v === "AdjacentAfter" ||
                 v === Word.LocationRelation.after || v === Word.LocationRelation.adjacentAfter) {
        scores[key][j]++;
      }
    }
    // Pick the occurrence with the highest score (earliest)
    for (const key of multiKeys) {
      let bestIdx = 0;
      for (let i = 1; i < scores[key].length; i++) {
        if (scores[key][i] > scores[key][bestIdx]) bestIdx = i;
      }
      representative[key] = ccMap[key].items[bestIdx];
    }
  }

  // Compare representative CCs across keys (one sync)
  const crossComps = [];
  for (let i = 0; i < keys.length; i++) {
    for (let j = i + 1; j < keys.length; j++) {
      crossComps.push({
        a: keys[i], b: keys[j],
        result: representative[keys[i]].getRange().compareLocationWith(representative[keys[j]].getRange())
      });
    }
  }

  if (crossComps.length === 0) return keys;
  await context.sync();

  const score = {};
  for (const key of keys) score[key] = 0;

  for (const { a, b, result } of crossComps) {
    const v = result.value;
    if (v === "Before" || v === "AdjacentBefore" ||
        v === Word.LocationRelation.before || v === Word.LocationRelation.adjacentBefore) {
      score[a]++;
    } else if (v === "After" || v === "AdjacentAfter" ||
               v === Word.LocationRelation.after || v === Word.LocationRelation.adjacentAfter) {
      score[b]++;
    }
  }

  // Higher score = comes before more keys = earlier in document
  return [...keys].sort((a, b) => score[b] - score[a]);
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
      // Sort keys by CC document order using Range.compareLocationWith
      const allKeys = await sortKeysByDocumentOrder(context, ccMap);

      if (allKeys.length === 0 && keysInBody.length === 0) {
        currentFields = [];
        currentStorageKey = "";
        document.getElementById("fields-section").style.display = "none";
        document.getElementById("actions").style.display = "none";
        document.getElementById("empty-state").style.display = "block";
        document.querySelector(".empty-desc").textContent =
          'No {{placeholders}} found. Type fields like {{client_name}} into your document, then scan again.';
        setScanButtonLoading(false);
        return; // scanInProgress cleared in outer finally
      }

      // Hydrate lastFilledValues from CCs
      lastFilledValues = {};
      for (const [key, data] of Object.entries(ccMap)) {
        const text = data.text.trim();
        // A CC is "filled" if its text is not its own placeholder pattern
        if (text && !isPlaceholderTextForKey(text, key)) {
          lastFilledValues[key] = data.text.replace(/\v/g, "\n");
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
    const importBtn = document.getElementById("fill-import-btn");
    if (importBtn) importBtn.disabled = true;
    closeImportPanel();
    scanHeaderFooters().then(() => {
      if (hfStatusEl) hfStatusEl.style.display = "none";
      if (fillBtn) { fillBtn.disabled = false; fillBtn.innerHTML = "Fill Document"; }
      if (importBtn) importBtn.disabled = false;
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

        const allKeys = await sortKeysByDocumentOrder(context, ccMap);
        // Hydrate filled values
        for (const [key, data] of Object.entries(ccMap)) {
          const text = data.text.trim();
          if (text && !isPlaceholderTextForKey(text, key)) {
            lastFilledValues[key] = data.text.replace(/\v/g, "\n");
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

let fillSortMode = "doc"; // "doc" or "az"
let fillFilterText = "";

function setFillSort(mode) {
  fillSortMode = mode;
  const btn = document.getElementById("fill-sort-btn");
  if (btn) btn.classList.toggle("active", mode === "az");
  document.getElementById("sort-opt-doc")?.classList.toggle("active", mode === "doc");
  document.getElementById("sort-opt-az")?.classList.toggle("active", mode === "az");
  closeSortMenu();
  applyFieldDisplayOrder();
}

function toggleSortMenu() {
  const menu = document.getElementById("fill-sort-menu");
  if (!menu) return;
  if (menu.style.display === "none") {
    menu.style.display = "block";
    // Close on outside click
    setTimeout(() => {
      document.addEventListener("click", closeSortMenuOnOutsideClick, { once: true });
    }, 0);
  } else {
    closeSortMenu();
  }
}

function closeSortMenu() {
  const menu = document.getElementById("fill-sort-menu");
  if (menu) menu.style.display = "none";
}

function closeSortMenuOnOutsideClick(e) {
  const wrap = document.querySelector(".fill-sort-wrap");
  if (wrap && !wrap.contains(e.target)) {
    closeSortMenu();
  }
}

function filterFillFields(query) {
  fillFilterText = query.toLowerCase();
  applyFieldDisplayOrder();
}

/**
 * Reorder and show/hide existing DOM field rows without rebuilding them.
 * This preserves all typed draft values.
 */
function applyFieldDisplayOrder() {
  const fieldsList = document.getElementById("fields-list");
  if (!fieldsList) return;

  const rows = Array.from(fieldsList.querySelectorAll(".field-row"));
  if (rows.length === 0) return;

  // Build ordered list of keys
  let orderedKeys = currentFields.map((f) => f.key);
  if (fillSortMode === "az") {
    orderedKeys = [...currentFields].sort((a, b) => a.label.localeCompare(b.label)).map((f) => f.key);
  }

  // Build a map of key -> DOM row
  const rowMap = {};
  for (const row of rows) {
    rowMap[row.dataset.key] = row;
  }

  // Reorder DOM nodes and apply filter visibility
  let visibleCount = 0;
  for (const key of orderedKeys) {
    const row = rowMap[key];
    if (!row) continue;

    // Filter
    if (fillFilterText) {
      const field = currentFields.find((f) => f.key === key);
      const matches = key.includes(fillFilterText) ||
        (field && field.label.toLowerCase().includes(fillFilterText));
      row.style.display = matches ? "" : "none";
      if (matches) visibleCount++;
    } else {
      row.style.display = "";
      visibleCount++;
    }

    // Move to end (reorders without destroying)
    fieldsList.appendChild(row);
  }

  // Show "no results" message if needed
  let noResults = fieldsList.querySelector(".fill-no-results");
  if (visibleCount === 0 && fillFilterText) {
    if (!noResults) {
      noResults = document.createElement("div");
      noResults.className = "fill-no-results";
      noResults.style.cssText = "padding:16px;text-align:center;color:#9ca3af;font-size:12px";
      noResults.textContent = "No fields match your search.";
      fieldsList.appendChild(noResults);
    }
    noResults.style.display = "";
  } else if (noResults) {
    noResults.style.display = "none";
  }
}

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
        <button
          type="button"
          class="field-label-link"
          id="label-${escapeAttr(field.key)}"
          onclick="navigateToChip('${escapeAttr(field.key)}')"
          title="Jump to {{${escapeAttr(field.key)}}} in document"
        >${escapeHtml(field.label)}</button>
        <button
          type="button"
          class="field-edit-btn"
          onclick="startLabelEdit('${escapeAttr(field.key)}')"
          title="Edit label"
        >
          <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
            <path d="M8.5 1.5l2 2-6.5 6.5H2V8L8.5 1.5z" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
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

  // Apply current sort/filter to the freshly built rows
  applyFieldDisplayOrder();
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

function buildResetBtn(key) {
  return `<button
    class="field-reset-btn"
    id="reset-btn-${escapeAttr(key)}"
    onclick="resetField('${escapeAttr(key)}')"
    title="Clear value"
    style="display:none"
  >
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
      <path d="M2 5.5A4 4 0 1 1 3.5 9" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
      <path d="M2 3v2.5h2.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>
  </button>`;
}

function buildValueInput(field) {
  const id = `val-${field.key}`;
  const resetBtn = buildResetBtn(field.key);
  if (field.type === "paragraph") {
    return `<div class="field-value-wrap">${resetBtn}<textarea id="${id}" class="field-value-textarea" placeholder="Enter ${escapeHtml(field.label).toLowerCase()}..." rows="3"></textarea></div>`;
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
    return `<div class="field-value-wrap">${resetBtn}<div class="date-dropdowns" id="${id}">
      <select class="date-select date-month" title="Month" onchange="updateDayOptions('${escapeAttr(field.key)}')"><option value="">Month</option>${monthOpts}</select>
      <select class="date-select date-day" title="Day"><option value="">Day</option>${dayOpts}</select>
      <select class="date-select date-year" title="Year" onchange="updateDayOptions('${escapeAttr(field.key)}')" ><option value="" selected>Year</option>${yearOpts}</select>
      <button type="button" class="date-today-btn" onclick="setDateToday('${escapeAttr(field.key)}')" title="Set to today">Today</button>
    </div></div>
    <div class="date-format-row">
      <span class="date-format-label">Format:</span>
      <select class="date-format-select" id="datefmt-${field.key}" onchange="setFieldDateFormat('${escapeAttr(field.key)}', this.value)" title="Date output format">
        <option value="" ${!fieldFmt ? "selected" : ""}>Default (${formatDatePreview(globalFmt)})</option>
        ${DATE_FORMATS.map((f) => `<option value="${f.value}" ${fieldFmt === f.value ? "selected" : ""}>${f.label}</option>`).join("")}
      </select>
    </div>`;
  }
  return `<div class="field-value-wrap">${resetBtn}<input id="${id}" class="field-value-input" type="text" placeholder="Enter ${escapeHtml(field.label).toLowerCase()}..." /></div>`;
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
  row.querySelectorAll(".field-value-wrap, .field-value-input, .field-value-textarea, .date-dropdowns, .date-format-row").forEach((el) => el.remove());
  row.insertAdjacentHTML("beforeend", buildValueInput(field));
  if (newType !== "date") {
    const newInput = row.querySelector(".field-value-input, .field-value-textarea");
    if (newInput) newInput.value = oldValue;
  }
  // Re-show reset button if this field has been filled
  if (lastFilledValues[key]) {
    const resetBtn = document.getElementById(`reset-btn-${key}`);
    if (resetBtn) resetBtn.style.display = "inline-flex";
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

function startLabelEdit(key) {
  // Cancel any existing edit first
  const existingEdit = document.querySelector(".field-label-edit-input");
  if (existingEdit && existingEdit.dataset.key) {
    finishLabelEdit(existingEdit.dataset.key, existingEdit.value);
  }

  const labelBtn = document.getElementById(`label-${key}`);
  if (!labelBtn) return;
  const field = currentFields.find((f) => f.key === key);
  if (!field) return;

  const input = document.createElement("input");
  input.type = "text";
  input.className = "field-label-edit-input";
  input.dataset.key = key;
  input.value = field.label;
  input.onblur = function () { finishLabelEdit(key, input.value); };
  input.onkeydown = function (e) {
    if (e.key === "Enter") { input.blur(); }
    if (e.key === "Escape") { input.value = field.label; input.blur(); }
  };

  labelBtn.replaceWith(input);
  input.focus();
  input.select();
}

function finishLabelEdit(key, newLabel) {
  const field = currentFields.find((f) => f.key === key);
  if (!field) return;

  const trimmed = newLabel.trim();
  if (trimmed) field.label = trimmed;
  saveFieldConfigs(currentStorageKey, currentFields);

  // Find the edit input for this specific field
  const input = document.querySelector(`.field-label-edit-input[data-key="${key}"]`);
  if (!input) return;

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "field-label-link";
  btn.id = `label-${key}`;
  btn.onclick = function () { navigateToChip(key); };
  btn.title = `Jump to {{${key}}} in document`;
  btn.textContent = field.label;

  input.replaceWith(btn);
}

function setFieldDateFormat(key, format) {
  const field = currentFields.find((f) => f.key === key);
  if (!field) return;
  field.dateFormat = format || undefined;
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

let fillInProgress = false;

async function fillDocument() {
  const btn = document.getElementById("fill-btn");
  const allValues = collectValues();

  const toFill = Object.fromEntries(Object.entries(allValues).filter(([, v]) => v.trim()));
  const emptyKeys = Object.keys(allValues).filter((k) => !allValues[k].trim());

  // If search is active and there are hidden empty fields, clear the filter to reveal them
  if (fillFilterText && emptyKeys.length > 0) {
    const searchInput = document.getElementById("fill-search");
    if (searchInput) searchInput.value = "";
    fillFilterText = "";
    applyFieldDisplayOrder();
  }

  document.querySelectorAll(".field-row.field-empty").forEach((r) => r.classList.remove("field-empty"));

  if (Object.keys(toFill).length === 0) {
    emptyKeys.forEach((key) => {
      document.querySelector(`.field-row[data-key="${key}"]`)?.classList.add("field-empty");
    });
    showStatus("Fill in at least one field to continue.", "error");
    scrollToFirstEmptyField();
    return;
  }

  emptyKeys.forEach((key) => {
    document.querySelector(`.field-row[data-key="${key}"]`)?.classList.add("field-empty");
  });

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Filling...';
  fillInProgress = true;
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
        scrollToFirstEmptyField();
      } else {
        showStatus("All fields filled successfully.", "success");
      }
    }
  } catch (err) {
    showStatus("Error: " + err.message, "error");
  } finally {
    fillInProgress = false;
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

// ── Import ────────────────────────────────────────────────────────────────────
// Import data from CSV file or pasted text to populate form fields.
// Populates the form only -- user still clicks "Fill Document" to apply.

const IMPORT_MAX_FILE_SIZE = 5 * 1024 * 1024; // 5MB
const IMPORT_ROW_CAP = 100;
let importAutoCloseTimer = null;
let importGeneration = 0;
let activeImportReader = null;
let pendingHorizontalHeaders = null;
let pendingHorizontalRows = null;
let pendingImportRawText = null;
let pendingImportSource = null; // "csv" | "paste"

const MONTH_NAMES = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];

function normalizeImportKey(rawKey) {
  return rawKey
    .trim()
    .replace(/^\{\{|\}\}$/g, "")
    .trim()
    .replace(/[\s-]+/g, "_")
    .toLowerCase()
    .replace(/[^a-z0-9_]/g, "");
}

function isHeaderRow(col1, col2) {
  const HEADER_WORDS = /^(key|field|name|placeholder|variable|value|header|column|label|data)$/i;
  return HEADER_WORDS.test(col1.trim()) && HEADER_WORDS.test(col2.trim());
}

function detectDelimiter(line) {
  const tabs = (line.match(/\t/g) || []).length;
  const commas = (line.match(/,/g) || []).length;
  return tabs >= commas && tabs > 0 ? "\t" : ",";
}

function parseCSVRaw(text) {
  if (!text || !text.trim()) return [];

  const allRows = [];
  let currentCell = "";
  let currentRow = [];
  let inQuotes = false;
  const chars = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < chars.length && chars[i + 1] === '"') {
          currentCell += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        currentCell += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ",") {
        currentRow.push(currentCell);
        currentCell = "";
      } else if (ch === "\n") {
        currentRow.push(currentCell);
        allRows.push(currentRow);
        currentRow = [];
        currentCell = "";
      } else {
        currentCell += ch;
      }
    }
  }
  currentRow.push(currentCell);
  if (currentRow.some((c) => c !== "")) {
    allRows.push(currentRow);
  }

  return allRows.filter((row) => row.some((c) => c !== ""));
}

function parsePastedRaw(text) {
  if (!text || !text.trim()) return { rows: [], delimiter: "," };

  const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const nonEmpty = lines.filter((l) => l.trim());
  if (nonEmpty.length === 0) return { rows: [], delimiter: "," };

  const delimiter = detectDelimiter(nonEmpty[0]);
  const rows = nonEmpty.map((line) => line.split(delimiter));

  return { rows, delimiter };
}

function parseCSV(text) {
  const allRows = parseCSVRaw(text);
  const rows = [];
  let skippedEmpty = 0;
  let isFirst = true;

  for (const cells of allRows) {
    if (cells.length < 2) continue;
    const key = cells[0].trim();
    const value = cells[1].trim();
    if (!key) continue;

    if (isFirst) {
      isFirst = false;
      if (isHeaderRow(key, value)) continue;
    }

    if (!value) {
      skippedEmpty++;
      continue;
    }

    rows.push({ key, value });
  }

  return { rows, skippedEmpty };
}

function parsePastedText(text) {
  const raw = parsePastedRaw(text);
  if (raw.rows.length === 0) return { rows: [], skippedEmpty: 0 };

  const rows = [];
  let skippedEmpty = 0;
  let isFirst = true;

  for (const parts of raw.rows) {
    if (parts.length < 2) continue;
    const key = parts[0].trim();
    const value = parts.slice(1).join(raw.delimiter).trim();
    if (!key) continue;

    if (isFirst) {
      isFirst = false;
      if (isHeaderRow(key, value)) continue;
    }

    if (!value) {
      skippedEmpty++;
      continue;
    }

    rows.push({ key, value });
  }

  return { rows, skippedEmpty };
}

function parseDateValue(value) {
  if (!value || !value.trim()) return null;
  const v = value.trim();

  function validDate(month, day, year) {
    return month >= 1 && month <= 12 && day >= 1 && day <= daysInMonth(month, year);
  }

  let m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const year = parseInt(m[1], 10), month = parseInt(m[2], 10), day = parseInt(m[3], 10);
    if (validDate(month, day, year)) return { month, day, year };
  }

  m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const a = parseInt(m[1], 10), b = parseInt(m[2], 10), year = parseInt(m[3], 10);
    if (a > 12 && b >= 1 && b <= 12 && validDate(b, a, year)) return { month: b, day: a, year };
    if (validDate(a, b, year)) return { month: a, day: b, year };
  }

  m = v.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) {
    const a = parseInt(m[1], 10), b = parseInt(m[2], 10), year = parseInt(m[3], 10);
    if (a > 12 && b >= 1 && b <= 12 && validDate(b, a, year)) return { month: b, day: a, year };
    if (validDate(a, b, year)) return { month: a, day: b, year };
  }

  m = v.match(/^([A-Za-z]{3,})\s+(\d{1,2}),?\s+(\d{4})$/);
  if (m) {
    const monthStr = m[1].toLowerCase();
    const day = parseInt(m[2], 10), year = parseInt(m[3], 10);
    const monthIdx = MONTH_NAMES.findIndex((n) => n.startsWith(monthStr));
    if (monthIdx !== -1 && validDate(monthIdx + 1, day, year)) return { month: monthIdx + 1, day, year };
  }

  return null;
}

function matchImportKeys(rows, fields) {
  const fieldLookup = Object.create(null);
  for (const f of fields) {
    fieldLookup[normalizeImportKey(f.key)] = f.key;
  }

  const matchMap = Object.create(null);
  const unmatched = [];

  for (const row of rows) {
    const normalized = normalizeImportKey(row.key);
    if (!normalized) {
      unmatched.push({ key: row.key, value: row.value });
      continue;
    }

    const fieldKey = fieldLookup[normalized];
    if (!fieldKey) {
      unmatched.push({ key: row.key, value: row.value });
      continue;
    }

    if (!matchMap[normalized]) {
      matchMap[normalized] = { fieldKey, value: row.value, originalKeys: [row.key] };
    } else {
      matchMap[normalized].value = row.value;
      matchMap[normalized].originalKeys.push(row.key);
    }
  }

  const matched = [];
  const duplicates = [];

  for (const normalized of Object.keys(matchMap)) {
    const entry = matchMap[normalized];
    matched.push({ fieldKey: entry.fieldKey, value: entry.value });
    if (entry.originalKeys.length > 1) {
      duplicates.push({ normalizedKey: entry.fieldKey, originalKeys: entry.originalKeys });
    }
  }

  return { matched, unmatched, duplicates };
}

// ── Horizontal Import Helpers ────────────────────────────────────────────────

function isHorizontalHeaderRow(cells) {
  if (!cells || cells.length < 3) return false;
  const nonEmpty = cells.filter((c) => c.trim() !== "");
  if (nonEmpty.length < 3) return false;
  return nonEmpty.every((c) => /[a-zA-Z]/.test(c));
}

function detectImportFormat(allRows) {
  const nonEmpty = allRows.filter((row) => row.some((c) => c.trim() !== ""));
  if (nonEmpty.length < 2) return "vertical";

  const headerRow = nonEmpty[0];
  if (headerRow.length < 3) return "vertical";
  if (!isHorizontalHeaderRow(headerRow)) return "vertical";

  const expectedCols = headerRow.length;
  const dataRows = nonEmpty.slice(1);
  for (const row of dataRows) {
    if (row.length !== expectedCols) return "vertical";
  }

  return "horizontal";
}

function extractHorizontalData(allRows) {
  const nonEmpty = allRows.filter((row) => row.some((c) => c.trim() !== ""));
  if (nonEmpty.length === 0) return { headers: [], dataRows: [] };
  return { headers: nonEmpty[0], dataRows: nonEmpty.slice(1) };
}

function horizontalRowToVertical(headers, dataRow) {
  const rows = [];
  let skippedEmpty = 0;
  for (let i = 0; i < headers.length; i++) {
    const key = (headers[i] || "").trim();
    const value = (dataRow[i] || "").trim();
    if (!key) continue;
    if (!value) { skippedEmpty++; continue; }
    rows.push({ key, value });
  }
  return { rows, skippedEmpty };
}

// ── Import UI ──────────────────────────────────────────────────────────────────

function toggleImportPanel() {
  if (hfScanInProgress || scanInProgress || fillInProgress) return;
  const panel = document.getElementById("import-panel");
  const btn = document.getElementById("fill-import-btn");
  if (!panel) return;

  if (panel.style.display === "none" || !panel.style.display) {
    panel.style.display = "block";
    btn?.classList.add("active");
    renderImportPanel();
  } else {
    closeImportPanel();
  }
}

function closeImportPanel() {
  if (importAutoCloseTimer) { clearTimeout(importAutoCloseTimer); importAutoCloseTimer = null; }
  importGeneration++;
  if (activeImportReader && activeImportReader.readyState === FileReader.LOADING) {
    activeImportReader.abort();
  }
  activeImportReader = null;
  pendingHorizontalHeaders = null;
  pendingHorizontalRows = null;
  pendingImportRawText = null;
  pendingImportSource = null;
  const panel = document.getElementById("import-panel");
  const btn = document.getElementById("fill-import-btn");
  if (panel) {
    panel.style.display = "none";
    panel.innerHTML = "";
  }
  btn?.classList.remove("active");
}

function renderRowPicker(headers, dataRows, page) {
  if (typeof page !== "number") page = 0;
  const totalPages = Math.ceil(dataRows.length / IMPORT_ROW_CAP);
  const startIdx = page * IMPORT_ROW_CAP;
  const pageRows = dataRows.slice(startIdx, startIdx + IMPORT_ROW_CAP);

  const panel = document.getElementById("import-panel");
  if (!panel) return;
  panel.innerHTML = "";

  // Close button
  const closeBtn = document.createElement("button");
  closeBtn.className = "import-close-btn";
  closeBtn.title = "Close";
  closeBtn.onclick = closeImportPanel;
  closeBtn.innerHTML = '<svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M2 2l8 8M10 2l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
  panel.appendChild(closeBtn);

  // Label
  const label = document.createElement("div");
  label.className = "import-row-picker-label";
  label.textContent = "Select a row to import";
  panel.appendChild(label);

  // Scrollable table container
  const scroll = document.createElement("div");
  scroll.className = "import-row-picker-scroll";

  const table = document.createElement("table");
  table.className = "import-row-picker-table";

  // Header row
  const thead = document.createElement("thead");
  const headerTr = document.createElement("tr");
  const radioTh = document.createElement("th");
  radioTh.className = "row-radio-col";
  headerTr.appendChild(radioTh);
  for (const h of headers) {
    const th = document.createElement("th");
    th.textContent = h;
    th.title = h;
    headerTr.appendChild(th);
  }
  thead.appendChild(headerTr);
  table.appendChild(thead);

  // Data rows for this page
  const tbody = document.createElement("tbody");

  for (let i = 0; i < pageRows.length; i++) {
    const actualIdx = startIdx + i;
    const tr = document.createElement("tr");
    if (i === 0) tr.className = "selected";

    // Radio cell
    const radioTd = document.createElement("td");
    radioTd.className = "row-radio-col";
    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "import-row";
    radio.value = String(actualIdx);
    radio.setAttribute("aria-label", `Row ${actualIdx + 1}: ${pageRows[i].slice(0, 3).map((c) => (c || "").trim()).filter(Boolean).join(", ")}`);
    if (i === 0) radio.checked = true;
    radio.onchange = function () {
      table.querySelectorAll("tbody tr").forEach((r) => r.classList.remove("selected"));
      tr.classList.add("selected");
    };
    radioTd.appendChild(radio);
    tr.appendChild(radioTd);

    // Data cells
    for (let j = 0; j < headers.length; j++) {
      const td = document.createElement("td");
      const cellValue = (pageRows[i][j] || "").trim();
      td.textContent = cellValue;
      td.title = cellValue;
      tr.appendChild(td);
    }

    // Clickable row
    tr.onclick = function () {
      table.querySelectorAll("tbody tr").forEach((r) => r.classList.remove("selected"));
      tr.classList.add("selected");
      radio.checked = true;
    };

    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  scroll.appendChild(table);
  panel.appendChild(scroll);

  // Pagination
  if (totalPages > 1) {
    const paginationWrap = document.createElement("div");
    paginationWrap.className = "import-pagination";

    const prevBtn = document.createElement("button");
    prevBtn.className = "import-page-btn";
    prevBtn.textContent = "Previous";
    prevBtn.disabled = page === 0;
    prevBtn.onclick = function () { renderRowPicker(headers, dataRows, page - 1); };

    const pageLabel = document.createElement("span");
    pageLabel.className = "import-page-label";
    pageLabel.textContent = `${startIdx + 1}\u2013${Math.min(startIdx + IMPORT_ROW_CAP, dataRows.length)} of ${dataRows.length}`;

    const nextBtn = document.createElement("button");
    nextBtn.className = "import-page-btn";
    nextBtn.textContent = "Next";
    nextBtn.disabled = page >= totalPages - 1;
    nextBtn.onclick = function () { renderRowPicker(headers, dataRows, page + 1); };

    paginationWrap.appendChild(prevBtn);
    paginationWrap.appendChild(pageLabel);
    paginationWrap.appendChild(nextBtn);
    panel.appendChild(paginationWrap);
  }

  // Import Row button
  const importBtn = document.createElement("button");
  importBtn.className = "import-select-btn";
  importBtn.textContent = "Import Row";
  importBtn.onclick = onHorizontalRowSelected;
  panel.appendChild(importBtn);

  // Back button
  const backBtn = document.createElement("button");
  backBtn.className = "import-back-btn";
  backBtn.textContent = "Back";
  backBtn.onclick = function () {
    pendingHorizontalHeaders = null;
    pendingHorizontalRows = null;
    pendingImportRawText = null;
    pendingImportSource = null;
    renderImportPanel();
  };
  panel.appendChild(backBtn);

  // Escape hatch
  const twoColLink = document.createElement("button");
  twoColLink.className = "import-twocol-link";
  twoColLink.textContent = "Import as Field Name and Value pairs instead";
  twoColLink.onclick = onUseAsTwoColumnFormat;
  panel.appendChild(twoColLink);

}

function onHorizontalRowSelected() {
  if (hfScanInProgress || scanInProgress || fillInProgress) {
    showImportSummary({ error: "Still finishing document setup. Please wait a moment." });
    return;
  }
  if (!pendingHorizontalHeaders || !pendingHorizontalRows) return;

  const selected = document.querySelector('input[name="import-row"]:checked');
  if (!selected) return;

  const idx = parseInt(selected.value, 10);
  const dataRow = pendingHorizontalRows[idx];
  if (!dataRow) return;

  const result = horizontalRowToVertical(pendingHorizontalHeaders, dataRow);
  pendingHorizontalHeaders = null;
  pendingHorizontalRows = null;
  pendingImportRawText = null;
  pendingImportSource = null;
  applyImportData(result, "horizontal");
}

function onUseAsTwoColumnFormat() {
  if (hfScanInProgress || scanInProgress || fillInProgress) {
    showImportSummary({ error: "Still finishing document setup. Please wait a moment." });
    return;
  }
  if (!pendingImportRawText || !pendingImportSource) return;

  const text = pendingImportRawText;
  const source = pendingImportSource;
  pendingHorizontalHeaders = null;
  pendingHorizontalRows = null;
  pendingImportRawText = null;
  pendingImportSource = null;

  const parseResult = source === "csv" ? parseCSV(text) : parsePastedText(text);
  applyImportData(parseResult);
}

function renderImportPanel() {
  if (importAutoCloseTimer) { clearTimeout(importAutoCloseTimer); importAutoCloseTimer = null; }
  const panel = document.getElementById("import-panel");
  if (!panel) return;

  panel.innerHTML = `
    <button class="import-close-btn" onclick="closeImportPanel()" title="Close">
      <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
        <path d="M2 2l8 8M10 2l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
      </svg>
    </button>
    <div class="import-section-label">Upload a spreadsheet</div>
    <button class="import-file-label" type="button" onclick="document.getElementById('import-file-input').click()">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 10V3M4 5.5l3-3 3 3" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/><path d="M2 10v2h10v-2" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/></svg>
      Choose .csv file
    </button>
    <input class="import-file-input" type="file" id="import-file-input" accept=".csv" onchange="onImportFileSelected(event)" />
    <div class="import-divider">or</div>
    <label class="import-section-label" for="import-paste-area">Paste your data</label>
    <textarea class="import-paste-area" id="import-paste-area" placeholder="Paste your data here..."></textarea>
    <div class="import-paste-hint">Two columns or a full spreadsheet</div>
    <button class="import-paste-btn" onclick="onPasteImport()">Import pasted data</button>
  `;
}

function onImportFileSelected(event) {
  // Cancel any in-flight file read from a previous selection
  importGeneration++;
  if (activeImportReader && activeImportReader.readyState === FileReader.LOADING) {
    activeImportReader.abort();
  }
  activeImportReader = null;

  if (hfScanInProgress || scanInProgress || fillInProgress) {
    showImportSummary({ error: "Still finishing document setup. Please wait a moment." });
    return;
  }
  const file = event.target.files?.[0];
  if (!file) return;

  if (file.size > IMPORT_MAX_FILE_SIZE) {
    showImportSummary({ error: "File is too large (max 5MB). Try a smaller file or paste your data instead." });
    return;
  }

  const ext = file.name.split(".").pop()?.toLowerCase();
  if (ext !== "csv") {
    showImportSummary({ error: "Please select a .csv file." });
    return;
  }

  const gen = importGeneration;
  const reader = new FileReader();
  activeImportReader = reader;
  reader.onload = function (e) {
    activeImportReader = null;
    if (gen !== importGeneration) return; // panel was closed or new import started
    if (hfScanInProgress || scanInProgress || fillInProgress) {
      showImportSummary({ error: "Document is being scanned. Please try again in a moment." });
      return;
    }
    const text = e.target?.result;
    if (typeof text !== "string") {
      showImportSummary({ error: "Could not read the file. Please try again." });
      return;
    }
    const allRows = parseCSVRaw(text);
    if (detectImportFormat(allRows) === "horizontal") {
      const hData = extractHorizontalData(allRows);
      pendingHorizontalHeaders = hData.headers;
      pendingHorizontalRows = hData.dataRows;
      pendingImportRawText = text;
      pendingImportSource = "csv";
      renderRowPicker(hData.headers, hData.dataRows);
    } else {
      const parseResult = parseCSV(text);
      applyImportData(parseResult);
    }
  };
  reader.onerror = function () {
    activeImportReader = null;
    if (gen !== importGeneration) return;
    showImportSummary({ error: "Could not read the file. Please try again." });
  };
  reader.readAsText(file);
}

function onPasteImport() {
  // Cancel any in-flight file read first (before scan guard, so stale reads never land)
  importGeneration++;
  if (activeImportReader && activeImportReader.readyState === FileReader.LOADING) {
    activeImportReader.abort();
  }
  activeImportReader = null;

  if (hfScanInProgress || scanInProgress || fillInProgress) {
    showImportSummary({ error: "Still finishing document setup. Please wait a moment." });
    return;
  }
  const textarea = document.getElementById("import-paste-area");
  const text = textarea ? textarea.value : "";
  if (!text.trim()) {
    showImportSummary({ error: "Paste some data first, then click Import." });
    return;
  }
  const raw = parsePastedRaw(text);
  if (detectImportFormat(raw.rows) === "horizontal") {
    const hData = extractHorizontalData(raw.rows);
    pendingHorizontalHeaders = hData.headers;
    pendingHorizontalRows = hData.dataRows;
    pendingImportRawText = text;
    pendingImportSource = "paste";
    renderRowPicker(hData.headers, hData.dataRows);
  } else {
    const parseResult = parsePastedText(text);
    applyImportData(parseResult);
  }
}

/**
 * Set a single form field's value programmatically.
 * Returns false if the field is a date type and parsing failed.
 */
function setFieldValue(fieldKey, value) {
  const field = currentFields.find((f) => f.key === fieldKey);
  if (!field) return false;

  // Clear any prior empty-field validation styling
  const fieldRow = document.querySelector(`.field-row[data-key="${fieldKey}"]`);
  if (fieldRow) fieldRow.classList.remove("field-empty");

  if (field.type === "date") {
    const parsed = parseDateValue(value);
    if (!parsed) return false;

    const container = document.getElementById(`val-${fieldKey}`);
    if (!container) return false;

    const monthSel = container.querySelector(".date-month");
    const daySel = container.querySelector(".date-day");
    const yearSel = container.querySelector(".date-year");
    if (!monthSel || !daySel || !yearSel) return false;

    // Add year option if outside dropdown range
    const yearStr = String(parsed.year);
    if (!yearSel.querySelector(`option[value="${yearStr}"]`)) {
      const opt = document.createElement("option");
      opt.value = yearStr;
      opt.textContent = yearStr;
      // Insert in sorted position
      const options = Array.from(yearSel.options);
      const insertBefore = options.find((o) => o.value && parseInt(o.value, 10) > parsed.year);
      if (insertBefore) {
        yearSel.insertBefore(opt, insertBefore);
      } else {
        yearSel.appendChild(opt);
      }
    }

    monthSel.value = String(parsed.month);
    yearSel.value = yearStr;
    updateDayOptions(fieldKey);
    daySel.value = String(parsed.day);
    return true;
  }

  // Text or paragraph
  const el = document.getElementById(`val-${fieldKey}`);
  if (!el) return false;
  el.value = value;
  return true;
}

function applyImportData(parseResult, sourceFormat) {
  const { rows, skippedEmpty } = parseResult;

  if (rows.length === 0) {
    const sf = sourceFormat || "vertical";
    const msg = skippedEmpty > 0
      ? (sf === "horizontal" ? "All cells were empty. No data to import." : "All rows had empty values. No data to import.")
      : "We couldn't read this data. Try a different format or check your file.";
    showImportSummary({ error: msg, skippedEmpty, sourceFormat: sf });
    return;
  }

  const result = matchImportKeys(rows, currentFields);

  if (result.matched.length === 0) {
    showImportSummary({
      filledCount: 0,
      unmatched: result.unmatched,
      skippedEmpty,
      sourceFormat: sourceFormat || "vertical",
      invalidDates: [],
      duplicates: result.duplicates,
    });
    return;
  }

  // Populate form fields and track invalid dates
  const invalidDates = [];
  let filledCount = 0;

  for (const { fieldKey, value } of result.matched) {
    const success = setFieldValue(fieldKey, value);
    if (success) {
      filledCount++;
    } else {
      invalidDates.push({ key: fieldKey, value });
    }
  }

  showImportSummary({
    filledCount,
    unmatched: result.unmatched,
    skippedEmpty,
    sourceFormat: sourceFormat || "vertical",
    invalidDates,
    duplicates: result.duplicates,
  });

  // Panel stays open until user dismisses via Done or X
}

function showImportSummary(data) {
  const panel = document.getElementById("import-panel");
  if (!panel) return;

  // No X button on summary screen -- Done button handles dismissal
  let html = '<div class="import-summary" role="status" aria-live="polite">';

  if (data.error) {
    html += `<div class="import-summary-bucket import-summary-error-bucket">${escapeHtml(data.error)}</div>`;
    if (data.skippedEmpty && data.skippedEmpty > 0) {
      const isHoriz = data.sourceFormat === "horizontal";
      const noun = isHoriz
        ? (data.skippedEmpty === 1 ? "empty cell" : "empty cells")
        : (data.skippedEmpty === 1 ? "row with empty value" : "rows with empty values");
      html += `<div class="import-summary-bucket import-summary-warning-bucket">${data.skippedEmpty} ${noun} skipped</div>`;
    }
  } else if (!data.filledCount || data.filledCount === 0) {
    // Nothing successfully filled -- single clean message instead of a wall of errors
    const hint = data.sourceFormat === "horizontal"
      ? 'If you pasted a simple list, try importing again and clicking "Import as Field Name and Value pairs instead".'
      : "Please check your spelling or formatting and try again.";
    html += `<div class="import-summary-bucket import-summary-error-bucket">None of this data matched your document fields. ${hint}</div>`;
  } else {
    // Filled bucket
    const noun = data.filledCount === 1 ? "field" : "fields";
    html += `<div class="import-summary-bucket import-summary-filled-bucket">${data.filledCount} ${noun} filled</div>`;

    // Unrecognized bucket (bulleted if multiple)
    if (data.unmatched && data.unmatched.length > 0) {
      const n = data.unmatched.length;
      if (n === 1) {
        html += `<div class="import-summary-bucket import-summary-error-bucket">1 not recognized: ${escapeHtml(data.unmatched[0].key)}</div>`;
      } else {
        html += `<div class="import-summary-bucket import-summary-error-bucket">${n} not recognized:<ul class="import-summary-list">`;
        const toShow = data.unmatched.slice(0, 8);
        for (const r of toShow) {
          html += `<li>${escapeHtml(r.key)}</li>`;
        }
        if (n > 8) html += `<li>and ${n - 8} more</li>`;
        html += '</ul></div>';
      }
    }

    // Invalid dates bucket (bulleted if multiple)
    if (data.invalidDates && data.invalidDates.length > 0) {
      const n = data.invalidDates.length;
      if (n === 1) {
        html += `<div class="import-summary-bucket import-summary-error-bucket">Could not parse date for ${escapeHtml(data.invalidDates[0].key)}: "${escapeHtml(data.invalidDates[0].value)}"</div>`;
      } else {
        html += `<div class="import-summary-bucket import-summary-error-bucket">Could not parse ${n} dates:<ul class="import-summary-list">`;
        for (const d of data.invalidDates) {
          html += `<li>${escapeHtml(d.key)}: "${escapeHtml(d.value)}"</li>`;
        }
        html += '</ul></div>';
      }
    }

    // Skipped empty bucket
    if (data.skippedEmpty && data.skippedEmpty > 0) {
      const isHoriz = data.sourceFormat === "horizontal";
      const noun = isHoriz
        ? (data.skippedEmpty === 1 ? "empty cell" : "empty cells")
        : (data.skippedEmpty === 1 ? "row with empty value" : "rows with empty values");
      html += `<div class="import-summary-bucket import-summary-warning-bucket">${data.skippedEmpty} ${noun} skipped</div>`;
    }

    // Duplicates bucket (bulleted if multiple)
    if (data.duplicates && data.duplicates.length > 0) {
      const n = data.duplicates.length;
      if (n === 1) {
        const dup = data.duplicates[0];
        const originals = dup.originalKeys.slice(0, 5).map((k) => escapeHtml(k)).join(", ");
        html += `<div class="import-summary-bucket import-summary-warning-bucket">${escapeHtml(dup.normalizedKey)} appeared ${dup.originalKeys.length} times (${originals}); used the last value</div>`;
      } else {
        html += `<div class="import-summary-bucket import-summary-warning-bucket">${n} fields had duplicate keys:<ul class="import-summary-list">`;
        const dupsToShow = data.duplicates.slice(0, 5);
        for (const dup of dupsToShow) {
          const originals = dup.originalKeys.slice(0, 5).map((k) => escapeHtml(k)).join(", ");
          html += `<li>${escapeHtml(dup.normalizedKey)}: ${originals}</li>`;
        }
        if (n > 5) html += `<li>and ${n - 5} more</li>`;
        html += '</ul></div>';
      }
    }
  }

  html += '</div>';
  html += '<button class="import-done-btn" onclick="closeImportPanel()">Done</button>';
  html += '<button class="import-retry-btn" onclick="renderImportPanel()">Import different data</button>';

  panel.innerHTML = html;
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

/** Scroll to the first empty/highlighted field and auto-focus its input. */
function scrollToFirstEmptyField() {
  const firstEmpty = document.querySelector(".field-row.field-empty");
  if (!firstEmpty) return;

  // Smooth scroll to center the field in view
  firstEmpty.scrollIntoView({ behavior: "smooth", block: "center" });

  // Auto-focus the input after scroll settles
  setTimeout(() => {
    const input = firstEmpty.querySelector(".field-value-input, .field-value-textarea, .date-month");
    if (input && !input.disabled) input.focus();
  }, 400);
}

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
  if (scanInProgress || hfScanInProgress) return;
  scanInProgress = true;
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

      // Check 1: CCs not in current field list (new fields added)
      for (const key of ccKeys) {
        if (!currentKeys.has(key)) { needsUpdate = true; break; }
      }

      // Check 1b: current fields no longer in CCs (fields deleted)
      if (!needsUpdate) {
        for (const key of currentKeys) {
          if (!ccKeys.has(key)) { needsUpdate = true; break; }
        }
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

      const allKeys = await sortKeysByDocumentOrder(context, ccMap);
      if (allKeys.length === 0) {
        // No DocFill fields remain -- clear state and show empty state
        currentFields = [];
        currentStorageKey = "";
        lastFilledValues = {};
        hasFilled = false;
        document.getElementById("fields-section").style.display = "none";
        document.getElementById("actions").style.display = "none";
        document.getElementById("empty-state").style.display = "block";
        document.querySelector(".empty-desc").textContent =
          'No {{placeholders}} found. Type fields like {{client_name}} into your document, then scan again.';
        return;
      }

      lastFilledValues = {};
      for (const [key, data] of Object.entries(ccMap)) {
        const text = data.text.trim();
        if (text && !isPlaceholderTextForKey(text, key)) lastFilledValues[key] = data.text.replace(/\v/g, "\n");
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
  } finally {
    scanInProgress = false;
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
  if (activeTab !== "create" || suppressSelectionPreview) return;
  clearTimeout(selectionDebounceTimer);
  selectionDebounceTimer = setTimeout(fetchCurrentSelection, 250);
}

async function fetchCurrentSelection() {
  if (activeTab !== "create" || selectionFetchInProgress || suppressSelectionPreview) return;
  selectionFetchInProgress = true;
  const myGeneration = selectionFetchGeneration;
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      // Abort if navigation happened during this async fetch
      if (myGeneration !== selectionFetchGeneration || suppressSelectionPreview) return;
      const text = sel.text.trim();
      lastSelectedText = text.includes("\r") || text.includes("\n") ? "" : text;
      updateSelectionPreview(lastSelectedText);
    });
  } catch { /* ignore */ }
  finally { selectionFetchInProgress = false; }
}

function updateSelectionPreview(text) {
  const preview = document.getElementById("selection-preview");
  const nameInput = document.getElementById("placeholder-name-input");
  const replaceBtn = document.getElementById("create-replace-btn");
  if (!preview || !nameInput) return;

  if (!text) {
    // State 1: Idle -- lock everything
    preview.className = "selection-preview selection-idle";
    preview.innerHTML = '<svg class="selection-idle-icon" width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M5 1h4M5 13h4M7 1v12" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg><span class="selection-idle-text">Highlight text in the document to begin</span>';
    nameInput.disabled = true;
    nameInput.value = "";
    if (replaceBtn) { replaceBtn.disabled = true; replaceBtn.classList.add("btn-disabled"); }
    return;
  }

  // State 2: Active -- unlock and populate
  const display = text.length > 60 ? text.substring(0, 60) + "\u2026" : text;
  preview.className = "selection-preview has-selection";
  preview.innerHTML = `<span class="selection-label">Selected</span><span class="selection-text">"${escapeHtml(display)}"</span>`;

  nameInput.disabled = false;
  if (replaceBtn) { replaceBtn.disabled = false; replaceBtn.classList.remove("btn-disabled"); }

  const suggested = suggestPlaceholderName(text);
  if (!nameInput.value || nameInput.value === lastSuggestedName) {
    nameInput.value = suggested;
    lastSuggestedName = suggested;
  }
  nameInput.focus();
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

  if (!text) return; // Button should be disabled, but guard anyway
  if (!name) { showCreateStatus("Enter a placeholder name.", "error"); nameInput.focus(); return; }
  if (!/^\w+$/.test(name)) { showCreateStatus("Use only letters, numbers, and underscores.", "error"); return; }

  const btn = document.getElementById("create-replace-btn");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Replacing...';
  hideCreateStatus();

  let shouldProceed = true;
  let exactCount = 0;
  let allCount = 0;

  // Only use matchWholeWord for simple word/phrase selections (no punctuation/symbols)
  const wholeWord = /^\w+(\s+\w+)*$/.test(text);

  try {
    await Word.run(async (context) => {
      // Exact-case search, filtering out ranges inside existing CCs
      const exactRaw = await searchAllBodies(context, text, { matchCase: true, matchWholeWord: wholeWord });
      const exactDeduped = await dedupeRanges(context, exactRaw);

      // Check parent CCs for exact matches -- only skip ranges inside DocFill CCs
      for (const r of exactDeduped) {
        r.parentContentControlOrNullObject.load("tag");
      }
      await context.sync();
      const exactItems = exactDeduped.filter((r) => {
        const parent = r.parentContentControlOrNullObject;
        return parent.isNullObject || !parent.tag || !parent.tag.startsWith(DOCFILL_TAG_PREFIX);
      });
      exactCount = exactItems.length;

      // Case-insensitive search for variants, also filtering
      const allRaw = await searchAllBodies(context, text, { matchCase: false, matchWholeWord: wholeWord });
      const allDeduped = await dedupeRanges(context, allRaw);
      for (const r of allDeduped) {
        r.parentContentControlOrNullObject.load("tag");
      }
      await context.sync();
      const allItems = allDeduped.filter((r) => {
        const parent = r.parentContentControlOrNullObject;
        return parent.isNullObject || !parent.tag || !parent.tag.startsWith(DOCFILL_TAG_PREFIX);
      });
      allCount = allItems.length;

      const variantCount = allCount - exactCount;

      if (allCount === 0) {
        shouldProceed = false;
        // Distinguish "inside CC" from "not found at all"
        if (allDeduped.length > 0) {
          showCreateStatus("This text is already inside a placeholder field.", "info");
        }
        return;
      }

      // Check for existing CCs for this key
      const existingCCs = context.document.contentControls.getByTag(keyToCCTag(name));
      existingCCs.load("items");
      await context.sync();
      const existingCount = existingCCs.items.length;

      if (exactCount === 1 && variantCount === 0 && existingCount === 0) {
        // Single exact match, no variants, no existing CCs: replace immediately
        const cc = exactItems[0].insertContentControl();
        cc.tag = keyToCCTag(name);
        cc.title = toTitleCase(name);
        cc.appearance = Word.ContentControlAppearance.boundingBox;
        cc.placeholderText = `{{${name}}}`;
        cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
        await context.sync();
      } else {
        // Multiple matches, variants, or existing CCs: show confirmation
        shouldProceed = false;
        pendingCreateText = text;
        pendingCreateName = name;
        pendingMatchCase = false; // case-insensitive to show ALL matches
        pendingMatchWholeWord = wholeWord;
        // Only enable match navigation for 2+ matches (nav bar not shown for single match)
        pendingMatchCount = allCount > 1 ? allCount : 0;

        // Find which match the user's selection is at/near (only for 2+ matches)
        if (pendingMatchCount > 0) {
          const sel = context.document.getSelection();
          const comparisons = allItems.map((r) => sel.compareLocationWith(r));
          await context.sync();
          let startIdx = 0;
          for (let i = 0; i < comparisons.length; i++) {
            const v = comparisons[i].value;
            if (v === "Equal" || v === "Inside" || v === "Contains" ||
                v === "ContainsStart" || v === "ContainsEnd" ||
                v === "InsideStart" || v === "InsideEnd" ||
                v === Word.LocationRelation.equal || v === Word.LocationRelation.inside ||
                v === Word.LocationRelation.contains ||
                v === Word.LocationRelation.containsStart || v === Word.LocationRelation.containsEnd ||
                v === Word.LocationRelation.insideStart || v === Word.LocationRelation.insideEnd) {
              startIdx = i;
              break;
            }
          }
          pendingMatchIndex = startIdx;
        }

        showReplaceAllConfirm(exactCount, allCount, name, existingCount);
      }
    });

    // Auto-navigate to the initial match after Word.run completes
    if (!shouldProceed && allCount > 1) {
      navigateMatch(0);
    }

    if (!shouldProceed && allCount === 0 && !document.querySelector("#create-status.info")) {
      showCreateStatus("Could not find that text in the document -- try selecting it again.", "error");
    } else if (shouldProceed) {
      onPlaceholderCreated(name, 1);
    }
  } catch (err) {
    showCreateStatus("Error: " + err.message, "error");
  }

  // Reset to idle only if we actually completed (not showing confirmation)
  if (shouldProceed || allCount === 0) {
    btn.innerHTML = "Convert to Placeholder";
    btn.disabled = true;
    btn.classList.add("btn-disabled");
    lastSelectedText = "";
    updateSelectionPreview("");
  }
}

function showReplaceAllConfirm(exactCount, allCount, name, existingCount) {
  const el = prepareCreateStatus();
  const variantCount = allCount - exactCount;
  const btnStyle = 'flex:1;padding:6px 0;background:#2563eb;color:#fff;border:none;border-radius:6px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;min-width:80px';
  const cancelStyle = 'padding:6px 10px;background:none;border:1.5px solid #bfdbfe;border-radius:6px;font-family:inherit;font-size:12px;color:#1d4ed8;cursor:pointer';
  const noteStyle = 'margin-top:6px;font-size:11px;color:#6b7280;line-height:1.3';

  let description;
  let buttons;
  const renameStyle = 'flex:1;padding:6px 0;background:none;border:1.5px solid #d1d5db;border-radius:6px;font-family:inherit;font-size:12px;font-weight:500;color:#374151;cursor:pointer;min-width:80px';
  const existingNote = existingCount > 0
    ? `<div style="${noteStyle}">Note: You already have a {{${escapeHtml(name)}}} field. Converting will link to the same field.</div>`
    : "";

  // When existing CCs exist, use "Link" language. Otherwise use "Convert/Replace" language.
  if (existingCount > 0) {
    // ── Existing field: "Link" wording ──
    // Preserve capitalization distinction in the description
    if (variantCount > 0 && exactCount > 0) {
      description = `Found <strong>${exactCount} exact</strong> and <strong>${variantCount}</strong> with different capitalization. Link to your existing <code>{{${escapeHtml(name)}}}</code> field?`;
    } else if (exactCount === 0) {
      description = `Found <strong>${allCount} match${allCount > 1 ? "es" : ""}</strong> with different capitalization. Link to your existing <code>{{${escapeHtml(name)}}}</code> field?`;
    } else {
      description = `Found <strong>${allCount} exact match${allCount > 1 ? "es" : ""}</strong>. Link ${allCount > 1 ? "them" : "it"} to your existing <code>{{${escapeHtml(name)}}}</code> field?`;
    }

    if (exactCount === 0) {
      // Only capitalization variants -- use 'all' mode (case-insensitive search)
      buttons = allCount === 1
        ? `<button onclick="confirmReplace('all')" style="${btnStyle}">Link this one</button>
           <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`
        : `<button onclick="confirmReplace('all')" style="${btnStyle}">Link all ${allCount}</button>
           <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`;
    } else if (allCount === 1) {
      // Single exact match
      buttons = `
        <button onclick="confirmReplace('single')" style="${btnStyle}">Link this one</button>
        <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`;
    } else if (variantCount === 0) {
      // Multiple exact matches only
      buttons = `
        <button onclick="confirmReplace('single')" style="${btnStyle}">Link this one</button>
        <button onclick="confirmReplace('exact')" style="${btnStyle}">Link all ${allCount}</button>
        <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`;
    } else {
      // Exact + variants: show "Link this one", optionally "Link N exact", "Link all N"
      buttons = `
        <button onclick="confirmReplace('single')" style="${btnStyle}">Link this one</button>
        ${exactCount > 1 ? `<button onclick="confirmReplace('exact')" style="${btnStyle}">Link ${exactCount} exact</button>` : ""}
        <button onclick="confirmReplace('all')" style="${btnStyle}">Link all ${allCount}</button>
        <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`;
    }

    const renameRow = `<div style="margin-top:6px"><button onclick="promptRenamePlaceholder()" style="${renameStyle};width:100%">Use different name</button></div>`;

    const navBar = pendingMatchCount > 1 ? `<div class="match-nav">
      <button class="match-nav-btn" id="match-nav-prev" onclick="navigateMatch(-1)" title="Previous match"><svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M8 2L4 6l4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg></button>
      <span class="match-nav-label" id="match-nav-label">${pendingMatchIndex + 1} of ${pendingMatchCount}</span>
      <button class="match-nav-btn" id="match-nav-next" onclick="navigateMatch(1)" title="Next match"><svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M4 2l4 4-4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg></button>
    </div>` : "";

    el.innerHTML = `${navBar}
      <div style="margin-bottom:8px">${description}${existingNote}</div>
      <div style="display:flex;gap:6px;flex-wrap:wrap">${buttons}</div>${renameRow}`;

  } else {
    // ── No existing field: "Convert/Replace" wording ──
    if (exactCount === 0) {
      description = `Found <strong>${allCount} match${allCount > 1 ? "es" : ""}</strong> with different capitalization. Replace with <code>{{${escapeHtml(name)}}}</code>?`;
      buttons = `
        <button onclick="confirmReplace('all')" style="${btnStyle}">${allCount === 1 ? "Convert" : `All ${allCount} matches`}</button>
        <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`;
    } else if (variantCount === 0) {
      description = `Found <strong>${exactCount} exact match${exactCount > 1 ? "es" : ""}</strong>. Replace with <code>{{${escapeHtml(name)}}}</code>?`;
      buttons = exactCount === 1
        ? `<button onclick="confirmReplace('single')" style="${btnStyle}">Convert</button>
           <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`
        : `<button onclick="confirmReplace('single')" style="${btnStyle}">This one only</button>
           <button onclick="confirmReplace('exact')" style="${btnStyle}">All ${exactCount} matches</button>
           <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`;
    } else {
      description = `Found <strong>${exactCount} exact</strong> and <strong>${variantCount}</strong> with different capitalization. Replace with <code>{{${escapeHtml(name)}}}</code>?`;
      buttons = `
        <button onclick="confirmReplace('single')" style="${btnStyle}">This one only</button>
        ${exactCount > 1 ? `<button onclick="confirmReplace('exact')" style="${btnStyle}">All ${exactCount} exact</button>` : ""}
        <button onclick="confirmReplace('all')" style="${btnStyle}">All ${allCount} matches</button>
        <button onclick="cancelCreateAction()" style="${cancelStyle}">Cancel</button>`;
    }

    const navBar = pendingMatchCount > 1 ? `<div class="match-nav">
      <button class="match-nav-btn" id="match-nav-prev" onclick="navigateMatch(-1)" title="Previous match"><svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M8 2L4 6l4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg></button>
      <span class="match-nav-label" id="match-nav-label">${pendingMatchIndex + 1} of ${pendingMatchCount}</span>
      <button class="match-nav-btn" id="match-nav-next" onclick="navigateMatch(1)" title="Next match"><svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M4 2l4 4-4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg></button>
    </div>` : "";

    el.innerHTML = `${navBar}
      <div style="margin-bottom:8px">${description}</div>
      <div style="display:flex;gap:6px;flex-wrap:wrap">${buttons}</div>`;
  }
  el.className = "info";
  el.style.display = "block";
}

/** User chose "Use different name" -- focus the name input and dismiss the dialog. */
function promptRenamePlaceholder() {
  hideCreateStatus();
  const nameInput = document.getElementById("placeholder-name-input");
  if (nameInput) {
    nameInput.value = "";
    nameInput.disabled = false;
    nameInput.focus();
    nameInput.placeholder = "Enter a unique name...";
  }
  // Keep pendingCreateText/Name so the user can submit with the new name
  // Reset shouldProceed state -- they'll click Convert again after renaming
  const btn = document.getElementById("create-replace-btn");
  if (btn) { btn.disabled = false; btn.classList.remove("btn-disabled"); btn.innerHTML = "Convert to Placeholder"; }
}

/** @param {'single'|'exact'|'all'} mode */
async function confirmReplace(mode) {
  hideCreateStatus();
  const text = pendingCreateText;
  const name = pendingCreateName;
  const savedMatchIndex = pendingMatchIndex;
  const hadMatchNav = pendingMatchCount > 0;

  // Clear match state before async work
  pendingCreateText = "";
  pendingCreateName = "";
  pendingMatchCount = 0;
  pendingMatchIndex = 0;
  pendingMatchCase = true;
  pendingMatchWholeWord = true;
  matchNavInFlight = false;
  chipNavGeneration++;
  suppressSelectionPreview = false;
  if (!text || !name) return;

  const btn = document.getElementById("create-replace-btn");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Replacing...';

  try {
    let count = 0;

    await Word.run(async (context) => {
      // When match navigation was active, use case-insensitive body-only search + index
      // (matches navigateMatch's search scope so indices align)
      const useMatchIndex = hadMatchNav && mode === "single";
      const matchCase = useMatchIndex ? false : (mode === "single" || mode === "exact");
      const wholeWord = /^\w+(\s+\w+)*$/.test(text);
      let items;
      if (useMatchIndex) {
        // Full search to match navigateMatch scope
        const rawItems = await searchAllBodies(context, text, { matchCase, matchWholeWord: wholeWord });
        items = await dedupeRanges(context, rawItems);
      } else {
        const rawItems = await searchAllBodies(context, text, { matchCase, matchWholeWord: wholeWord });
        items = await dedupeRanges(context, rawItems);
      }
      if (items.length === 0) return;

      // Check parent CCs to skip ranges already inside DocFill controls
      for (const range of items) {
        const parentCC = range.parentContentControlOrNullObject;
        parentCC.load("tag");
      }
      await context.sync();

      // Skip ranges inside DocFill content controls only
      const freeRanges = items.filter((range) => {
        const parent = range.parentContentControlOrNullObject;
        return parent.isNullObject || !parent.tag || !parent.tag.startsWith(DOCFILL_TAG_PREFIX);
      });

      if (freeRanges.length === 0) return;

      if (mode === "single") {
        let idx = -1;

        if (useMatchIndex) {
          // Use the navigated match index directly
          idx = savedMatchIndex < freeRanges.length ? savedMatchIndex : -1;
        } else {
          // Fallback: find by selection position (original behavior)
          const sel = context.document.getSelection();
          for (let i = 0; i < freeRanges.length; i++) {
            const loc = sel.compareLocationWith(freeRanges[i]);
            await context.sync();
            const v = loc.value;
            if (v === "Equal" || v === "Inside" || v === "Contains" ||
                v === Word.LocationRelation.equal ||
                v === Word.LocationRelation.inside ||
                v === Word.LocationRelation.contains) {
              idx = i;
              break;
            }
          }
        }

        if (idx === -1) {
          showCreateStatus("Selection changed. Please select the text again.", "error");
          return;
        }
        const cc = freeRanges[idx].insertContentControl();
        cc.tag = keyToCCTag(name);
        cc.title = toTitleCase(name);
        cc.appearance = Word.ContentControlAppearance.boundingBox;
        cc.placeholderText = `{{${name}}}`;
        cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
        count = 1;
      } else {
        for (const range of freeRanges) {
          const cc = range.insertContentControl();
          cc.tag = keyToCCTag(name);
          cc.title = toTitleCase(name);
          cc.appearance = Word.ContentControlAppearance.boundingBox;
          cc.placeholderText = `{{${name}}}`;
          cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
        }
        count = freeRanges.length;
      }
      await context.sync();
    });
    if (count > 0) {
      onPlaceholderCreated(name, count);
    }
  } catch (err) {
    showCreateStatus("Error: " + err.message, "error");
  }

  // Reset to State 1 (idle)
  btn.innerHTML = "Convert to Placeholder";
  btn.disabled = true;
  btn.classList.add("btn-disabled");
  lastSelectedText = "";
  updateSelectionPreview("");
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
    <div class="created-row" onclick="navigateToChip('${escapeAttr(e.name)}')">
      <span class="created-row-name">{{${escapeHtml(e.name)}}}</span>
      <span class="created-row-right">
        <span class="created-row-badge">${e.count}</span>
        <button class="created-row-remove" onclick="event.stopPropagation(); confirmDeletePlaceholder('${escapeAttr(e.name)}')">&times;</button>
      </span>
    </div>`).join("");
}

function filterCreatedList(query) {
  renderCreatedList(query);
}

/** Delete a placeholder: remove all its CCs from the document and refresh the list. */
function confirmDeletePlaceholder(name) {
  const el = prepareCreateStatus();
  el.innerHTML = `
    <div style="margin-bottom:6px;font-weight:600">Remove {{${escapeHtml(name)}}} from template?</div>
    <div style="margin-bottom:10px;font-size:12px;color:#64748b">This keeps the visible text but removes all DocFill controls for this field across the document.</div>
    <div style="display:flex;gap:8px">
      <button onclick="deleteCreatedPlaceholder('${escapeAttr(name)}')" style="flex:1;padding:7px 0;background:#dc2626;color:#fff;border:none;border-radius:7px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer">Remove</button>
      <button onclick="hideCreateStatus()" style="padding:7px 12px;background:none;border:1.5px solid #e5e7eb;border-radius:7px;font-family:inherit;font-size:12px;color:#374151;cursor:pointer">Cancel</button>
    </div>`;
  el.className = "info";
  el.style.display = "block";
}

async function deleteCreatedPlaceholder(name) {
  hideCreateStatus();
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(keyToCCTag(name));
      ccs.load("items,text");
      await context.sync();

      // Convert back to plain text: strip {{braces}}, keep the word
      for (const cc of ccs.items) {
        let plainText = cc.text || name;
        const m = plainText.match(/^\{\{(\w+)\}\}$/);
        if (m) plainText = m[1];
        cc.insertText(plainText, Word.InsertLocation.replace);
      }
      await context.sync();

      // Remove CC wrappers in a separate sync
      const ccs2 = context.document.contentControls.getByTag(keyToCCTag(name));
      ccs2.load("items");
      await context.sync();
      for (const cc of ccs2.items) {
        cc.delete(true);
      }
      if (ccs2.items.length > 0) await context.sync();
    });

    // Remove from Created list
    createdPlaceholders = createdPlaceholders.filter((e) => e.name !== name);
    renderCreatedList(document.getElementById("created-list-search")?.value);

    // Reconcile Fill state: remove from currentFields and lastFilledValues
    currentFields = currentFields.filter((f) => f.key !== name);
    delete lastFilledValues[name];
    if (Object.keys(lastFilledValues).length === 0) hasFilled = false;

    showCreateStatus(`{{${name}}} removed from template.`, "success");
  } catch (err) {
    showCreateStatus("Error removing placeholder: " + err.message, "error");
  }
}

let suppressionTimer = null;
let chipNavGeneration = 0;
let statusPreserveTimer = null;
let scrollLockRaf = null;
let scrollLockTimers = [];

function captureTaskPaneScroll() {
  const els = [
    document.scrollingElement,
    document.documentElement,
    document.body,
    document.querySelector("main"),
    document.querySelector(".created-list-scroll"),
  ].filter(Boolean);
  const unique = [...new Set(els)];
  return {
    windowX: window.scrollX,
    windowY: window.scrollY,
    elements: unique.map((el) => ({ el, top: el.scrollTop, left: el.scrollLeft })),
  };
}

function restoreTaskPaneScroll(snapshot) {
  window.scrollTo(snapshot.windowX, snapshot.windowY);
  for (const item of snapshot.elements) {
    item.el.scrollTop = item.top;
    item.el.scrollLeft = item.left;
  }
}

function startTaskPaneScrollLock(snapshot, generation, durationMs) {
  if (generation !== chipNavGeneration) return;
  if (scrollLockRaf) cancelAnimationFrame(scrollLockRaf);
  scrollLockTimers.forEach(clearTimeout);
  scrollLockTimers = [];
  const until = performance.now() + (durationMs || 1500);
  const tick = () => {
    if (generation !== chipNavGeneration) return;
    restoreTaskPaneScroll(snapshot);
    if (performance.now() < until) scrollLockRaf = requestAnimationFrame(tick);
  };
  tick();
  for (const delay of [0, 50, 150, 300, 600, 1000, 1500]) {
    scrollLockTimers.push(setTimeout(() => {
      if (generation === chipNavGeneration) restoreTaskPaneScroll(snapshot);
    }, delay));
  }
}

async function navigateToChip(name) {
  const idx = chipNavIndex[name] || 0;

  chipNavGeneration++;
  const myGeneration = chipNavGeneration;

  const scrollSnapshot = captureTaskPaneScroll();
  startTaskPaneScrollLock(scrollSnapshot, myGeneration);

  // Clear stale Create action state
  pendingCreateText = "";
  pendingCreateName = "";
  pendingMatchCount = 0;
  pendingMatchIndex = 0;
  pendingMatchCase = true;
  pendingMatchWholeWord = true;
  matchNavInFlight = false;
  lastSelectedText = "";
  lastSuggestedName = "";
  clearTimeout(selectionDebounceTimer);
  // Preserve layout while removing stale actionable content
  clearTimeout(statusPreserveTimer);
  const statusEl = document.getElementById("create-status");
  if (statusEl && statusEl.style.display !== "none" && statusEl.innerHTML) {
    const h = statusEl.offsetHeight;
    statusEl.innerHTML = "";
    statusEl.className = "";
    statusEl.style.height = h + "px";
    statusEl.style.visibility = "hidden";
    statusPreserveTimer = setTimeout(() => {
      // Only release if still empty/hidden (newer status hasn't taken over)
      if (!statusEl.innerHTML && statusEl.style.visibility === "hidden") {
        statusEl.style.height = "";
        statusEl.style.visibility = "";
        statusEl.style.display = "none";
      }
    }, 1600);
  }
  const nameInput = document.getElementById("placeholder-name-input");
  if (nameInput) { nameInput.disabled = true; nameInput.value = ""; }
  const replaceBtn = document.getElementById("create-replace-btn");
  if (replaceBtn) { replaceBtn.disabled = true; replaceBtn.classList.add("btn-disabled"); }

  suppressSelectionPreview = true;
  clearTimeout(suppressionTimer);
  selectionFetchGeneration++;

  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(keyToCCTag(name));
      ccs.load("items");
      await context.sync();

      if (ccs.items.length === 0) {
        createdPlaceholders = createdPlaceholders.filter((e) => e.name !== name);
        renderCreatedList(document.getElementById("created-list-search")?.value);
        showChipToast(`Placeholder not found. {{${name}}} removed.`);
        return;
      }

      const entry = createdPlaceholders.find((e) => e.name === name);
      if (entry && entry.count !== ccs.items.length) {
        entry.count = ccs.items.length;
        renderCreatedList(document.getElementById("created-list-search")?.value);
      }

      const targetIdx = idx % ccs.items.length;

      // Re-lock scroll right before and after cc.select()
      startTaskPaneScrollLock(scrollSnapshot, myGeneration);
      ccs.items[targetIdx].select();
      await context.sync();
      startTaskPaneScrollLock(scrollSnapshot, myGeneration);

      chipNavIndex[name] = (targetIdx + 1) % ccs.items.length;

      if (ccs.items.length > 1) {
        showChipToast(`{{${name}}} (${targetIdx + 1} of ${ccs.items.length})`);
      } else {
        showChipToast(`{{${name}}}`);
      }
    });
  } catch (err) {
    showChipToast("Error: " + err.message);
  } finally {
    suppressionTimer = setTimeout(() => {
      if (myGeneration === chipNavGeneration) suppressSelectionPreview = false;
    }, 500);
  }
}

/** Show a floating toast that doesn't shift layout. Auto-dismisses after 2s. */
let chipToastTimer = null;
// ── Match Navigation (Create Mode Confirmation) ─────────────────────────────

async function navigateMatch(delta) {
  if (matchNavInFlight) return;
  matchNavInFlight = true;

  pendingMatchIndex += delta;
  // Wrap around
  if (pendingMatchIndex < 0) pendingMatchIndex = pendingMatchCount - 1;
  if (pendingMatchIndex >= pendingMatchCount) pendingMatchIndex = 0;

  // Disable all dialog buttons except Cancel during navigation
  const dialog = document.getElementById("create-status");
  if (dialog) dialog.querySelectorAll("button:not([onclick*='cancel'])").forEach((b) => { b.disabled = true; });

  const scrollSnapshot = captureTaskPaneScroll();
  chipNavGeneration++;
  const myGeneration = chipNavGeneration;
  startTaskPaneScrollLock(scrollSnapshot, myGeneration);
  suppressSelectionPreview = true;
  clearTimeout(suppressionTimer);
  selectionFetchGeneration++;

  try {
    await Word.run(async (context) => {
      // Search all bodies (body + headers/footers) for consistent count
      const searchOpts = { matchCase: pendingMatchCase, matchWholeWord: pendingMatchWholeWord };
      const rawItems = await searchAllBodies(context, pendingCreateText, searchOpts);
      const deduped = await dedupeRanges(context, rawItems);

      // Batch parent-CC checks -- only skip ranges inside DocFill CCs
      for (const r of deduped) r.parentContentControlOrNullObject.load("tag");
      await context.sync();
      const freeRanges = deduped.filter((r) => {
        const parent = r.parentContentControlOrNullObject;
        return parent.isNullObject || !parent.tag || !parent.tag.startsWith(DOCFILL_TAG_PREFIX);
      });

      if (chipNavGeneration !== myGeneration) return;

      pendingMatchCount = freeRanges.length;
      if (pendingMatchIndex >= pendingMatchCount) {
        pendingMatchIndex = Math.max(0, pendingMatchCount - 1);
      }

      if (pendingMatchCount > 0 && pendingMatchIndex < freeRanges.length) {
        startTaskPaneScrollLock(scrollSnapshot, myGeneration);
        freeRanges[pendingMatchIndex].select();
        await context.sync();
        startTaskPaneScrollLock(scrollSnapshot, myGeneration);
      }
    });
  } catch { /* best effort */ }

  matchNavInFlight = false;

  if (chipNavGeneration !== myGeneration) return;

  updateMatchCounter();

  suppressionTimer = setTimeout(() => {
    if (chipNavGeneration === myGeneration) suppressSelectionPreview = false;
  }, 500);
}

function updateMatchCounter() {
  const label = document.getElementById("match-nav-label");
  const prevBtn = document.getElementById("match-nav-prev");
  const nextBtn = document.getElementById("match-nav-next");

  if (pendingMatchCount === 0) {
    if (label) label.textContent = "No matches found";
    if (prevBtn) prevBtn.disabled = true;
    if (nextBtn) nextBtn.disabled = true;
    const dialog = document.getElementById("create-status");
    if (dialog) dialog.querySelectorAll("button:not([onclick*='cancel'])").forEach((b) => { b.disabled = true; });
  } else {
    if (label) label.textContent = `${pendingMatchIndex + 1} of ${pendingMatchCount}`;
    // Re-enable all buttons (nav buttons always enabled for wrap-around)
    const dialog = document.getElementById("create-status");
    if (dialog) dialog.querySelectorAll("button").forEach((b) => { b.disabled = false; });
  }
}

function showChipToast(msg) {
  let toast = document.getElementById("chip-toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "chip-toast";
    toast.className = "chip-toast";
    document.getElementById("app").appendChild(toast);
  }
  toast.textContent = msg;
  toast.style.display = "block";
  toast.style.opacity = "1";
  clearTimeout(chipToastTimer);
  chipToastTimer = setTimeout(() => {
    toast.style.opacity = "0";
    setTimeout(() => { toast.style.display = "none"; }, 300);
  }, 2000);
}

function switchToFill() {
  switchTab("fill");
  // switchTab("fill") handles refresh via checkForNewPlaceholders or first-scan
}

// ── Create Status ──────────────────────────────────────────────────────────────


/** Prepare #create-status for new content: cancel any pending preserve timer. */
function prepareCreateStatus() {
  clearTimeout(statusPreserveTimer);
  const el = document.getElementById("create-status");
  if (el) { el.style.height = ""; el.style.visibility = ""; }
  return el;
}

function showCreateStatus(msg, type) {
  const el = prepareCreateStatus();
  el.textContent = msg;
  el.className = type;
  el.style.display = "block";
}

/** Cancel a pending Create action: hide status and reset button to proper state. */
function cancelCreateAction() {
  hideCreateStatus();
  pendingCreateText = "";
  pendingCreateName = "";
  pendingMatchCount = 0;
  pendingMatchIndex = 0;
  pendingMatchCase = true;
  pendingMatchWholeWord = true;
  matchNavInFlight = false;
  chipNavGeneration++;
  suppressSelectionPreview = false;
  const btn = document.getElementById("create-replace-btn");
  if (btn) {
    btn.innerHTML = "Convert to Placeholder";
    if (lastSelectedText) {
      btn.disabled = false;
      btn.classList.remove("btn-disabled");
    } else {
      btn.disabled = true;
      btn.classList.add("btn-disabled");
    }
  }
}

function hideCreateStatus() {
  clearTimeout(statusPreserveTimer);
  const el = document.getElementById("create-status");
  if (el) { el.style.display = "none"; el.style.height = ""; el.style.visibility = ""; }
}
