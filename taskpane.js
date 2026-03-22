/* global Office, Word */

"use strict";

// ── State ─────────────────────────────────────────────────────────────────────

/** @type {{ key: string, label: string, type: string }[]} */
let currentFields = [];
let currentStorageKey = "";
let originalOoxml = null;
let hasFilled = false;
/** @type {Record<string, string>} */
let lastFilledValues = {};

// ── Create Mode State ──────────────────────────────────────────────────────────

let activeTab = "fill";
let lastSelectedText = "";
let lastSuggestedName = "";
/** @type {{ name: string, count: number }[]} */
let createdPlaceholders = [];
let pendingCreateText = "";
let pendingCreateName = "";
let selectionDebounceTimer = null;
let selectionFetchInProgress = false;

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
    await Word.run(async (context) => {
      const body = context.document.body;
      const ooxmlResult = body.getOoxml();
      body.load("text");
      await context.sync();

      const raw = body.text || "";
      const matches = raw.match(/\{\{(\w+)\}\}/g) || [];
      const docKeys = [...new Set(matches)].map((m) => m.replace(/\{\{|\}\}/g, ""));

      // Sync state with doc: if a key appears as a {{placeholder}} it wasn't filled
      // (or was undone via Ctrl+Z). Remove it from tracking so Phase 1 won't mis-restore.
      for (const key of docKeys) {
        delete lastFilledValues[key];
      }
      if (Object.keys(lastFilledValues).length === 0) hasFilled = false;

      // Freeze OOXML snapshot when nothing is filled (or everything was undone)
      if (Object.keys(lastFilledValues).length === 0) {
        originalOoxml = ooxmlResult.value;
      }

      // Merge with previously filled fields so they stay visible in the form
      const filledNotInDoc = Object.keys(lastFilledValues).filter((k) => !docKeys.includes(k));
      const allKeys = new Set([...docKeys, ...filledNotInDoc]);

      // Preserve original field order; append any brand-new keys at the end
      const orderedExisting = currentFields.map((f) => f.key).filter((k) => allKeys.has(k));
      const brandNewKeys = [...allKeys].filter((k) => !currentFields.some((f) => f.key === k));
      const keys = [...orderedExisting, ...brandNewKeys];

      if (keys.length === 0) {
        showStatus(
          "No {{placeholders}} found. Add fields like {{client_name}} to your document and rescan.",
          "error"
        );
        setScanButtonLoading(false);
        return;
      }

      currentStorageKey = buildStorageKey(keys);
      const saved = loadFieldConfigs(currentStorageKey);

      currentFields = keys.map((key) => ({
        key,
        label: saved[key]?.label || toTitleCase(key),
        type: saved[key]?.type || guessFieldType(key),
      }));

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
        <button
          class="field-edit-btn"
          title="Change field type"
          onclick="toggleTypePanel('${escapeAttr(field.key)}')"
        >
          <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
            <path d="M8.5 1.5a1.414 1.414 0 0 1 2 2L3.5 10.5 1 11l.5-2.5 7-7z"
              stroke="currentColor" stroke-width="1.2"
              stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </button>
      </div>
      <div class="field-type-panel" id="type-panel-${escapeAttr(field.key)}">
        ${["text", "date", "number", "paragraph"].map((t) => `
          <button
            class="type-pill ${field.type === t ? "active" : ""}"
            onclick="setFieldType('${escapeAttr(field.key)}', '${t}')"
          >${t === "paragraph" ? "Long text" : t.charAt(0).toUpperCase() + t.slice(1)}</button>
        `).join("")}
      </div>
      ${buildValueInput(field)}
    `;

    fieldsList.appendChild(row);

    // Restore filled value and show reset button if this field has been filled
    if (lastFilledValues[field.key]) {
      const input = row.querySelector(".field-value-input, .field-value-textarea");
      if (input) input.value = lastFilledValues[field.key];
      const resetBtn = document.getElementById(`reset-btn-${field.key}`);
      if (resetBtn) resetBtn.style.display = "inline-flex";
    }
  });
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
  const inputType = field.type === "date" ? "date" : field.type === "number" ? "number" : "text";
  return `<input
    id="${id}"
    class="field-value-input"
    type="${inputType}"
    placeholder="${inputType === "date" ? "" : "Enter " + escapeHtml(field.label).toLowerCase() + "..."}"
  />`;
}

// ── Field Edit Handlers ────────────────────────────────────────────────────────

function toggleTypePanel(key) {
  const panel = document.getElementById(`type-panel-${key}`);
  if (!panel) return;
  const isOpen = panel.classList.contains("open");
  // Close all panels first
  document.querySelectorAll(".field-type-panel.open").forEach((p) => p.classList.remove("open"));
  if (!isOpen) panel.classList.add("open");
}

function setFieldType(key, newType) {
  const field = currentFields.find((f) => f.key === key);
  if (!field || field.type === newType) {
    toggleTypePanel(key); // close panel if same type selected
    return;
  }

  const oldValue = document.getElementById(`val-${key}`)?.value || "";
  field.type = newType;
  saveFieldConfigs(currentStorageKey, currentFields);

  // Rebuild the value input
  const row = document.querySelector(`.field-row[data-key="${key}"]`);
  if (!row) return;
  const oldInput = row.querySelector(".field-value-input, .field-value-textarea");
  if (oldInput) oldInput.remove();
  row.insertAdjacentHTML("beforeend", buildValueInput(field));
  if (newType !== "date") {
    const newInput = row.querySelector(".field-value-input, .field-value-textarea");
    if (newInput) newInput.value = oldValue;
  }

  // Update pill active states
  row.querySelectorAll(".type-pill").forEach((pill) => {
    pill.classList.toggle("active", pill.textContent.trim().toLowerCase().replace(" ", "") ===
      (newType === "paragraph" ? "longtext" : newType));
  });

  // Close the panel
  const panel = document.getElementById(`type-panel-${key}`);
  if (panel) panel.classList.remove("open");
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
  if (/amount|fee|price|cost|total|number|qty|quantity|count|rate|salary|budget|hours|days/.test(k)) return "number";
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

  // Detect if two re-filled fields share the same value (Phase 1 may confuse them)
  const refillVals = Object.entries(toFill).filter(([k]) => lastFilledValues[k]).map(([, v]) => v);
  const hasDuplicateRefillValues = refillVals.some((v, i) => refillVals.indexOf(v) !== i);

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Filling...';
  hideStatus();

  try {
    let totalReplaced = 0;

    await Word.run(async (context) => {
      // Phase 1: restore any previously filled values back to {{placeholders}}
      const keysToRestore = Object.keys(toFill).filter((k) => lastFilledValues[k]);
      if (keysToRestore.length > 0) {
        const restoreSearches = {};
        for (const key of keysToRestore) {
          restoreSearches[key] = context.document.body.search(lastFilledValues[key], { matchCase: true });
          restoreSearches[key].load("items");
        }
        await context.sync();
        for (const [key, results] of Object.entries(restoreSearches)) {
          results.items.forEach((item) => item.insertText(`{{${key}}}`, Word.InsertLocation.replace));
        }
        await context.sync();
      }

      // Phase 2: fill placeholders with new values
      for (const [key, value] of Object.entries(toFill)) {
        const results = context.document.body.search(`{{${key}}}`, { matchCase: true });
        results.load("items");
        await context.sync();
        totalReplaced += results.items.length;
        results.items.forEach((item) => item.insertText(value, Word.InsertLocation.replace));
      }
      await context.sync();
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
      } else if (hasDuplicateRefillValues) {
        showStatus("✓ Filled. Note: some fields shared the same value — if anything looks mixed up, clear all fields and re-fill.", "info");
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
  currentFields.forEach((field) => {
    const el = document.getElementById(`val-${field.key}`);
    let value = el ? el.value.trim() : "";
    if (field.type === "date" && value) value = formatDate(value);
    values[field.key] = value;
  });
  return values;
}

function formatDate(isoDate) {
  try {
    const [year, month, day] = isoDate.split("-").map(Number);
    return new Date(year, month - 1, day).toLocaleDateString("en-US", {
      year: "numeric", month: "long", day: "numeric",
    });
  } catch {
    return isoDate;
  }
}

// ── Clear Form ─────────────────────────────────────────────────────────────────

async function clearForm() {
  if (hasFilled && originalOoxml) {
    showClearConfirm();
    return;
  }
  doFormClear();
}

function showClearConfirm() {
  const el = document.getElementById("status");
  el.innerHTML = `
    <div style="margin-bottom:10px">Reset the document to its original template?</div>
    <div style="display:flex;gap:8px">
      <button onclick="confirmReset()" style="flex:1;padding:7px 0;background:#2563eb;color:#fff;border:none;border-radius:7px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer">Reset Document</button>
      <button onclick="hideStatus()" style="padding:7px 12px;background:none;border:1.5px solid #bfdbfe;border-radius:7px;font-family:inherit;font-size:12px;color:#1d4ed8;cursor:pointer">Cancel</button>
    </div>
  `;
  el.className = "info";
  el.style.display = "block";
}

async function confirmReset() {
  hideStatus();
  const clearBtn = document.querySelector(".btn-clear");
  if (clearBtn) { clearBtn.disabled = true; clearBtn.textContent = "Resetting..."; }

  try {
    await Word.run(async (context) => {
      context.document.body.insertOoxml(originalOoxml, Word.InsertLocation.replace);
      await context.sync();
    });
    hasFilled = false;
  } catch (err) {
    showStatus("Failed to reset document: " + err.message, "error");
    if (clearBtn) { clearBtn.disabled = false; clearBtn.textContent = "Clear all fields"; }
    return;
  }

  if (clearBtn) { clearBtn.disabled = false; clearBtn.textContent = "Clear all fields"; }
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
      const results = context.document.body.search(filledValue, { matchCase: true });
      results.load("items");
      await context.sync();
      if (results.items.length > 0) {
        found = true;
        results.items.forEach((item) => item.insertText(`{{${key}}}`, Word.InsertLocation.replace));
        await context.sync();
      }
    });

    if (found) {
      delete lastFilledValues[key];
      const input = document.getElementById(`val-${key}`);
      if (input) input.value = "";
      if (resetBtn) { resetBtn.style.display = "none"; resetBtn.disabled = false; }
      if (Object.keys(lastFilledValues).length === 0) hasFilled = false;
    } else {
      if (resetBtn) resetBtn.disabled = false;
      showStatus("Could not find this field's value in the document — it may have been edited directly.", "error");
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
  document.querySelectorAll(".field-row.field-empty").forEach((r) => r.classList.remove("field-empty"));
  document.querySelectorAll(".field-reset-btn").forEach((btn) => { btn.style.display = "none"; });
  lastFilledValues = {};
  hideStatus();
}

// ── localStorage ───────────────────────────────────────────────────────────────

const LS_PREFIX = "template-filler:";

function buildStorageKey(keys) {
  return LS_PREFIX + [...keys].sort().join(",");
}

function loadFieldConfigs(storageKey) {
  try { return JSON.parse(localStorage.getItem(storageKey) || "{}"); }
  catch { return {}; }
}

function saveFieldConfigs(storageKey, fields) {
  const data = {};
  fields.forEach((f) => { data[f.key] = { label: f.label, type: f.type }; });
  try { localStorage.setItem(storageKey, JSON.stringify(data)); } catch { /* ignore */ }
}

// ── UI Helpers ─────────────────────────────────────────────────────────────────

function showStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = type;
  el.style.display = "block";
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
      const results = context.document.body.search(text, { matchCase: true });
      results.load("items");
      await context.sync();

      occurrenceCount = results.items.length;

      if (occurrenceCount === 0) {
        shouldProceed = false;
        return;
      }

      if (occurrenceCount > 1) {
        shouldProceed = false;
        pendingCreateText = text;
        pendingCreateName = name;
        showReplaceAllConfirm(occurrenceCount, name);
        return;
      }

      results.items.forEach((item) => item.insertText(`{{${name}}}`, Word.InsertLocation.replace));
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

function showReplaceAllConfirm(count, name) {
  const el = document.getElementById("create-status");
  el.innerHTML = `
    <div style="margin-bottom:8px">Found <strong>${count} occurrences</strong> of this text. Replace with <code>{{${escapeHtml(name)}}}</code>?</div>
    <div style="display:flex;gap:6px;flex-wrap:wrap">
      <button onclick="confirmReplace(false)" style="flex:1;padding:6px 0;background:#2563eb;color:#fff;border:none;border-radius:6px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;min-width:80px">This word only</button>
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
  pendingCreateText = "";
  pendingCreateName = "";
  if (!text || !name) return;

  const btn = document.getElementById("create-replace-btn");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Replacing...';

  try {
    let count = 0;
    await Word.run(async (context) => {
      const results = context.document.body.search(text, { matchCase: true });
      results.load("items");
      await context.sync();
      if (results.items.length === 0) return;
      if (replaceAll) {
        count = results.items.length;
        results.items.forEach((item) => item.insertText(`{{${name}}}`, Word.InsertLocation.replace));
      } else {
        count = 1;
        results.items[0].insertText(`{{${name}}}`, Word.InsertLocation.replace);
      }
      await context.sync();
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
    .map((e) => `<span class="created-chip">{{${escapeHtml(e.name)}}}${e.count > 1 ? `<span class="chip-count">×${e.count}</span>` : ""}</span>`)
    .join("");
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
