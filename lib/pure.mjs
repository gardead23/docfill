/**
 * Pure utility functions extracted from taskpane.js for testability.
 * These are duplicated here (not imported by taskpane.js) because the add-in
 * runs as a plain browser script with no module system.
 *
 * Keep these in sync with taskpane.js. Tests validate the logic here;
 * taskpane.js contains the identical implementations inline.
 */

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

function guessFieldType(key) {
  const k = key.toLowerCase();
  if (/date|day|month|year|when|start|end|deadline|due|expir|signed|effective/.test(k)) return "date";
  if (/description|notes?|bio|summary|detail|scope|address|comments?|message|body|terms/.test(k)) return "paragraph";
  return "text";
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

function daysInMonth(month, year) {
  if (!month) return 31;
  if (!year) year = new Date().getFullYear();
  return new Date(year, month, 0).getDate();
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

function buildStorageKey(keys, fingerprint) {
  const LS_PREFIX = "template-filler:";
  const base = [...keys].sort().join(",");
  return fingerprint
    ? LS_PREFIX + fingerprint + ":" + base
    : LS_PREFIX + base;
}

// ── DocFill Content Control Helpers ───────────────────────────────────────────

const DOCFILL_TAG_PREFIX = "docfill:";

/** Check whether a content control belongs to DocFill (by tag prefix). */
function isDocFillCC(cc) {
  return !!(cc && cc.tag && cc.tag.startsWith(DOCFILL_TAG_PREFIX));
}

/** Extract the placeholder key from a DocFill CC tag (always lowercase). */
function ccTagToKey(tag) {
  const raw = tag.startsWith(DOCFILL_TAG_PREFIX) ? tag.slice(DOCFILL_TAG_PREFIX.length) : tag;
  return raw.toLowerCase();
}

/** Build a DocFill CC tag from a placeholder key (always lowercase). */
function keyToCCTag(key) {
  return DOCFILL_TAG_PREFIX + key.toLowerCase();
}

/** Return the placeholder display text for a key. */
function placeholderText(key) {
  return `{{${key}}}`;
}

/** Check whether text looks like any placeholder pattern (case-insensitive). */
function isPlaceholderText(text) {
  return /^\{\{\w+\}\}$/.test(text.trim());
}

/** Check if text is a placeholder for a specific key (case-insensitive). */
function isPlaceholderTextForKey(text, key) {
  const m = text.trim().match(/^\{\{(\w+)\}\}$/);
  return m !== null && m[1].toLowerCase() === key.toLowerCase();
}

/** Check whether a CC's text matches its placeholder (i.e., it is unfilled). */
function isCCUnfilled(ccText, key) {
  return isPlaceholderTextForKey(ccText.trim(), key) || ccText.trim() === "";
}

// ── Import Helpers ───────────────────────────────────────────────────────────

const MONTH_NAMES = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];

/**
 * Normalize an imported key for matching against field keys.
 * Strips {{braces}}, replaces spaces/hyphens with underscores,
 * lowercases, and removes any remaining non-word characters.
 */
function normalizeImportKey(rawKey) {
  return rawKey
    .trim()
    .replace(/^\{\{|\}\}$/g, "")
    .trim()
    .replace(/[\s-]+/g, "_")
    .toLowerCase()
    .replace(/[^a-z0-9_]/g, "");
}

/**
 * Detect whether the first row is a header row.
 * Returns true only when BOTH columns look like header labels.
 * This prevents silently dropping rows like "name,Danny".
 */
function isHeaderRow(col1, col2) {
  const HEADER_WORDS = /^(key|field|name|placeholder|variable|value|header|column|label|data)$/i;
  return HEADER_WORDS.test(col1.trim()) && HEADER_WORDS.test(col2.trim());
}

/**
 * Detect whether pasted text uses tabs or commas as delimiter.
 * Only used for paste — CSV uses the RFC 4180 parser directly.
 */
function detectDelimiter(line) {
  const tabs = (line.match(/\t/g) || []).length;
  const commas = (line.match(/,/g) || []).length;
  return tabs >= commas && tabs > 0 ? "\t" : ",";
}

/**
 * Parse CSV text (RFC 4180) into key-value rows.
 * Handles quoted fields, commas inside quotes, escaped quotes (""),
 * Windows line endings, header row detection, and empty value skipping.
 * Returns { rows: [{key, value}], skippedEmpty: number }.
 */
function parseCSV(text) {
  if (!text || !text.trim()) return { rows: [], skippedEmpty: 0 };

  // RFC 4180 state-machine: parse into rows of cells
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
          i++; // skip escaped quote
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
  // Final cell/row
  currentRow.push(currentCell);
  if (currentRow.some((c) => c !== "")) {
    allRows.push(currentRow);
  }

  // Process rows into {key, value} pairs
  const rows = [];
  let skippedEmpty = 0;
  let isFirst = true;

  for (const cells of allRows) {
    if (cells.length < 2) continue; // skip single-column rows
    const key = cells[0].trim();
    const value = cells[1].trim();
    if (!key) continue; // skip rows with empty key

    // Header detection on first valid two-column row
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

/**
 * Parse pasted tab-separated or comma-separated text into key-value rows.
 * Auto-detects delimiter from the first non-empty line.
 * Returns { rows: [{key, value}], skippedEmpty: number }.
 */
function parsePastedText(text) {
  if (!text || !text.trim()) return { rows: [], skippedEmpty: 0 };

  const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const nonEmpty = lines.filter((l) => l.trim());
  if (nonEmpty.length === 0) return { rows: [], skippedEmpty: 0 };

  const delimiter = detectDelimiter(nonEmpty[0]);

  const rows = [];
  let skippedEmpty = 0;
  let isFirst = true;

  for (const line of lines) {
    if (!line.trim()) continue;
    const parts = line.split(delimiter);
    if (parts.length < 2) continue;
    const key = parts[0].trim();
    const value = parts.slice(1).join(delimiter).trim(); // rejoin in case of extra delimiters
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

/**
 * Parse a date string in various common formats.
 * Supported: ISO (2026-03-22), US slash (03/22/2026), day-first when first>12,
 * month names (March 22, 2026 / Mar 22, 2026), US dash (03-22-2026).
 * No new Date() fallback — returns null if no explicit format matches.
 */
function parseDateValue(value) {
  if (!value || !value.trim()) return null;
  const v = value.trim();

  // Helper: validate day against actual month length
  function validDate(month, day, year) {
    return month >= 1 && month <= 12 && day >= 1 && day <= daysInMonth(month, year);
  }

  // ISO: YYYY-MM-DD
  let m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const year = parseInt(m[1], 10), month = parseInt(m[2], 10), day = parseInt(m[3], 10);
    if (validDate(month, day, year)) return { month, day, year };
  }

  // Slash: MM/DD/YYYY or DD/MM/YYYY (day-first when first > 12)
  m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const a = parseInt(m[1], 10), b = parseInt(m[2], 10), year = parseInt(m[3], 10);
    if (a > 12 && b >= 1 && b <= 12 && validDate(b, a, year)) return { month: b, day: a, year };
    if (validDate(a, b, year)) return { month: a, day: b, year };
  }

  // Dash: MM-DD-YYYY or DD-MM-YYYY (same heuristic)
  m = v.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) {
    const a = parseInt(m[1], 10), b = parseInt(m[2], 10), year = parseInt(m[3], 10);
    if (a > 12 && b >= 1 && b <= 12 && validDate(b, a, year)) return { month: b, day: a, year };
    if (validDate(a, b, year)) return { month: a, day: b, year };
  }

  // Month name: "March 22, 2026" or "Mar 22, 2026" (comma optional, min 3 chars)
  m = v.match(/^([A-Za-z]{3,})\s+(\d{1,2}),?\s+(\d{4})$/);
  if (m) {
    const monthStr = m[1].toLowerCase();
    const day = parseInt(m[2], 10), year = parseInt(m[3], 10);
    const monthIdx = MONTH_NAMES.findIndex((n) => n.startsWith(monthStr));
    if (monthIdx !== -1 && validDate(monthIdx + 1, day, year)) return { month: monthIdx + 1, day, year };
  }

  return null;
}

/**
 * Match imported key-value rows to document field keys.
 * Normalizes both sides for case-insensitive, space/underscore/brace matching.
 * Deduplicates: last value wins, but duplicates are tracked with original keys.
 * Returns { matched: [{fieldKey, value}], unmatched: [{key, value}], duplicates: [{normalizedKey, originalKeys: [string]}] }.
 */
function matchImportKeys(rows, fields) {
  // Build lookup: normalizedKey -> fieldKey (null-prototype to avoid inherited keys)
  const fieldLookup = Object.create(null);
  for (const f of fields) {
    fieldLookup[normalizeImportKey(f.key)] = f.key;
  }

  // Track matches and originals per normalized key (null-prototype)
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
      matchMap[normalized].value = row.value; // last wins
      matchMap[normalized].originalKeys.push(row.key);
    }
  }

  // Build results
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

export {
  toTitleCase,
  escapeHtml,
  escapeAttr,
  guessFieldType,
  suggestPlaceholderName,
  daysInMonth,
  formatDate,
  buildStorageKey,
  DOCFILL_TAG_PREFIX,
  isDocFillCC,
  ccTagToKey,
  keyToCCTag,
  placeholderText,
  isPlaceholderText,
  isPlaceholderTextForKey,
  isCCUnfilled,
  normalizeImportKey,
  isHeaderRow,
  detectDelimiter,
  parseCSV,
  parsePastedText,
  parseDateValue,
  matchImportKeys,
};
