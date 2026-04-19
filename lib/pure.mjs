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
 * RFC 4180 CSV state-machine. Returns string[][] (2D array of all cells).
 * Handles quoted fields, escaped quotes (""), Windows line endings.
 * Filters out completely empty rows.
 */
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

  // Filter completely empty rows
  return allRows.filter((row) => row.some((c) => c !== ""));
}

/**
 * Split pasted text into a 2D array by auto-detected delimiter.
 * Returns { rows: string[][], delimiter: string }.
 * Filters out empty lines.
 */
function parsePastedRaw(text) {
  if (!text || !text.trim()) return { rows: [], delimiter: "," };

  const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const nonEmpty = lines.filter((l) => l.trim());
  if (nonEmpty.length === 0) return { rows: [], delimiter: "," };

  const delimiter = detectDelimiter(nonEmpty[0]);
  const rows = nonEmpty.map((line) => line.split(delimiter));

  return { rows, delimiter };
}

/**
 * Parse CSV text (RFC 4180) into vertical key-value rows.
 * Uses parseCSVRaw for the state machine, then extracts first two columns.
 * Returns { rows: [{key, value}], skippedEmpty: number }.
 */
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

/**
 * Parse pasted text into vertical key-value rows.
 * Uses parsePastedRaw for splitting, then extracts key + rejoins remaining columns as value.
 * Returns { rows: [{key, value}], skippedEmpty: number }.
 */
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

// ── Horizontal Import Helpers ────────────────────────────────────────────────

/**
 * Check if a row of cells looks like column headers (not data values).
 * Requires 3+ non-empty cells, each containing at least one letter.
 */
function isHorizontalHeaderRow(cells) {
  if (!cells || cells.length < 3) return false;
  const nonEmpty = cells.filter((c) => c.trim() !== "");
  if (nonEmpty.length < 3) return false;
  return nonEmpty.every((c) => /[a-zA-Z]/.test(c));
}

/**
 * Detect whether a 2D array represents horizontal (multi-column) or vertical (two-column) data.
 * Returns "horizontal" only if: 3+ columns, first non-empty row is header-like,
 * at least 1 data row after header, and all rows have the same column count.
 * 2-column data is always "vertical".
 */
function detectImportFormat(allRows) {
  // Filter leading empty rows
  const nonEmpty = allRows.filter((row) => row.some((c) => c.trim() !== ""));
  if (nonEmpty.length < 2) return "vertical"; // need header + at least 1 data row

  const headerRow = nonEmpty[0];
  if (headerRow.length < 3) return "vertical"; // 2 columns = always vertical
  if (!isHorizontalHeaderRow(headerRow)) return "vertical";

  // All rows (header + data) must have the same column count.
  // Real spreadsheet exports have uniform widths; vertical paste with
  // extra delimiters has inconsistent widths (e.g., 3 cols then 2 cols).
  const expectedCols = headerRow.length;
  const dataRows = nonEmpty.slice(1);
  for (const row of dataRows) {
    if (row.length !== expectedCols) return "vertical";
  }

  return "horizontal";
}

/**
 * Extract headers and data rows from a 2D array for horizontal import.
 * First non-empty row becomes headers, remaining non-empty rows become data.
 */
function extractHorizontalData(allRows) {
  const nonEmpty = allRows.filter((row) => row.some((c) => c.trim() !== ""));
  if (nonEmpty.length === 0) return { headers: [], dataRows: [] };
  return { headers: nonEmpty[0], dataRows: nonEmpty.slice(1) };
}

/**
 * Convert one horizontal data row into vertical {key, value} pairs.
 * Pairs headers[i] with dataRow[i] up to headers.length.
 * Skips empty keys. Counts empty values in skippedEmpty (not included in rows).
 */
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
  parseCSVRaw,
  parsePastedRaw,
  parseCSV,
  parsePastedText,
  parseDateValue,
  matchImportKeys,
  isHorizontalHeaderRow,
  detectImportFormat,
  extractHorizontalData,
  horizontalRowToVertical,
};
