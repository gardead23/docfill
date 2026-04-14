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
};
