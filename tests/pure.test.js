import { describe, it, expect } from "vitest";
import {
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
} from "../lib/pure.mjs";

// ── toTitleCase ──────────────────────────────────────────────────────────────

describe("toTitleCase", () => {
  it("converts snake_case to Title Case", () => {
    expect(toTitleCase("client_name")).toBe("Client Name");
  });

  it("converts camelCase to Title Case", () => {
    expect(toTitleCase("clientName")).toBe("Client Name");
  });

  it("handles single word", () => {
    expect(toTitleCase("name")).toBe("Name");
  });

  it("handles multiple underscores", () => {
    expect(toTitleCase("first_middle_last_name")).toBe("First Middle Last Name");
  });
});

// ── escapeHtml ───────────────────────────────────────────────────────────────

describe("escapeHtml", () => {
  it("escapes HTML special characters", () => {
    expect(escapeHtml('<script>"alert"</script>')).toBe(
      "&lt;script&gt;&quot;alert&quot;&lt;/script&gt;"
    );
  });

  it("escapes ampersands", () => {
    expect(escapeHtml("A & B")).toBe("A &amp; B");
  });

  it("handles plain text unchanged", () => {
    expect(escapeHtml("hello world")).toBe("hello world");
  });
});

// ── escapeAttr ───────────────────────────────────────────────────────────────

describe("escapeAttr", () => {
  it("escapes double quotes", () => {
    expect(escapeAttr('say "hello"')).toBe("say &quot;hello&quot;");
  });

  it("escapes single quotes", () => {
    expect(escapeAttr("it's")).toBe("it&#39;s");
  });
});

// ── guessFieldType ───────────────────────────────────────────────────────────

describe("guessFieldType", () => {
  it("detects date fields", () => {
    expect(guessFieldType("start_date")).toBe("date");
    expect(guessFieldType("deadline")).toBe("date");
    expect(guessFieldType("effective")).toBe("date");
    expect(guessFieldType("signed")).toBe("date");
    expect(guessFieldType("expiration")).toBe("date");
  });

  it("detects paragraph fields", () => {
    expect(guessFieldType("description")).toBe("paragraph");
    expect(guessFieldType("project_scope")).toBe("paragraph");
    expect(guessFieldType("notes")).toBe("paragraph");
    expect(guessFieldType("comments")).toBe("paragraph");
    expect(guessFieldType("address")).toBe("paragraph");
    expect(guessFieldType("terms")).toBe("paragraph");
  });

  it("defaults to text", () => {
    expect(guessFieldType("client_name")).toBe("text");
    expect(guessFieldType("total_fee")).toBe("text");
    expect(guessFieldType("email")).toBe("text");
  });
});

// ── suggestPlaceholderName ───────────────────────────────────────────────────

describe("suggestPlaceholderName", () => {
  it("converts to lowercase snake_case", () => {
    expect(suggestPlaceholderName("John Smith")).toBe("john_smith");
  });

  it("strips special characters", () => {
    expect(suggestPlaceholderName("$1,000.00")).toBe("100000");
  });

  it("truncates to 40 chars", () => {
    const long = "a".repeat(50);
    expect(suggestPlaceholderName(long).length).toBe(40);
  });

  it("trims trailing underscores", () => {
    expect(suggestPlaceholderName("hello ")).toBe("hello");
  });
});

// ── daysInMonth ──────────────────────────────────────────────────────────────

describe("daysInMonth", () => {
  it("returns 31 for January", () => {
    expect(daysInMonth(1, 2026)).toBe(31);
  });

  it("returns 28 for February in a non-leap year", () => {
    expect(daysInMonth(2, 2026)).toBe(28);
  });

  it("returns 29 for February in a leap year", () => {
    expect(daysInMonth(2, 2024)).toBe(29);
  });

  it("returns 30 for April", () => {
    expect(daysInMonth(4, 2026)).toBe(30);
  });

  it("returns 31 when month is 0/falsy", () => {
    expect(daysInMonth(0, 2026)).toBe(31);
    expect(daysInMonth(null, 2026)).toBe(31);
  });

  it("handles December", () => {
    expect(daysInMonth(12, 2026)).toBe(31);
  });
});

// ── formatDate ───────────────────────────────────────────────────────────────

describe("formatDate", () => {
  it("formats ISO", () => {
    expect(formatDate("2026-03-22", "iso")).toBe("2026-03-22");
  });

  it("formats short-us", () => {
    expect(formatDate("2026-03-22", "short-us")).toBe("03/22/2026");
  });

  it("formats short-intl", () => {
    expect(formatDate("2026-03-22", "short-intl")).toBe("22/03/2026");
  });

  it("formats long", () => {
    const result = formatDate("2026-03-22", "long");
    expect(result).toContain("March");
    expect(result).toContain("22");
    expect(result).toContain("2026");
  });

  it("formats abbr", () => {
    const result = formatDate("2026-03-22", "abbr");
    expect(result).toContain("Mar");
    expect(result).toContain("22");
    expect(result).toContain("2026");
  });

  it("defaults to long format", () => {
    const result = formatDate("2026-03-22", "unknown");
    expect(result).toContain("March");
  });
});

// ── buildStorageKey ──────────────────────────────────────────────────────────

describe("buildStorageKey", () => {
  it("sorts keys alphabetically", () => {
    expect(buildStorageKey(["b", "a", "c"], "")).toBe("template-filler:a,b,c");
  });

  it("includes fingerprint when provided", () => {
    expect(buildStorageKey(["name"], "abc123")).toBe("template-filler:abc123:name");
  });

  it("omits fingerprint when empty", () => {
    expect(buildStorageKey(["name"], "")).toBe("template-filler:name");
  });

  it("produces same key regardless of input order", () => {
    const key1 = buildStorageKey(["z", "a", "m"], "fp");
    const key2 = buildStorageKey(["a", "m", "z"], "fp");
    expect(key1).toBe(key2);
  });
});

// ── DocFill CC Helpers ───────────────────────────────────────────────────────

describe("isDocFillCC", () => {
  it("returns true for docfill-prefixed tag", () => {
    expect(isDocFillCC({ tag: "docfill:client_name" })).toBe(true);
  });

  it("returns false for non-prefixed tag", () => {
    expect(isDocFillCC({ tag: "client_name" })).toBe(false);
  });

  it("returns false for empty tag", () => {
    expect(isDocFillCC({ tag: "" })).toBe(false);
  });

  it("returns false for null/undefined", () => {
    expect(isDocFillCC(null)).toBe(false);
    expect(isDocFillCC(undefined)).toBe(false);
    expect(isDocFillCC({ tag: null })).toBe(false);
  });

  it("returns false for other add-in tags", () => {
    expect(isDocFillCC({ tag: "other-addin:field" })).toBe(false);
  });
});

describe("ccTagToKey", () => {
  it("strips docfill prefix", () => {
    expect(ccTagToKey("docfill:client_name")).toBe("client_name");
  });

  it("returns tag as-is if no prefix", () => {
    expect(ccTagToKey("client_name")).toBe("client_name");
  });

  it("handles keys with underscores", () => {
    expect(ccTagToKey("docfill:first_middle_last")).toBe("first_middle_last");
  });

  it("normalizes to lowercase", () => {
    expect(ccTagToKey("docfill:Describe")).toBe("describe");
    expect(ccTagToKey("docfill:CLIENT_NAME")).toBe("client_name");
    expect(ccTagToKey("Client_Name")).toBe("client_name");
  });
});

describe("keyToCCTag", () => {
  it("adds docfill prefix", () => {
    expect(keyToCCTag("client_name")).toBe("docfill:client_name");
  });

  it("handles empty key", () => {
    expect(keyToCCTag("")).toBe("docfill:");
  });

  it("normalizes to lowercase", () => {
    expect(keyToCCTag("Describe")).toBe("docfill:describe");
    expect(keyToCCTag("CLIENT_NAME")).toBe("docfill:client_name");
  });
});

describe("placeholderText", () => {
  it("wraps key in double braces", () => {
    expect(placeholderText("name")).toBe("{{name}}");
  });
});

describe("isPlaceholderText", () => {
  it("returns true for lowercase placeholder", () => {
    expect(isPlaceholderText("{{client_name}}")).toBe(true);
  });

  it("returns true for mixed-case placeholder", () => {
    expect(isPlaceholderText("{{ClientName}}")).toBe(true);
    expect(isPlaceholderText("{{CLIENT_NAME}}")).toBe(true);
  });

  it("returns true with whitespace around", () => {
    expect(isPlaceholderText("  {{name}}  ")).toBe(true);
  });

  it("returns false for filled values", () => {
    expect(isPlaceholderText("Acme Corp")).toBe(false);
    expect(isPlaceholderText("")).toBe(false);
    expect(isPlaceholderText("{{not closed")).toBe(false);
  });
});

describe("isPlaceholderTextForKey", () => {
  it("matches same key case-insensitively", () => {
    expect(isPlaceholderTextForKey("{{client_name}}", "client_name")).toBe(true);
    expect(isPlaceholderTextForKey("{{ClientName}}", "clientname")).toBe(true);
    expect(isPlaceholderTextForKey("{{CLIENT_NAME}}", "client_name")).toBe(true);
  });

  it("returns false for different key", () => {
    expect(isPlaceholderTextForKey("{{other_key}}", "client_name")).toBe(false);
  });

  it("returns false for non-placeholder text", () => {
    expect(isPlaceholderTextForKey("Acme Corp", "client_name")).toBe(false);
  });
});

describe("isCCUnfilled", () => {
  it("returns true when text matches placeholder", () => {
    expect(isCCUnfilled("{{client_name}}", "client_name")).toBe(true);
  });

  it("returns true for mixed-case placeholder", () => {
    expect(isCCUnfilled("{{ClientName}}", "clientname")).toBe(true);
  });

  it("returns true when text is empty", () => {
    expect(isCCUnfilled("", "client_name")).toBe(true);
  });

  it("returns true when text is whitespace around placeholder", () => {
    expect(isCCUnfilled("  {{client_name}}  ", "client_name")).toBe(true);
  });

  it("returns false when text is a filled value", () => {
    expect(isCCUnfilled("Acme Corp", "client_name")).toBe(false);
  });

  it("returns false when text is a different key's placeholder", () => {
    expect(isCCUnfilled("{{other_key}}", "client_name")).toBe(false);
  });
});

// ── Import Helpers ──────────────────────────────────────────────────────────

describe("normalizeImportKey", () => {
  it("passes through simple snake_case", () => {
    expect(normalizeImportKey("client_name")).toBe("client_name");
  });

  it("converts spaces to underscores and lowercases", () => {
    expect(normalizeImportKey("Client Name")).toBe("client_name");
  });

  it("lowercases uppercase keys", () => {
    expect(normalizeImportKey("CLIENT_NAME")).toBe("client_name");
  });

  it("strips surrounding curly braces", () => {
    expect(normalizeImportKey("{{client_name}}")).toBe("client_name");
  });

  it("strips braces with inner spaces and mixed case", () => {
    expect(normalizeImportKey("{{ Client Name }}")).toBe("client_name");
  });

  it("converts hyphens to underscores", () => {
    expect(normalizeImportKey("client-name")).toBe("client_name");
  });

  it("trims whitespace", () => {
    expect(normalizeImportKey("  client_name  ")).toBe("client_name");
  });

  it("returns empty string for empty input", () => {
    expect(normalizeImportKey("")).toBe("");
  });

  it("strips special characters", () => {
    expect(normalizeImportKey("field!!name")).toBe("fieldname");
  });

  it("preserves numeric prefixes", () => {
    expect(normalizeImportKey("123_field")).toBe("123_field");
  });
});

describe("isHeaderRow", () => {
  it("detects key/value header pair", () => {
    expect(isHeaderRow("key", "value")).toBe(true);
  });

  it("detects Field/Value header pair (case-insensitive)", () => {
    expect(isHeaderRow("Field", "Value")).toBe(true);
  });

  it("detects placeholder/value header pair", () => {
    expect(isHeaderRow("placeholder", "value")).toBe(true);
  });

  it("detects name/value header pair", () => {
    expect(isHeaderRow("name", "value")).toBe(true);
  });

  it("rejects when col2 is not a header word", () => {
    expect(isHeaderRow("name", "Danny")).toBe(false);
  });

  it("rejects when col1 is not a header word", () => {
    expect(isHeaderRow("client_name", "value")).toBe(false);
  });

  it("rejects empty strings", () => {
    expect(isHeaderRow("", "")).toBe(false);
  });

  it("detects label/data header pair", () => {
    expect(isHeaderRow("label", "data")).toBe(true);
  });
});

describe("detectDelimiter", () => {
  it("detects tab delimiter", () => {
    expect(detectDelimiter("a\tb")).toBe("\t");
  });

  it("detects comma delimiter", () => {
    expect(detectDelimiter("a,b")).toBe(",");
  });

  it("prefers tab when more tabs than commas", () => {
    expect(detectDelimiter("a\tb\tc")).toBe("\t");
  });

  it("defaults to comma when no delimiters", () => {
    expect(detectDelimiter("no delimiter")).toBe(",");
  });

  it("prefers tab when equal counts", () => {
    expect(detectDelimiter("a\tb,c")).toBe("\t");
  });
});

describe("parseCSV", () => {
  it("parses basic two-column CSV", () => {
    const result = parseCSV("name,John\nemail,j@x.com");
    expect(result.rows).toEqual([
      { key: "name", value: "John" },
      { key: "email", value: "j@x.com" },
    ]);
    expect(result.skippedEmpty).toBe(0);
  });

  it("skips header row when both columns are header words", () => {
    const result = parseCSV("Key,Value\nname,John");
    expect(result.rows).toEqual([{ key: "name", value: "John" }]);
  });

  it("does NOT skip first row when col2 is not a header word", () => {
    const result = parseCSV("name,Danny\nemail,d@x.com");
    expect(result.rows).toEqual([
      { key: "name", value: "Danny" },
      { key: "email", value: "d@x.com" },
    ]);
  });

  it("handles quoted values with commas", () => {
    const result = parseCSV('name,"Smith, John"');
    expect(result.rows).toEqual([{ key: "name", value: "Smith, John" }]);
  });

  it("handles escaped quotes inside quoted fields", () => {
    const result = parseCSV('name,"He said ""hi"""');
    expect(result.rows).toEqual([{ key: "name", value: 'He said "hi"' }]);
  });

  it("skips blank lines", () => {
    const result = parseCSV("name,John\n\nemail,j@x.com\n");
    expect(result.rows).toEqual([
      { key: "name", value: "John" },
      { key: "email", value: "j@x.com" },
    ]);
  });

  it("skips rows with empty values and counts them", () => {
    const result = parseCSV("name,\nemail,j@x.com");
    expect(result.rows).toEqual([{ key: "email", value: "j@x.com" }]);
    expect(result.skippedEmpty).toBe(1);
  });

  it("handles Windows line endings", () => {
    const result = parseCSV("name,John\r\nemail,j@x.com\r\n");
    expect(result.rows).toEqual([
      { key: "name", value: "John" },
      { key: "email", value: "j@x.com" },
    ]);
  });

  it("returns empty result for empty input", () => {
    expect(parseCSV("")).toEqual({ rows: [], skippedEmpty: 0 });
    expect(parseCSV("   ")).toEqual({ rows: [], skippedEmpty: 0 });
  });

  it("skips single-column rows", () => {
    const result = parseCSV("name,John\njust_a_key\nemail,j@x.com");
    expect(result.rows).toEqual([
      { key: "name", value: "John" },
      { key: "email", value: "j@x.com" },
    ]);
  });

  it("trims whitespace from keys and values", () => {
    const result = parseCSV("  name  ,  John  ");
    expect(result.rows).toEqual([{ key: "name", value: "John" }]);
  });

  it("handles quoted value with newlines", () => {
    const result = parseCSV('name,"line1\nline2"');
    expect(result.rows).toEqual([{ key: "name", value: "line1\nline2" }]);
  });
});

describe("parsePastedText", () => {
  it("parses tab-separated text", () => {
    const result = parsePastedText("name\tJohn\nemail\tj@x.com");
    expect(result.rows).toEqual([
      { key: "name", value: "John" },
      { key: "email", value: "j@x.com" },
    ]);
    expect(result.skippedEmpty).toBe(0);
  });

  it("parses comma-separated text", () => {
    const result = parsePastedText("name,John\nemail,j@x.com");
    expect(result.rows).toEqual([
      { key: "name", value: "John" },
      { key: "email", value: "j@x.com" },
    ]);
  });

  it("skips header row when both columns are header words", () => {
    const result = parsePastedText("Field\tValue\nname\tJohn");
    expect(result.rows).toEqual([{ key: "name", value: "John" }]);
  });

  it("skips empty lines", () => {
    const result = parsePastedText("name\tJohn\n\nemail\tj@x.com");
    expect(result.rows).toEqual([
      { key: "name", value: "John" },
      { key: "email", value: "j@x.com" },
    ]);
  });

  it("skips rows with empty values and counts them", () => {
    const result = parsePastedText("name\t\nemail\tj@x.com");
    expect(result.rows).toEqual([{ key: "email", value: "j@x.com" }]);
    expect(result.skippedEmpty).toBe(1);
  });

  it("returns empty result for empty input", () => {
    expect(parsePastedText("")).toEqual({ rows: [], skippedEmpty: 0 });
    expect(parsePastedText("   ")).toEqual({ rows: [], skippedEmpty: 0 });
  });

  it("rejoins extra delimiters into value", () => {
    const result = parsePastedText("address\t123 Main St\tApt 4");
    expect(result.rows).toEqual([{ key: "address", value: "123 Main St\tApt 4" }]);
  });
});

describe("parseDateValue", () => {
  it("parses ISO format", () => {
    expect(parseDateValue("2026-03-22")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("parses US slash format", () => {
    expect(parseDateValue("03/22/2026")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("parses day-first slash when first number > 12", () => {
    expect(parseDateValue("22/03/2026")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("defaults to US interpretation when ambiguous", () => {
    expect(parseDateValue("03/04/2026")).toEqual({ month: 3, day: 4, year: 2026 });
  });

  it("parses long month name", () => {
    expect(parseDateValue("March 22, 2026")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("parses abbreviated month name", () => {
    expect(parseDateValue("Mar 22, 2026")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("parses month name case-insensitively", () => {
    expect(parseDateValue("march 22, 2026")).toEqual({ month: 3, day: 22, year: 2026 });
    expect(parseDateValue("MARCH 22, 2026")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("parses month name without comma", () => {
    expect(parseDateValue("March 22 2026")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("parses US dash format", () => {
    expect(parseDateValue("03-22-2026")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("parses day-first dash when first number > 12", () => {
    expect(parseDateValue("22-03-2026")).toEqual({ month: 3, day: 22, year: 2026 });
  });

  it("returns null for invalid date string", () => {
    expect(parseDateValue("not a date")).toBeNull();
  });

  it("returns null for empty string", () => {
    expect(parseDateValue("")).toBeNull();
  });

  it("returns null for whitespace only", () => {
    expect(parseDateValue("   ")).toBeNull();
  });

  it("returns null for partial date", () => {
    expect(parseDateValue("March 2026")).toBeNull();
  });

  it("rejects invalid month numbers", () => {
    expect(parseDateValue("13/01/2026")).toEqual({ month: 1, day: 13, year: 2026 }); // day-first heuristic
    expect(parseDateValue("2026-13-01")).toBeNull(); // month 13 invalid in ISO
  });

  it("rejects invalid day for month (Feb 31)", () => {
    expect(parseDateValue("2026-02-31")).toBeNull();
    expect(parseDateValue("02/31/2026")).toBeNull();
  });

  it("rejects April 31", () => {
    expect(parseDateValue("04/31/2026")).toBeNull();
  });

  it("accepts Feb 29 in a leap year", () => {
    expect(parseDateValue("02/29/2024")).toEqual({ month: 2, day: 29, year: 2024 });
  });

  it("rejects Feb 29 in a non-leap year", () => {
    expect(parseDateValue("02/29/2025")).toBeNull();
  });

  it("rejects ambiguous one-letter month prefixes", () => {
    expect(parseDateValue("M 1, 2026")).toBeNull();
    expect(parseDateValue("J 1, 2026")).toBeNull();
  });

  it("rejects two-letter ambiguous month prefixes", () => {
    expect(parseDateValue("Ma 1, 2026")).toBeNull();
    expect(parseDateValue("Ju 1, 2026")).toBeNull();
  });
});

describe("matchImportKeys", () => {
  const fields = [
    { key: "client_name" },
    { key: "email" },
    { key: "start_date" },
  ];

  it("matches exact keys", () => {
    const rows = [{ key: "client_name", value: "Acme" }];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([{ fieldKey: "client_name", value: "Acme" }]);
    expect(result.unmatched).toEqual([]);
    expect(result.duplicates).toEqual([]);
  });

  it("matches case-insensitively", () => {
    const rows = [{ key: "CLIENT_NAME", value: "Acme" }];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([{ fieldKey: "client_name", value: "Acme" }]);
  });

  it("matches space-to-underscore", () => {
    const rows = [{ key: "Client Name", value: "Acme" }];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([{ fieldKey: "client_name", value: "Acme" }]);
  });

  it("matches with brace stripping", () => {
    const rows = [{ key: "{{client_name}}", value: "Acme" }];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([{ fieldKey: "client_name", value: "Acme" }]);
  });

  it("reports unmatched keys", () => {
    const rows = [{ key: "unknown_field", value: "x" }];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([]);
    expect(result.unmatched).toEqual([{ key: "unknown_field", value: "x" }]);
  });

  it("handles mixed matched and unmatched", () => {
    const rows = [
      { key: "client_name", value: "Acme" },
      { key: "unknown", value: "x" },
      { key: "email", value: "a@b.com" },
    ];
    const result = matchImportKeys(rows, fields);
    expect(result.matched.length).toBe(2);
    expect(result.unmatched.length).toBe(1);
  });

  it("deduplicates with last-wins and tracks duplicates", () => {
    const rows = [
      { key: "client_name", value: "First" },
      { key: "client_name", value: "Second" },
    ];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([{ fieldKey: "client_name", value: "Second" }]);
    expect(result.duplicates.length).toBe(1);
    expect(result.duplicates[0].originalKeys).toEqual(["client_name", "client_name"]);
  });

  it("tracks original keys when different forms normalize to same key", () => {
    const rows = [
      { key: "Client Name", value: "First" },
      { key: "client-name", value: "Second" },
    ];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([{ fieldKey: "client_name", value: "Second" }]);
    expect(result.duplicates[0].originalKeys).toEqual(["Client Name", "client-name"]);
  });

  it("returns empty results for empty rows", () => {
    const result = matchImportKeys([], fields);
    expect(result).toEqual({ matched: [], unmatched: [], duplicates: [] });
  });

  it("marks all as unmatched when fields is empty", () => {
    const rows = [{ key: "name", value: "John" }];
    const result = matchImportKeys(rows, []);
    expect(result.matched).toEqual([]);
    expect(result.unmatched).toEqual([{ key: "name", value: "John" }]);
  });

  it("handles rows with empty normalized keys", () => {
    const rows = [{ key: "!!!", value: "x" }];
    const result = matchImportKeys(rows, fields);
    expect(result.unmatched).toEqual([{ key: "!!!", value: "x" }]);
  });

  it("does not match inherited object properties like constructor", () => {
    const rows = [{ key: "constructor", value: "x" }];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([]);
    expect(result.unmatched).toEqual([{ key: "constructor", value: "x" }]);
  });

  it("does not match __proto__", () => {
    const rows = [{ key: "__proto__", value: "x" }];
    const result = matchImportKeys(rows, fields);
    expect(result.matched).toEqual([]);
    expect(result.unmatched).toEqual([{ key: "__proto__", value: "x" }]);
  });

  it("matches constructor key when a field with that name exists", () => {
    const fieldsWithConstructor = [{ key: "constructor" }];
    const rows = [{ key: "constructor", value: "Acme" }];
    const result = matchImportKeys(rows, fieldsWithConstructor);
    expect(result.matched).toEqual([{ fieldKey: "constructor", value: "Acme" }]);
  });
});
