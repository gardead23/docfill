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
} from "../lib/pure.js";

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
