"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");

const report = require("./generate_report");
const theme = require("./theme.default");

const UNKNOWN_GUID = "11111111-2222-3333-4444-555555555555";
const GRAPH_GUID = "00000003-0000-0000-c000-000000000000";

test("P0-1 redaction: unknown GUID is redacted, known app GUID resolves, opt-out preserves raw id", () => {
  // Default behavior: redact unknown directory object GUIDs.
  const redacted = report.sanitizeText(UNKNOWN_GUID);
  assert.equal(redacted, "Directory object");
  assert.ok(!redacted.includes(UNKNOWN_GUID), "raw GUID must not leak");

  // Known application GUID maps to a friendly label.
  assert.equal(report.sanitizeText(GRAPH_GUID), "Microsoft Graph");

  // Opt-out: raw ids should be preserved. Use try/finally so the module-level
  // flag is always reset even if an assertion throws.
  try {
    report.__setShowRawIds(true);
    assert.equal(report.humanizeIdentifier(UNKNOWN_GUID), UNKNOWN_GUID);
    assert.equal(report.sanitizeText(UNKNOWN_GUID), UNKNOWN_GUID);
  } finally {
    report.__setShowRawIds(false);
  }

  // Confirm the reset restored default redaction.
  assert.equal(report.sanitizeText(UNKNOWN_GUID), "Directory object");
});

test("P1-1 fit invariant: addTableWrapper clamps rowH so the table fits inside its box", () => {
  let recordedOptions = null;
  const slide = {
    addShape() {},
    addTable(rows, options) {
      recordedOptions = options;
    },
  };
  const ctx = {
    theme,
    pres: {
      shapes: {
        ROUNDED_RECTANGLE: "ROUNDED_RECTANGLE",
        RECTANGLE: "RECTANGLE",
      },
    },
  };

  const rows = [];
  for (let i = 0; i < 9; i += 1) {
    rows.push([`r${i}c0`, `r${i}c1`]);
  }
  const options = {
    x: 0.42,
    y: 1.5,
    w: 6,
    h: 2.8,
    rows,
    colW: [3, 3],
    rowH: 0.5,
  };

  report.addTableWrapper(slide, ctx, options);

  assert.ok(recordedOptions, "addTable should have been called");
  // Floating-point knife-edge: the clamp makes rowH * rowCount conceptually equal
  // to (h - 0.08). Use a tiny tolerance so a rounding tie does not flag a phantom
  // overflow in production code.
  assert.ok(
    recordedOptions.rowH * options.rows.length <= options.h - 0.08 + 1e-9,
    "clamped rowH must keep the table inside its reserved box"
  );
  // Sanity: the unclamped rowH (0.5) would have overflowed.
  assert.ok(0.5 * options.rows.length > options.h - 0.08, "unclamped rowH would overflow");
});

test("P1-2 score: inferAssessment uses provided score or derives from policy states", () => {
  const provided = report.inferAssessment({ assessment: { score: 72 } }, []);
  assert.equal(provided.score, 72);
  assert.equal(provided.derived, false);

  const allEnabled = report.inferAssessment({}, [{ state: "enabled" }, { state: "enabled" }]);
  assert.equal(allEnabled.score, 100);
  assert.equal(allEnabled.derived, true);

  const mixed = report.inferAssessment({}, [{ state: "enabled" }, { state: "report_only" }]);
  assert.equal(mixed.score, 75);
  assert.equal(mixed.derived, true);
});

test("measureTableHeight: more rows is taller, a long cell yields a larger uniform row height", () => {
  const headers = [{ label: "A" }, { label: "B" }];
  const colW = [1.5, 1.5];

  const small = report.measureTableHeight({ headers, rows: [["x", "y"]], colW });
  const large = report.measureTableHeight({
    headers,
    rows: [["x", "y"], ["x", "y"], ["x", "y"], ["x", "y"]],
    colW,
  });
  assert.ok(large.totalH > small.totalH, "more rows should yield a larger totalH");

  const shortRow = report.measureTableHeight({ headers, rows: [["a", "b"]], colW });
  const longRow = report.measureTableHeight({
    headers,
    rows: [["this is a very long cell value that will wrap across many lines indeed", "b"]],
    colW,
  });
  assert.ok(
    longRow.uniformRowH > shortRow.uniformRowH,
    "a long wrapping cell should produce a taller uniform row height"
  );
});

test("wrappedLines: empty string is one line, long text in narrow width wraps", () => {
  assert.equal(report.wrappedLines("", 2, 10), 1);
  assert.ok(
    report.wrappedLines("a".repeat(200), 1, 10) > 1,
    "a long string in a narrow column should wrap to more than one line"
  );
});

test("chunk: splits into pages and treats size 0 as a single whole-list page", () => {
  const pages = report.chunk([1, 2, 3, 4, 5, 6, 7], 3);
  assert.equal(pages.length, 3);
  assert.deepEqual(pages.map((p) => p.length), [3, 3, 1]);

  const single = report.chunk([1, 2, 3], 0);
  assert.equal(single.length, 1);
  assert.deepEqual(single[0], [1, 2, 3]);
});

test("collectActiveGrantControls: maps known controls, preserves unknowns, recognizes block", () => {
  const mfa = report.collectActiveGrantControls({
    grantControls: { controls: ["mfa"], authStrength: null },
  });
  assert.deepEqual(mfa, ["Multifactor authentication"]);

  const unknown = report.collectActiveGrantControls({
    grantControls: { controls: ["Custom control X"], authStrength: null },
  });
  // Unrecognized controls pass through verbatim (sanitized form), not dropped.
  assert.deepEqual(unknown, [report.sanitizeText("Custom control X", 36)]);
  assert.equal(unknown.length, 1);

  const block = report.collectActiveGrantControls({
    grantControls: { controls: ["block"], authStrength: null },
  });
  assert.deepEqual(block, ["Block access"]);
});

test("normalizeState: maps Graph state strings to canonical values", () => {
  assert.equal(report.normalizeState("enabled"), "enabled");
  assert.equal(report.normalizeState("enabledForReportingButNotEnforced"), "report_only");
  assert.equal(report.normalizeState("disabled"), "disabled");
  assert.equal(report.normalizeState("somethingUnrecognized"), "disabled");
});

test("normalizePolicies: block grant produces Block action, normal grant produces Grant", () => {
  const blocked = report.normalizePolicies([{ grantControls: { builtInControls: ["block"] } }]);
  assert.equal(blocked.length, 1);
  assert.equal(blocked[0].action, "Block");

  const granted = report.normalizePolicies([{ grantControls: { builtInControls: ["mfa"] } }]);
  assert.equal(granted.length, 1);
  assert.equal(granted[0].action, "Grant");
});

test("deepMerge: override wins on leaves, base-only keys preserved, nested objects merged", () => {
  const base = { a: 1, nested: { x: 1, y: 2, deep: { keep: true } } };
  const override = { a: 2, nested: { y: 99, z: 3, deep: { add: 1 } } };
  const merged = report.deepMerge(base, override);

  assert.equal(merged.a, 2); // override wins
  assert.equal(merged.nested.x, 1); // base-only leaf preserved
  assert.equal(merged.nested.y, 99); // override wins on nested leaf
  assert.equal(merged.nested.z, 3); // override-only key added
  assert.deepEqual(merged.nested.deep, { keep: true, add: 1 }); // deep merge
});
