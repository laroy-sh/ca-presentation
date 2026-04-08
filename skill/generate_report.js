const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ─── DESIGN TOKENS (dark theme) ────────────────────────────────
const C = {
  bg:         "0F172A",
  bgCard:     "1E293B",
  bgCardDim:  "162032",
  border:     "334155",
  borderTeal: "0E7490",
  teal:       "06B6D4",
  tealDark:   "0891B2",
  white:      "F1F5F9",
  muted:      "94A3B8",
  mutedDark:  "64748B",
  green:      "10B981",
  greenDim:   "064E3B",
  amber:      "F59E0B",
  amberDim:   "78350F",
  red:        "EF4444",
  redDim:     "7F1D1D",
  grantGreen: "10B981",
  blockRed:   "EF4444",
  circle:     "1E3A5F",
};

const FONT = "Calibri";
const FONT_LIGHT = "Calibri Light";
const W = 10;
const H = 5.625;
const M = 0.4;

// ─── COLOR KEY RESOLVER ────────────────────────────────────────
function resolveColor(colorKey) {
  return C[colorKey] || C.muted;
}

function resolveRatingDim(ratingColor) {
  if (ratingColor === "red") return C.redDim;
  if (ratingColor === "teal") return "0E3B4E";
  return C.greenDim;
}

// ─── ALL POSSIBLE CONTROLS ────────────────────────────────────
const ALL_GRANT_CONTROLS = [
  "Multifactor authentication",
  "Authentication strength",
  "Compliant device",
  "Hybrid Azure AD joined",
  "Approved client app",
  "App protection policy",
  "Change password",
  "Terms of use",
];

const ALL_SESSION_CONTROLS = [
  "App enforced restrictions",
  "Conditional Access App Control",
  "Sign-in frequency",
  "Persistent browser session",
  "Continuous access evaluation",
  "Disable resilience defaults",
  "Token protection",
];

// ─── HELPERS ───────────────────────────────────────────────────
function stateColor(state) {
  if (state === "enabled") return { accent: C.green, dim: C.greenDim, label: "ENABLED" };
  if (state === "report_only") return { accent: C.amber, dim: C.amberDim, label: "REPORT-ONLY" };
  return { accent: C.red, dim: C.redDim, label: "DISABLED" };
}

function getAction(p) {
  const g = p.grantControls || {};
  if (g.controls && g.controls.some(c => c.toLowerCase().includes("block"))) return "Block";
  if (g.operator === "Not configured") return "\u2014";
  return "Grant";
}

function getActiveGrantControls(policy) {
  const g = policy.grantControls || {};
  const active = new Set();
  (g.controls || []).forEach(c => {
    const cl = c.toLowerCase();
    if (cl.includes("multifactor")) active.add("Multifactor authentication");
    if (cl.includes("compliant")) active.add("Compliant device");
    if (cl.includes("hybrid") || cl.includes("azure ad joined")) active.add("Hybrid Azure AD joined");
    if (cl.includes("approved client")) active.add("Approved client app");
    if (cl.includes("app protection")) active.add("App protection policy");
    if (cl.includes("password change") || cl.includes("change password") || (cl.includes("password") && cl.includes("require"))) active.add("Change password");
    if (cl.includes("terms")) active.add("Terms of use");
  });
  if (g.authStrength) active.add("Authentication strength");
  return active;
}

function isBlockAccess(policy) {
  const g = policy.grantControls || {};
  return (g.controls || []).some(c => c.toLowerCase().includes("block"));
}

// ─── DECORATIVE BACKGROUND CIRCLES ─────────────────────────────
function addBgCircles(slide, pres) {
  slide.addShape(pres.shapes.OVAL, {
    x: 6.5, y: -1.5, w: 5, h: 5,
    fill: { color: C.circle, transparency: 70 }
  });
  slide.addShape(pres.shapes.OVAL, {
    x: 7.5, y: 2.0, w: 4, h: 4,
    fill: { color: C.circle, transparency: 80 }
  });
}

// ═══════════════════════════════════════════════════════════════
// SECTION A: EXECUTIVE OVERVIEW
// ═══════════════════════════════════════════════════════════════

function addTitleSlide(pres, analysis) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);
  const meta = analysis.meta;

  slide.addText("\u{1F6E1}", {
    x: M, y: 0.5, w: 1.2, h: 1.2,
    fontSize: 48, align: "center", valign: "middle", margin: 0
  });

  slide.addText("CONDITIONAL ACCESS", {
    x: M, y: 1.9, w: 6, h: 0.35,
    fontSize: 13, fontFace: FONT, bold: true, color: C.teal,
    charSpacing: 4, margin: 0
  });

  slide.addText("Security Posture\nReport", {
    x: M, y: 2.25, w: 6, h: 1.5,
    fontSize: 40, fontFace: FONT, bold: true, color: C.white,
    margin: 0, lineSpacingMultiple: 1.0
  });

  slide.addText(meta.date, {
    x: M, y: 3.55, w: 4, h: 0.3,
    fontSize: 13, fontFace: FONT_LIGHT, color: C.muted, margin: 0
  });

  slide.addShape(pres.shapes.LINE, {
    x: M, y: 3.95, w: 5.5, h: 0,
    line: { color: C.teal, width: 1.5 }
  });

  const statsY = 4.1;
  const stats = [
    { n: meta.policyCount, label: "Policies", color: C.teal },
    { n: meta.enabledCount, label: "Enabled", color: C.green },
    { n: meta.reportOnlyCount, label: "Report-Only", color: C.amber },
    { n: meta.disabledCount, label: "Disabled", color: C.red },
  ];

  stats.forEach((s, i) => {
    const x = M + i * 1.7;
    slide.addText(String(s.n), {
      x, y: statsY, w: 1.2, h: 0.5,
      fontSize: 28, fontFace: FONT, bold: true, color: s.color, margin: 0
    });
    slide.addText(s.label, {
      x, y: statsY + 0.45, w: 1.5, h: 0.25,
      fontSize: 10, fontFace: FONT_LIGHT, color: C.muted, margin: 0
    });
  });

  if (meta.statsFooter) {
    slide.addText(meta.statsFooter, {
      x: M, y: 5.15, w: 6, h: 0.25,
      fontSize: 9, fontFace: FONT_LIGHT, color: C.mutedDark, margin: 0
    });
  }
}

function addExecutiveSummary(pres, analysis) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);
  const meta = analysis.meta;
  const es = analysis.executiveSummary;

  slide.addText("Executive Summary", {
    x: M, y: 0.25, w: W - 2 * M, h: 0.5,
    fontSize: 24, fontFace: FONT, bold: true, color: C.white, margin: 0
  });

  const badgeY = 0.85;
  const badges = [
    { n: meta.policyCount, label: "Total Policies", color: C.teal },
    { n: meta.enabledCount, label: "Enabled", color: C.green },
    { n: meta.reportOnlyCount, label: "Report-Only", color: C.amber },
    { n: meta.disabledCount, label: "Disabled", color: C.red },
  ];
  badges.forEach((b, i) => {
    const bx = M + i * 2.2;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: bx, y: badgeY, w: 2.0, h: 0.55,
      fill: { color: C.bgCard }, rectRadius: 0.05,
      line: { color: C.border, width: 0.5 }
    });
    slide.addText(String(b.n), {
      x: bx + 0.1, y: badgeY + 0.05, w: 0.6, h: 0.45,
      fontSize: 22, fontFace: FONT, bold: true, color: b.color, margin: 0
    });
    slide.addText(b.label, {
      x: bx + 0.7, y: badgeY + 0.1, w: 1.2, h: 0.35,
      fontSize: 10, fontFace: FONT_LIGHT, color: C.muted, margin: 0
    });
  });

  slide.addText("KEY FINDINGS", {
    x: M, y: 1.6, w: W - 2 * M, h: 0.25,
    fontSize: 9, fontFace: FONT, bold: true, color: C.teal, charSpacing: 2, margin: 0
  });

  const colW = (W - 2 * M - 0.4) / 2;
  const colY = 1.95;
  const itemH = 0.48;
  const colH = 0.5 + Math.max(es.strengths.length, es.concerns.length) * itemH + 0.1;

  // Strengths
  const strX = M;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: strX, y: colY, w: colW, h: colH,
    fill: { color: C.bgCard }, line: { color: C.border, width: 0.5 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: strX, y: colY, w: 0.04, h: colH,
    fill: { color: C.green }
  });
  slide.addText("Strengths", {
    x: strX + 0.15, y: colY + 0.1, w: colW - 0.3, h: 0.3,
    fontSize: 14, fontFace: FONT, bold: true, color: C.green, margin: 0
  });
  es.strengths.forEach((s, i) => {
    slide.addText("\u2713", {
      x: strX + 0.15, y: colY + 0.5 + i * itemH, w: 0.25, h: 0.3,
      fontSize: 12, fontFace: FONT, bold: true, color: C.green, margin: 0
    });
    slide.addText(s, {
      x: strX + 0.4, y: colY + 0.5 + i * itemH, w: colW - 0.6, h: 0.4,
      fontSize: 10, fontFace: FONT_LIGHT, color: C.white, margin: 0
    });
  });

  // Concerns
  const conX = strX + colW + 0.4;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: conX, y: colY, w: colW, h: colH,
    fill: { color: C.bgCard }, line: { color: C.border, width: 0.5 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: conX, y: colY, w: 0.04, h: colH,
    fill: { color: C.red }
  });
  slide.addText("Concerns", {
    x: conX + 0.15, y: colY + 0.1, w: colW - 0.3, h: 0.3,
    fontSize: 14, fontFace: FONT, bold: true, color: C.red, margin: 0
  });
  es.concerns.forEach((c, i) => {
    slide.addText("\u2717", {
      x: conX + 0.15, y: colY + 0.5 + i * itemH, w: 0.25, h: 0.3,
      fontSize: 12, fontFace: FONT, bold: true, color: C.red, margin: 0
    });
    slide.addText(c, {
      x: conX + 0.4, y: colY + 0.5 + i * itemH, w: colW - 0.6, h: 0.4,
      fontSize: 10, fontFace: FONT_LIGHT, color: C.white, margin: 0
    });
  });
}

function addPolicyLandscape(pres, analysis) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  const meta = analysis.meta;
  const pl = analysis.policyLandscape;

  slide.addText("Policy Landscape", {
    x: M, y: 0.25, w: W - 2 * M, h: 0.5,
    fontSize: 24, fontFace: FONT, bold: true, color: C.white, margin: 0
  });

  slide.addText("STATE DISTRIBUTION", {
    x: M, y: 0.85, w: 4, h: 0.25,
    fontSize: 9, fontFace: FONT, bold: true, color: C.teal, charSpacing: 2, margin: 0
  });

  const total = meta.policyCount;
  const barY = 1.2;
  const maxBarW = 5.0;
  const bars = [
    { n: meta.enabledCount, label: "Enabled", color: C.green },
    { n: meta.reportOnlyCount, label: "Report-Only", color: C.amber },
    { n: meta.disabledCount, label: "Disabled", color: C.red },
  ];

  bars.forEach((b, i) => {
    const by = barY + i * 0.7;
    const bw = Math.max(0.3, (b.n / total) * maxBarW);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: M + 1.5, y: by, w: bw, h: 0.4,
      fill: { color: b.color }, rectRadius: 0.05
    });
    slide.addText(b.label, {
      x: M, y: by, w: 1.4, h: 0.4,
      fontSize: 10, fontFace: FONT, color: C.muted, align: "right", margin: 0
    });
    slide.addText(String(b.n), {
      x: M + 1.5 + bw + 0.1, y: by, w: 0.5, h: 0.4,
      fontSize: 14, fontFace: FONT, bold: true, color: b.color, margin: 0
    });
  });

  slide.addText("CATEGORY BREAKDOWN", {
    x: M, y: 3.35, w: 4, h: 0.25,
    fontSize: 9, fontFace: FONT, bold: true, color: C.teal, charSpacing: 2, margin: 0
  });

  const categories = pl.categories;
  const catCount = Math.min(categories.length, 7);
  const totalW = W - 2 * M;
  const cardGap = 0.12;
  const cardW = (totalW - (catCount - 1) * cardGap) / catCount;
  const cardY = 3.7;

  categories.slice(0, 7).forEach((cat, i) => {
    const cx = M + i * (cardW + cardGap);
    const catColor = resolveColor(cat.colorKey);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 1.3,
      fill: { color: C.bgCard }, line: { color: C.border, width: 0.5 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 0.04,
      fill: { color: catColor }
    });
    slide.addText(String(cat.count), {
      x: cx, y: cardY + 0.15, w: cardW, h: 0.5,
      fontSize: 24, fontFace: FONT, bold: true, color: catColor,
      align: "center", margin: 0
    });
    slide.addText(cat.label, {
      x: cx + 0.05, y: cardY + 0.7, w: cardW - 0.1, h: 0.5,
      fontSize: 8, fontFace: FONT_LIGHT, color: C.muted,
      align: "center", margin: 0
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SECTION B: DEEP ANALYSIS (conditional slides)
// ═══════════════════════════════════════════════════════════════

function addGeolocationStrategy(pres, analysis) {
  const geo = analysis.geolocationStrategy;
  if (!geo || !geo.available) return;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);

  slide.addText("Geolocation Strategy", {
    x: M, y: 0.2, w: 5, h: 0.4,
    fontSize: 22, fontFace: FONT, bold: true, color: C.white, margin: 0
  });
  slide.addText("Layered Defense-in-Depth", {
    x: M, y: 0.6, w: 5, h: 0.3,
    fontSize: 12, fontFace: FONT_LIGHT, color: C.muted, margin: 0
  });

  const layerX = M;
  const layerW = W - 2 * M;
  const layers = geo.layers || [];

  let curY = 1.05;

  layers.forEach((layer) => {
    const accentColor = resolveColor(layer.accentColor);

    if (layer.countries) {
      // Layer with country list
      const countries = layer.countries || [];
      const rows = Math.ceil(countries.length / 4);
      const layerH = 0.6 + rows * 0.25 + 0.15;

      slide.addShape(pres.shapes.RECTANGLE, {
        x: layerX, y: curY, w: layerW, h: layerH,
        fill: { color: C.bgCard }, line: { color: C.border, width: 0.5 }
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: layerX, y: curY, w: 0.04, h: layerH,
        fill: { color: accentColor }
      });
      slide.addText(`LAYER ${layer.id}`, {
        x: layerX + 0.15, y: curY + 0.05, w: 1.0, h: 0.2,
        fontSize: 8, fontFace: FONT, bold: true, color: accentColor, charSpacing: 1.5, margin: 0
      });
      slide.addText(layer.title, {
        x: layerX + 0.15, y: curY + 0.3, w: 3, h: 0.2,
        fontSize: 11, fontFace: FONT, bold: true, color: C.white, margin: 0
      });

      countries.forEach((co, ci) => {
        const row = Math.floor(ci / 4);
        const col = ci % 4;
        const cx = layerX + 0.15 + col * 2.0;
        const cy = curY + 0.6 + row * 0.25;
        const sc = co.state === "Enabled" ? C.green : C.amber;
        slide.addText(co.name, {
          x: cx, y: cy, w: 1.3, h: 0.2,
          fontSize: 8, fontFace: FONT, color: C.white, margin: 0
        });
        slide.addText(co.state, {
          x: cx + 1.3, y: cy, w: 0.6, h: 0.2,
          fontSize: 7, fontFace: FONT, bold: true, color: sc, margin: 0
        });
      });

      curY += layerH + 0.15;
    } else if (layer.locationCount != null) {
      // Layer with location count
      const layerH = 0.95;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: layerX, y: curY, w: layerW, h: layerH,
        fill: { color: C.bgCard }, line: { color: C.border, width: 0.5 }
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: layerX, y: curY, w: 0.04, h: layerH,
        fill: { color: accentColor }
      });
      slide.addText(`LAYER ${layer.id}`, {
        x: layerX + 0.15, y: curY + 0.05, w: 1.0, h: 0.2,
        fontSize: 8, fontFace: FONT, bold: true, color: accentColor, charSpacing: 1.5, margin: 0
      });
      slide.addText(layer.title, {
        x: layerX + 0.15, y: curY + 0.3, w: 5, h: 0.2,
        fontSize: 11, fontFace: FONT, bold: true, color: C.white, margin: 0
      });
      // Count badge
      slide.addShape(pres.shapes.OVAL, {
        x: layerX + 6.5, y: curY + 0.1, w: 0.7, h: 0.7,
        fill: { color: C.circle }, line: { color: C.teal, width: 1 }
      });
      slide.addText(String(layer.locationCount), {
        x: layerX + 6.5, y: curY + 0.1, w: 0.7, h: 0.7,
        fontSize: 20, fontFace: FONT, bold: true, color: C.teal,
        align: "center", valign: "middle", margin: 0
      });
      slide.addText("LOCATIONS", {
        x: layerX + 7.3, y: curY + 0.3, w: 1.5, h: 0.2,
        fontSize: 8, fontFace: FONT, bold: true, color: C.muted, charSpacing: 1, margin: 0
      });
      slide.addText(layer.description, {
        x: layerX + 0.15, y: curY + 0.6, w: layerW - 0.3, h: 0.3,
        fontSize: 9, fontFace: FONT_LIGHT, color: C.muted, margin: 0
      });
      curY += layerH + 0.15;
    } else {
      // Simple description layer
      const layerH = 1.0;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: layerX, y: curY, w: layerW, h: layerH,
        fill: { color: C.bgCard }, line: { color: accentColor, width: 0.8 }
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: layerX, y: curY, w: 0.04, h: layerH,
        fill: { color: accentColor }
      });
      slide.addText(`LAYER ${layer.id}`, {
        x: layerX + 0.15, y: curY + 0.05, w: 1.0, h: 0.2,
        fontSize: 8, fontFace: FONT, bold: true, color: accentColor, charSpacing: 1.5, margin: 0
      });

      // Enabled badge
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: layerX + 1.2, y: curY + 0.03, w: 0.85, h: 0.22,
        fill: { color: C.greenDim }, rectRadius: 0.03
      });
      slide.addText("ENABLED", {
        x: layerX + 1.2, y: curY + 0.03, w: 0.85, h: 0.22,
        fontSize: 7, fontFace: FONT, bold: true, color: C.green, align: "center", margin: 0
      });

      slide.addText(layer.title, {
        x: layerX + 0.15, y: curY + 0.3, w: 3, h: 0.2,
        fontSize: 11, fontFace: FONT, bold: true, color: C.white, margin: 0
      });
      slide.addText(layer.description, {
        x: layerX + 0.15, y: curY + 0.55, w: layerW - 0.3, h: 0.35,
        fontSize: 9, fontFace: FONT_LIGHT, color: C.muted, margin: 0
      });
      curY += layerH + 0.15;
    }
  });

  if (geo.note) {
    slide.addText(`Note: ${geo.note}`, {
      x: M, y: Math.min(curY + 0.1, 4.95), w: W - 2 * M, h: 0.3,
      fontSize: 8, fontFace: FONT_LIGHT, italic: true, color: C.muted, margin: 0
    });
  }
}

function addMfaMatrix(pres, analysis) {
  const mfa = analysis.mfaMatrix;
  if (!mfa || !mfa.available) return;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);

  slide.addText("MFA Enforcement Matrix", {
    x: M, y: 0.2, w: W - 2 * M, h: 0.45,
    fontSize: 22, fontFace: FONT, bold: true, color: C.white, margin: 0
  });

  const headerRow = [
    { text: "Policy", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Scope", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Auth Strength", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Frequency", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Conditions", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
  ];

  const rows = [headerRow];
  mfa.policies.forEach((d, i) => {
    const rowFill = i % 2 === 0 ? C.bgCard : C.bg;
    rows.push([
      { text: d.name, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.white } },
      { text: d.scope, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.white } },
      { text: d.authStrength, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.white } },
      { text: d.frequency || "\u2014", options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.white } },
      { text: d.conditions || "\u2014", options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.white } },
    ]);
  });

  slide.addTable(rows, {
    x: M, y: 0.8, w: W - 2 * M,
    colW: [2.0, 2.0, 1.5, 1.0, 2.2],
    border: { pt: 0.5, color: C.border },
    rowH: 0.4, autoPage: false
  });

  if (mfa.callout) {
    const callY = 0.8 + (rows.length) * 0.4 + 0.3;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: callY, w: W - 2 * M, h: 0.7,
      fill: { color: C.bgCard }, line: { color: C.amber, width: 0.8 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: callY, w: 0.04, h: 0.7,
      fill: { color: C.amber }
    });
    slide.addText(mfa.callout.title, {
      x: M + 0.15, y: callY + 0.05, w: W - 2 * M - 0.3, h: 0.25,
      fontSize: 10, fontFace: FONT, bold: true, color: C.amber, margin: 0
    });
    slide.addText(mfa.callout.text, {
      x: M + 0.15, y: callY + 0.35, w: W - 2 * M - 0.3, h: 0.25,
      fontSize: 9, fontFace: FONT_LIGHT, color: C.muted, margin: 0
    });
  }
}

function addRiskPolicies(pres, analysis) {
  const risk = analysis.riskPolicies;
  if (!risk || !risk.available) return;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);

  slide.addText("Identity Protection & Risk-Based Policies", {
    x: M, y: 0.2, w: W - 2 * M, h: 0.45,
    fontSize: 22, fontFace: FONT, bold: true, color: C.white, margin: 0
  });

  const cards = risk.policies;
  const cardCount = Math.min(cards.length, 3);
  const cardGap = 0.3;
  const cardW = (W - 2 * M - (cardCount - 1) * cardGap) / cardCount;
  const cardY = 0.9;
  const cardH = 2.9;

  cards.slice(0, 3).forEach((card, i) => {
    const cx = M + i * (cardW + cardGap);
    const accentColor = resolveColor(card.accentColor);

    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: C.bgCard }, line: { color: C.border, width: 0.5 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 0.04,
      fill: { color: accentColor }
    });

    slide.addText(card.title, {
      x: cx + 0.15, y: cardY + 0.15, w: cardW - 0.3, h: 0.3,
      fontSize: 13, fontFace: FONT, bold: true, color: C.white, margin: 0
    });
    slide.addText(card.policyName, {
      x: cx + 0.15, y: cardY + 0.5, w: cardW - 0.3, h: 0.25,
      fontSize: 9, fontFace: FONT_LIGHT, color: C.muted, margin: 0
    });

    const fields = [
      { label: "Grant Control", value: card.grantControl },
      { label: "Operator", value: card.operator },
      { label: "Scope", value: card.scope },
      { label: "Modified:", value: card.modified },
    ];
    fields.forEach((f, fi) => {
      const fy = cardY + 0.9 + fi * 0.45;
      slide.addText(f.label, {
        x: cx + 0.15, y: fy, w: cardW - 0.3, h: 0.18,
        fontSize: 8, fontFace: FONT, bold: true, color: C.mutedDark, margin: 0
      });
      slide.addText(f.value || "\u2014", {
        x: cx + 0.15, y: fy + 0.18, w: cardW - 0.3, h: 0.2,
        fontSize: 10, fontFace: FONT, color: C.white, margin: 0
      });
    });
  });

  if (risk.callout) {
    const callY = 4.15;
    const callColor = resolveColor(risk.callout.color || "green");
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: callY, w: W - 2 * M, h: 0.6,
      fill: { color: C.bgCard }, line: { color: callColor, width: 0.8 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: callY, w: 0.04, h: 0.6,
      fill: { color: callColor }
    });
    slide.addText(risk.callout.text, {
      x: M + 0.15, y: callY + 0.1, w: W - 2 * M - 0.3, h: 0.4,
      fontSize: 10, fontFace: FONT, bold: true, color: callColor, margin: 0
    });
  }
}

function addAuthStrengths(pres, analysis) {
  const auth = analysis.authStrengths;
  if (!auth || !auth.available) return;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);

  slide.addText("Authentication Strength Policies", {
    x: M, y: 0.2, w: W - 2 * M, h: 0.45,
    fontSize: 22, fontFace: FONT, bold: true, color: C.white, margin: 0
  });

  const strengths = auth.strengths.slice(0, 4);
  const cardCount = strengths.length;
  const cardGap = 0.2;
  const cardW = (W - 2 * M - (cardCount - 1) * cardGap) / cardCount;
  const cardY = 0.85;
  const cardH = 4.2;

  strengths.forEach((s, i) => {
    const cx = M + i * (cardW + cardGap);
    const rColor = resolveColor(s.ratingColor);

    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: C.bgCard }, line: { color: C.border, width: 0.5 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cardY, w: cardW, h: 0.04,
      fill: { color: rColor }
    });

    slide.addText(s.name, {
      x: cx + 0.12, y: cardY + 0.2, w: cardW - 0.24, h: 0.35,
      fontSize: 12, fontFace: FONT, bold: true, color: C.white, margin: 0
    });

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx + 0.12, y: cardY + 0.65, w: 1.0, h: 0.25,
      fill: { color: resolveRatingDim(s.ratingColor) }, rectRadius: 0.03
    });
    slide.addText(s.rating, {
      x: cx + 0.12, y: cardY + 0.65, w: 1.0, h: 0.25,
      fontSize: 9, fontFace: FONT, bold: true, color: rColor,
      align: "center", valign: "middle", margin: 0
    });

    const fields = [
      { label: "Type:", value: s.type },
      { label: "Methods:", value: s.methods },
    ];
    fields.forEach((f, fi) => {
      const fy = cardY + 1.15 + fi * 0.7;
      slide.addText(f.label, {
        x: cx + 0.12, y: fy, w: cardW - 0.24, h: 0.2,
        fontSize: 8, fontFace: FONT, bold: true, color: C.mutedDark, margin: 0
      });
      slide.addText(f.value, {
        x: cx + 0.12, y: fy + 0.2, w: cardW - 0.24, h: 0.35,
        fontSize: 9, fontFace: FONT_LIGHT, color: C.white, margin: 0
      });
    });

    if (s.note) {
      const noteColor = s.rating === "Weak" ? C.red : C.muted;
      slide.addText(s.note, {
        x: cx + 0.12, y: cardY + 2.85, w: cardW - 0.24, h: 0.5,
        fontSize: 9, fontFace: FONT_LIGHT, italic: true, color: noteColor, margin: 0
      });
    }
  });
}

function addPimCoverage(pres, analysis) {
  const pim = analysis.pimCoverage;
  if (!pim || !pim.available) return;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);

  slide.addText("Privileged Access \u2014 PIM Role Coverage", {
    x: M, y: 0.2, w: W - 2 * M, h: 0.4,
    fontSize: 20, fontFace: FONT, bold: true, color: C.white, margin: 0
  });
  slide.addText("Cross-referencing PIM role assignments against Conditional Access policy targeting.", {
    x: M, y: 0.6, w: W - 2 * M, h: 0.25,
    fontSize: 9, fontFace: FONT_LIGHT, color: C.muted, margin: 0
  });

  const headerRow = [
    { text: "PIM Role", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Eligible", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
    { text: "Active", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
    { text: "CA Policy Coverage", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Direct?", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
  ];

  const rows = [headerRow];
  pim.roles.slice(0, 10).forEach((d, i) => {
    const rowFill = i % 2 === 0 ? C.bgCard : C.bg;
    const directColor = d.direct === "Yes" ? C.green : d.direct === "Partial" ? C.amber : C.red;
    rows.push([
      { text: d.role, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, color: C.white } },
      { text: d.eligible, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, color: C.muted, align: "center" } },
      { text: d.active, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, color: C.muted, align: "center" } },
      { text: d.caCoverage, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.white } },
      { text: d.direct, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, bold: true, color: directColor, align: "center" } },
    ]);
  });

  slide.addTable(rows, {
    x: M, y: 0.95, w: W - 2 * M,
    colW: [2.0, 0.7, 0.7, 3.4, 0.8],
    border: { pt: 0.5, color: C.border },
    rowH: 0.3, autoPage: false
  });

  if (pim.note) {
    slide.addText(pim.note, {
      x: M, y: 0.95 + rows.length * 0.3 + 0.15, w: W - 2 * M, h: 0.6,
      fontSize: 8, fontFace: FONT_LIGHT, italic: true, color: C.mutedDark, margin: 0
    });
  }
}

function addReportOnlyPipeline(pres, analysis) {
  const ro = analysis.reportOnlyPipeline;
  if (!ro || !ro.policies || ro.policies.length === 0) return;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);

  slide.addText("Report-Only Pipeline", {
    x: M, y: 0.2, w: 5, h: 0.4,
    fontSize: 22, fontFace: FONT, bold: true, color: C.white, margin: 0
  });
  slide.addText("Policies Under Evaluation", {
    x: M, y: 0.6, w: 5, h: 0.25,
    fontSize: 11, fontFace: FONT_LIGHT, color: C.muted, margin: 0
  });

  const headerRow = [
    { text: "Policy", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Grant Control", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Target Apps", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
    { text: "Modified", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
    { text: "Priority", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
  ];

  const rows = [headerRow];
  ro.policies.slice(0, 7).forEach((d, i) => {
    const rowFill = i % 2 === 0 ? C.bgCard : C.bg;
    const prioColor = d.priority === "High" ? C.red : d.priority === "Medium" ? C.amber : C.muted;
    rows.push([
      { text: d.name, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, color: C.white } },
      { text: d.grantControl, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.white } },
      { text: d.targetApps || "\u2014", options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.muted } },
      { text: d.modified || "\u2014", options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, color: C.muted, align: "center" } },
      { text: d.priority, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, bold: true, color: prioColor, align: "center" } },
    ]);
  });

  slide.addTable(rows, {
    x: M, y: 0.95, w: W - 2 * M,
    colW: [2.5, 1.5, 2.0, 1.1, 0.8],
    border: { pt: 0.5, color: C.border },
    rowH: 0.3, autoPage: false
  });

  if (ro.callout) {
    const callY = 0.95 + rows.length * 0.3 + 0.2;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: callY, w: W - 2 * M, h: 1.0,
      fill: { color: C.bgCard }, line: { color: C.red, width: 0.8 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: callY, w: 0.04, h: 1.0,
      fill: { color: C.red }
    });
    slide.addText(ro.callout.title, {
      x: M + 0.15, y: callY + 0.1, w: W - 2 * M - 0.3, h: 0.25,
      fontSize: 11, fontFace: FONT, bold: true, color: C.red, margin: 0
    });
    slide.addText(ro.callout.text, {
      x: M + 0.15, y: callY + 0.4, w: W - 2 * M - 0.3, h: 0.5,
      fontSize: 9, fontFace: FONT_LIGHT, color: C.muted, margin: 0
    });
  }
}

function addMsManagedOverlap(pres, analysis) {
  const ms = analysis.msManagedOverlap;
  if (!ms || !ms.policies || ms.policies.length === 0) return;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);

  slide.addText("Microsoft-Managed Policies", {
    x: M, y: 0.2, w: 5, h: 0.4,
    fontSize: 22, fontFace: FONT, bold: true, color: C.white, margin: 0
  });
  slide.addText(`Overlap Analysis \u2014 All ${ms.policies.length} auto-created policies are disabled`, {
    x: M, y: 0.6, w: W - 2 * M, h: 0.25,
    fontSize: 11, fontFace: FONT_LIGHT, color: C.muted, margin: 0
  });

  const headerRow = [
    { text: "MS-Managed Policy", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 7.5, fontFace: FONT } },
    { text: "What It Does", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 7.5, fontFace: FONT } },
    { text: "Custom Equivalent", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 7.5, fontFace: FONT } },
    { text: "Overlap", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 7.5, fontFace: FONT, align: "center" } },
  ];

  const rows = [headerRow];
  ms.policies.forEach((d, i) => {
    const rowFill = i % 2 === 0 ? C.bgCard : C.bg;
    const overlapColor = d.overlap === "Full" ? C.green : d.overlap === "Partial" ? C.amber : C.red;
    rows.push([
      { text: d.name, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, color: C.white } },
      { text: d.description, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.muted } },
      { text: d.customEquivalent, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT_LIGHT, color: C.white } },
      { text: d.overlap, options: { fill: { color: rowFill }, fontSize: 7.5, fontFace: FONT, bold: true, color: overlapColor, align: "center" } },
    ]);
  });

  slide.addTable(rows, {
    x: M, y: 0.95, w: W - 2 * M,
    colW: [2.5, 2.5, 2.3, 0.8],
    border: { pt: 0.5, color: C.border },
    rowH: 0.35, autoPage: false
  });

  if (ms.callout) {
    const callY = 0.95 + rows.length * 0.35 + 0.2;
    const callH = 1.2;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: callY, w: W - 2 * M, h: callH,
      fill: { color: C.bgCard }, line: { color: C.red, width: 0.8 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: callY, w: 0.04, h: callH,
      fill: { color: C.red }
    });
    slide.addText(ms.callout.title, {
      x: M + 0.15, y: callY + 0.1, w: W - 2 * M - 0.3, h: 0.3,
      fontSize: 12, fontFace: FONT, bold: true, color: C.red, margin: 0
    });
    slide.addText(ms.callout.text, {
      x: M + 0.15, y: callY + 0.45, w: W - 2 * M - 0.3, h: 0.65,
      fontSize: 9, fontFace: FONT_LIGHT, color: C.muted, margin: 0
    });
  }
}

function addRecommendations(pres, analysis) {
  const rec = analysis.recommendations;
  if (!rec) return;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addText("Security Gaps & Recommendations", {
    x: M, y: 0.25, w: W - 2 * M, h: 0.5,
    fontSize: 24, fontFace: FONT, bold: true, color: C.white, margin: 0
  });

  const groups = [
    { label: "HIGH", color: C.red, dimColor: C.redDim, items: rec.high || [] },
    { label: "MEDIUM", color: C.amber, dimColor: C.amberDim, items: rec.medium || [] },
    { label: "LOW", color: C.muted, dimColor: C.bgCard, items: rec.low || [] },
  ].filter(g => g.items.length > 0);

  let curY = 0.9;

  groups.forEach((group) => {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: M, y: curY, w: 0.8, h: 0.28,
      fill: { color: group.dimColor }, rectRadius: 0.05
    });
    slide.addText(group.label, {
      x: M, y: curY, w: 0.8, h: 0.28,
      fontSize: 9, fontFace: FONT, bold: true, color: group.color,
      align: "center", valign: "middle", margin: 0
    });

    curY += 0.35;

    group.items.forEach((item) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: M + 0.3, y: curY, w: W - 2 * M - 0.3, h: 0.35,
        fill: { color: C.bgCard }, line: { color: C.border, width: 0.5 }
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: M + 0.3, y: curY, w: 0.04, h: 0.35,
        fill: { color: group.color }
      });
      slide.addText(item, {
        x: M + 0.5, y: curY + 0.03, w: W - 2 * M - 0.8, h: 0.28,
        fontSize: 10, fontFace: FONT, color: C.white, margin: 0
      });
      curY += 0.4;
    });

    curY += 0.1;
  });
}

// ═══════════════════════════════════════════════════════════════
// SECTION C: FULL POLICY MATRIX
// ═══════════════════════════════════════════════════════════════

function addSummarySlides(pres, policies) {
  const rowsPerSlide = 18;
  const pages = Math.ceil(policies.length / rowsPerSlide);

  for (let page = 0; page < pages; page++) {
    const slide = pres.addSlide();
    slide.background = { color: C.bg };

    const start = page * rowsPerSlide;
    const subset = policies.slice(start, start + rowsPerSlide);

    slide.addText(pages > 1 ? `Full Policy Matrix (${page + 1}/${pages})` : "Full Policy Matrix", {
      x: M, y: 0.25, w: W - 2 * M, h: 0.45,
      fontSize: 20, fontFace: FONT, bold: true, color: C.white, margin: 0
    });

    const headerRow = [
      { text: "#", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
      { text: "Policy Name", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT } },
      { text: "State", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
      { text: "Action", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
      { text: "Last Modified", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 8, fontFace: FONT, align: "center" } },
    ];

    const rows = [headerRow];
    subset.forEach((p, i) => {
      const sc = stateColor(p.state);
      const rowFill = (start + i) % 2 === 0 ? C.bgCard : C.bg;
      const action = getAction(p);
      rows.push([
        { text: String(start + i + 1), options: { fill: { color: rowFill }, fontSize: 7, fontFace: FONT, align: "center", color: C.muted } },
        { text: p.name, options: { fill: { color: rowFill }, fontSize: 7, fontFace: FONT, color: C.white } },
        { text: sc.label, options: { fill: { color: sc.dim }, fontSize: 7, fontFace: FONT, bold: true, color: sc.accent, align: "center" } },
        { text: action, options: { fill: { color: rowFill }, fontSize: 7, fontFace: FONT, color: action === "Block" ? C.red : C.green, align: "center" } },
        { text: p.lastModified || "\u2014", options: { fill: { color: rowFill }, fontSize: 7, fontFace: FONT, color: C.muted, align: "center" } },
      ]);
    });

    slide.addTable(rows, {
      x: M, y: 0.8, w: W - 2 * M,
      colW: [0.35, 4.6, 1.3, 0.9, 1.05],
      border: { pt: 0.5, color: C.border },
      rowH: 0.24, autoPage: false
    });
  }
}

// ═══════════════════════════════════════════════════════════════
// SECTION D: PER-POLICY DETAIL
// ═══════════════════════════════════════════════════════════════

function addSectionDivider(pres, title, count, accentColor) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.05,
    fill: { color: accentColor }
  });

  slide.addText(title, {
    x: M, y: 2.0, w: W - 2 * M, h: 0.8,
    fontSize: 36, fontFace: FONT, bold: true, color: C.white,
    align: "center", margin: 0
  });

  slide.addText(`${count} ${count === 1 ? "policy" : "policies"}`, {
    x: M, y: 2.85, w: W - 2 * M, h: 0.4,
    fontSize: 15, fontFace: FONT_LIGHT, color: C.muted,
    align: "center", margin: 0
  });

  slide.addShape(pres.shapes.OVAL, {
    x: (W - 0.15) / 2, y: 3.4, w: 0.15, h: 0.15,
    fill: { color: accentColor }
  });
}

function addPolicySlide(pres, policy, index, total, genDate) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  const sc = stateColor(policy.state);
  const blocking = isBlockAccess(policy);

  // Header
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.82,
    fill: { color: C.bgCard }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.035,
    fill: { color: sc.accent }
  });

  slide.addText(policy.name, {
    x: M, y: 0.1, w: 6.8, h: 0.45,
    fontSize: 17, fontFace: FONT, bold: true, color: C.white,
    margin: 0, shrinkText: true
  });

  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.8, y: 0.15, w: 1.8, h: 0.3,
    fill: { color: sc.dim }, rectRadius: 0.05
  });
  slide.addText(sc.label, {
    x: 7.8, y: 0.15, w: 1.8, h: 0.3,
    fontSize: 9, fontFace: FONT, bold: true, color: sc.accent,
    align: "center", valign: "middle", margin: 0
  });

  if (policy.lastModified) {
    slide.addText(`Modified: ${policy.lastModified}`, {
      x: M, y: 0.55, w: 4, h: 0.2,
      fontSize: 8, fontFace: FONT_LIGHT, color: C.muted, margin: 0
    });
  }

  // Conditions bar
  const condY = 0.92;
  const condH = 0.55;
  const cond = policy.conditions || {};

  slide.addText("CONDITIONS", {
    x: M, y: condY, w: 2, h: 0.18,
    fontSize: 7, fontFace: FONT, bold: true, color: C.mutedDark,
    charSpacing: 1.5, margin: 0
  });

  const condItems = [
    { label: "Platforms", value: cond.platforms, excl: cond.platformsExclude },
    { label: "Locations", value: cond.locationsInclude, excl: cond.locationsExclude },
    { label: "Client Apps", value: cond.clientApps },
    { label: "Risk", value: cond.signInRisk || cond.userRisk },
    { label: "Devices", value: cond.deviceFilter || cond.authFlow },
  ];

  const condW = 1.72;
  const condGapVal = 0.1;
  condItems.forEach((ci, i) => {
    const cx = M + i * (condW + condGapVal);
    const configured = !!ci.value;
    const cardColor = configured ? C.bgCard : C.bgCardDim;
    const borderColor = configured ? C.borderTeal : C.border;

    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: condY + 0.2, w: condW, h: condH - 0.2,
      fill: { color: cardColor },
      line: { color: borderColor, width: 0.5 }
    });
    if (configured) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: condY + 0.2, w: 0.03, h: condH - 0.2,
        fill: { color: C.teal }
      });
    }
    slide.addText(ci.label, {
      x: cx + 0.08, y: condY + 0.22, w: condW - 0.15, h: 0.15,
      fontSize: 7, fontFace: FONT, bold: true,
      color: configured ? C.white : C.mutedDark, margin: 0
    });
    const valText = configured ? (ci.value.length > 25 ? ci.value.substring(0, 23) + "\u2026" : ci.value) : "Not configured";
    slide.addText(valText, {
      x: cx + 0.08, y: condY + 0.37, w: condW - 0.15, h: 0.13,
      fontSize: 6, fontFace: FONT_LIGHT,
      color: configured ? C.muted : C.mutedDark,
      margin: 0, shrinkText: true
    });
  });

  // Main flow area
  const flowY = 1.6;
  const flowH = 3.55;

  // Users card
  const usersX = M;
  const usersW = 2.0;
  addFlowCard(slide, pres, usersX, flowY, usersW, flowH * 0.55, "USERS", () => {
    const u = policy.users || {};
    const lines = [];
    if (u.includeUsers) lines.push({ label: "Include:", value: u.includeUsers, icon: "\u2705" });
    else if (!u.includeGroups && !u.includeRoles) lines.push({ label: "Include:", value: "All users", icon: "\u2705" });
    if (u.excludeUsers) lines.push({ label: "Exclude:", value: u.excludeUsers, icon: "\u{1F6AB}" });
    if (u.includeGroups) lines.push({ label: "Groups:", value: u.includeGroups, icon: "\u2705" });
    if (u.excludeGroups) lines.push({ label: "Groups (excl):", value: u.excludeGroups, icon: "\u{1F6AB}" });
    if (u.includeRoles) lines.push({ label: "Roles:", value: u.includeRoles, icon: "\u2705" });
    if (u.excludeRoles) lines.push({ label: "Roles (excl):", value: u.excludeRoles, icon: "\u{1F6AB}" });
    return lines;
  });

  // Grant/Block label + arrow
  const arrowX = usersX + usersW + 0.05;
  const arrowLabelY = flowY + 0.1;
  const actionLabel = blocking ? "Block access" : "Grant access";
  const actionColor = blocking ? C.blockRed : C.grantGreen;

  slide.addText(actionLabel, {
    x: arrowX, y: arrowLabelY, w: 3.2, h: 0.28,
    fontSize: 12, fontFace: FONT, bold: true, color: actionColor,
    align: "center", margin: 0
  });

  slide.addShape(pres.shapes.LINE, {
    x: arrowX + 0.1, y: arrowLabelY + 0.35, w: 3.0, h: 0,
    line: { color: actionColor, width: 2 }
  });
  slide.addText("\u25B6", {
    x: arrowX + 2.85, y: arrowLabelY + 0.23, w: 0.3, h: 0.25,
    fontSize: 10, color: actionColor, align: "center", margin: 0
  });

  // Grant Controls panel
  const grantX = arrowX + 0.15;
  const grantY = arrowLabelY + 0.55;
  const grantW = 2.9;
  const grantH = flowH - 0.7;

  slide.addShape(pres.shapes.RECTANGLE, {
    x: grantX, y: grantY, w: grantW, h: grantH,
    fill: { color: C.bgCard },
    line: { color: C.border, width: 0.5 }
  });

  slide.addText("GRANT CONTROLS", {
    x: grantX + 0.1, y: grantY + 0.05, w: grantW - 0.2, h: 0.18,
    fontSize: 7, fontFace: FONT, bold: true, color: C.mutedDark,
    charSpacing: 1.5, margin: 0
  });

  const g = policy.grantControls || {};
  if (g.operator && g.operator !== "Not configured") {
    slide.addText(g.operator, {
      x: grantX + 0.1, y: grantY + 0.25, w: grantW - 0.2, h: 0.15,
      fontSize: 7, fontFace: FONT, bold: true, color: C.muted, margin: 0
    });
  }

  const activeGrants = getActiveGrantControls(policy);
  const grantStartY = grantY + 0.45;
  ALL_GRANT_CONTROLS.forEach((ctrl, i) => {
    const active = activeGrants.has(ctrl);
    const cy = grantStartY + i * 0.2;
    const dot = active ? "\u25CF" : "\u25CB";
    const textColor = active ? C.white : C.mutedDark;
    const dotColor = active ? C.teal : C.mutedDark;

    slide.addText(dot, {
      x: grantX + 0.1, y: cy, w: 0.15, h: 0.17,
      fontSize: 7, color: dotColor, margin: 0
    });

    let ctrlLabel = ctrl;
    if (ctrl === "Authentication strength" && g.authStrength) {
      ctrlLabel = `Auth strength: ${g.authStrength}`;
    }
    slide.addText(ctrlLabel, {
      x: grantX + 0.28, y: cy, w: grantW - 0.4, h: 0.17,
      fontSize: 7.5, fontFace: active ? FONT : FONT_LIGHT, bold: active,
      color: textColor, margin: 0, shrinkText: true
    });
  });

  // Apps card
  const appsX = grantX + grantW + 0.15;
  const appsW = 2.0;
  addFlowCard(slide, pres, appsX, flowY, appsW, flowH * 0.55, "APPS", () => {
    const a = policy.applications || {};
    const lines = [];
    if (a.include) lines.push({ label: "Include:", value: a.include, icon: "\u2705" });
    if (a.exclude) lines.push({ label: "Exclude:", value: a.exclude, icon: "\u{1F6AB}" });
    if (a.userActions) lines.push({ label: "Actions:", value: a.userActions });
    if (a.authContext) lines.push({ label: "Auth ctx:", value: a.authContext });
    if (lines.length === 0) lines.push({ label: "", value: "All cloud apps", icon: "\u2705" });
    return lines;
  });

  // Session Controls panel
  const sessX = appsX;
  const sessY = flowY + flowH * 0.55 + 0.08;
  const sessW = appsW;
  const sessH = flowH * 0.45 - 0.08;

  slide.addShape(pres.shapes.RECTANGLE, {
    x: sessX, y: sessY, w: sessW, h: sessH,
    fill: { color: C.bgCard },
    line: { color: C.border, width: 0.5 }
  });

  slide.addText("SESSION CONTROLS", {
    x: sessX + 0.08, y: sessY + 0.05, w: sessW - 0.16, h: 0.18,
    fontSize: 7, fontFace: FONT, bold: true, color: C.mutedDark,
    charSpacing: 1, margin: 0
  });

  const sessData = policy.sessionControls || {};
  const activeSign = !!sessData.signInFrequency;
  const sessStartY = sessY + 0.28;

  ALL_SESSION_CONTROLS.forEach((ctrl, i) => {
    const cy = sessStartY + i * 0.17;
    let active = false;
    let detail = "";
    if (ctrl === "Sign-in frequency" && activeSign) {
      active = true;
      detail = sessData.signInFrequency;
    }

    const dot = active ? "\u25CF" : "\u25CB";
    slide.addText(dot, {
      x: sessX + 0.08, y: cy, w: 0.12, h: 0.15,
      fontSize: 6, color: active ? C.teal : C.mutedDark, margin: 0
    });

    const label = detail ? `${ctrl}: ${detail}` : ctrl;
    slide.addText(label, {
      x: sessX + 0.22, y: cy, w: sessW - 0.32, h: 0.15,
      fontSize: 6, fontFace: active ? FONT : FONT_LIGHT, bold: active,
      color: active ? C.white : C.mutedDark, margin: 0, shrinkText: true
    });
  });

  // Footer
  slide.addText(`Policy ${index + 1} of ${total}`, {
    x: M, y: H - 0.3, w: 3, h: 0.2,
    fontSize: 7, fontFace: FONT_LIGHT, color: C.mutedDark, margin: 0
  });
  slide.addText(genDate, {
    x: W - M - 3, y: H - 0.3, w: 3, h: 0.2,
    fontSize: 7, fontFace: FONT_LIGHT, color: C.mutedDark, align: "right", margin: 0
  });
}

function addFlowCard(slide, pres, x, y, w, h, title, getLines) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.bgCard },
    line: { color: C.border, width: 0.5 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.03, h,
    fill: { color: C.teal }
  });

  slide.addText(title, {
    x: x + 0.1, y: y + 0.05, w: w - 0.2, h: 0.18,
    fontSize: 7, fontFace: FONT, bold: true, color: C.mutedDark,
    charSpacing: 1.5, margin: 0
  });

  const lines = getLines();
  const textItems = [];
  lines.forEach((line) => {
    if (line.label) {
      textItems.push({
        text: (line.icon ? line.icon + " " : "") + line.label + " ",
        options: { fontSize: 8, fontFace: FONT, bold: true, color: C.muted, breakLine: false }
      });
    }
    textItems.push({
      text: line.value,
      options: { fontSize: 8, fontFace: FONT_LIGHT, color: C.white, breakLine: true }
    });
  });

  slide.addText(textItems, {
    x: x + 0.08, y: y + 0.27, w: w - 0.16, h: h - 0.35,
    valign: "top", margin: 0, shrinkText: true, paraSpaceAfter: 3
  });
}

// ═══════════════════════════════════════════════════════════════
// SECTION E: CLOSING
// ═══════════════════════════════════════════════════════════════

function addClosingSlide(pres, analysis) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  addBgCircles(slide, pres);
  const meta = analysis.meta;
  const assess = analysis.assessment;

  slide.addText("Security Posture Assessment", {
    x: M, y: 1.8, w: W - 2 * M, h: 0.8,
    fontSize: 32, fontFace: FONT, bold: true, color: C.white,
    align: "center", margin: 0
  });

  // Verdict badge
  const verdictColor = resolveColor(assess.verdictColor);
  const verdictDim = assess.verdictColor === "red" ? C.redDim : assess.verdictColor === "green" ? C.greenDim : C.amberDim;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: (W - 5) / 2, y: 2.7, w: 5, h: 0.4,
    fill: { color: verdictDim }, rectRadius: 0.05
  });
  slide.addText(assess.verdict, {
    x: (W - 5) / 2, y: 2.7, w: 5, h: 0.4,
    fontSize: 13, fontFace: FONT, bold: true, color: verdictColor,
    align: "center", valign: "middle", margin: 0
  });

  slide.addText(assess.prioritySummary, {
    x: M, y: 3.25, w: W - 2 * M, h: 0.3,
    fontSize: 11, fontFace: FONT_LIGHT, color: C.muted,
    align: "center", margin: 0
  });

  if (assess.criticalGap) {
    slide.addText(`Critical gap: ${assess.criticalGap}`, {
      x: M, y: 3.6, w: W - 2 * M, h: 0.3,
      fontSize: 10, fontFace: FONT, bold: true, color: C.red,
      align: "center", margin: 0
    });
  }

  slide.addText(`${meta.policyCount} policies  \u2022  ${meta.enabledCount} enabled  \u2022  ${meta.reportOnlyCount} report-only  \u2022  ${meta.disabledCount} disabled`, {
    x: M, y: 4.1, w: W - 2 * M, h: 0.3,
    fontSize: 11, fontFace: FONT_LIGHT, color: C.mutedDark,
    align: "center", margin: 0
  });

  slide.addShape(pres.shapes.LINE, {
    x: 3, y: 4.55, w: 4, h: 0,
    line: { color: C.border, width: 0.5 }
  });

  slide.addText(`Generated: ${meta.date}  |  Data: ${assess.dataSources}`, {
    x: M, y: 4.65, w: W - 2 * M, h: 0.25,
    fontSize: 8, fontFace: FONT_LIGHT, color: C.mutedDark,
    align: "center", margin: 0
  });

  if (meta.nextReview) {
    slide.addText(`Next Review: ${meta.nextReview}`, {
      x: M, y: 4.9, w: W - 2 * M, h: 0.25,
      fontSize: 8, fontFace: FONT_LIGHT, color: C.mutedDark,
      align: "center", margin: 0
    });
  }
}

// ═══════════════════════════════════════════════════════════════
// MAIN
// ═══════════════════════════════════════════════════════════════

async function main() {
  // Resolve paths relative to this script's directory
  const scriptDir = __dirname;
  const projectDir = path.resolve(scriptDir, "..");

  const analysisPath = path.resolve(projectDir, "analysis.json");
  const policiesPath = path.resolve(projectDir, "policies.json");

  if (!fs.existsSync(analysisPath)) {
    console.error("Error: analysis.json not found at", analysisPath);
    console.error("Run the CA Documenter skill first to generate analysis.json.");
    process.exit(1);
  }
  if (!fs.existsSync(policiesPath)) {
    console.error("Error: policies.json not found at", policiesPath);
    process.exit(1);
  }

  const analysis = JSON.parse(fs.readFileSync(analysisPath, "utf8"));
  const policies = JSON.parse(fs.readFileSync(policiesPath, "utf8"));
  const meta = analysis.meta;

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "CA Documenter";
  pres.title = "Conditional Access Policies - Security Posture Report";

  const enabled = policies.filter(p => p.state === "enabled");
  const reportOnly = policies.filter(p => p.state === "report_only");
  const disabled = policies.filter(p => p.state === "disabled");

  // Section A: Executive Overview
  addTitleSlide(pres, analysis);
  addExecutiveSummary(pres, analysis);
  addPolicyLandscape(pres, analysis);

  // Section B: Deep Analysis (conditional)
  addGeolocationStrategy(pres, analysis);
  addMfaMatrix(pres, analysis);
  addRiskPolicies(pres, analysis);
  addAuthStrengths(pres, analysis);
  addPimCoverage(pres, analysis);
  addReportOnlyPipeline(pres, analysis);
  addMsManagedOverlap(pres, analysis);
  addRecommendations(pres, analysis);

  // Section C: Full Policy Matrix
  addSummarySlides(pres, policies);

  // Section D: Per-Policy Detail
  addSectionDivider(pres, "Enabled Policies", enabled.length, C.green);
  enabled.forEach((p, i) => addPolicySlide(pres, p, i, policies.length, meta.date));

  addSectionDivider(pres, "Report-Only Policies", reportOnly.length, C.amber);
  reportOnly.forEach((p, i) => addPolicySlide(pres, p, enabled.length + i, policies.length, meta.date));

  addSectionDivider(pres, "Disabled Policies", disabled.length, C.red);
  disabled.forEach((p, i) => addPolicySlide(pres, p, enabled.length + reportOnly.length + i, policies.length, meta.date));

  // Section E: Closing
  addClosingSlide(pres, analysis);

  const outputPath = path.resolve(projectDir, "CA_Security_Posture_Report.pptx");
  await pres.writeFile({ fileName: outputPath });
  console.log(`Generated ${outputPath}`);
}

main().catch(console.error);
