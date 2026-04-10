"use strict";

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const defaultTheme = require("./theme.default");

const SLIDE = {
  W: 10,
  H: 5.625,
  M: 0.42,
};

let _skipGuidSanitize = false;

const CONTENT_LIMITS = {
  executiveCardsPerSlide: 4,
  roadmapItemsPerColumn: 5,
  categoryCardsPerSlide: 6,
  layerCardsPerSlide: 3,
  riskCardsPerSlide: 3,
  authCardsPerSlide: 4,
  tableRows: {
    mfa: 7,
    pim: 10,
    reportOnly: 9,
    msManaged: 8,
    matrix: 14,
  },
};

const ALL_GRANT_CONTROLS = [
  "Multifactor authentication",
  "Authentication strength",
  "Compliant device",
  "Hybrid Azure AD joined",
  "Approved client app",
  "App protection policy",
  "Change password",
  "Terms of use",
  "Block access",
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

const KNOWN_LABELS = {
  all: "All",
  none: "None",
  "00000003-0000-0000-c000-000000000000": "Microsoft Graph",
  "00000002-0000-0ff1-ce00-000000000000": "Office 365 Exchange Online",
  "00000003-0000-0ff1-ce00-000000000000": "Microsoft SharePoint Online",
  "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe": "Microsoft Teams Services",
  Office365: "Office 365",
  MicrosoftAdminPortals: "Microsoft Admin Portals",
};

function parseArgs(argv) {
  const options = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--analysis") { options.analysis = argv[i + 1]; i += 1; }
    else if (arg === "--policies") { options.policies = argv[i + 1]; i += 1; }
    else if (arg === "--output") { options.output = argv[i + 1]; i += 1; }
    else if (arg === "--theme") { options.theme = argv[i + 1]; i += 1; }
    else if (arg === "--help" || arg === "-h") { options.help = true; }
  }
  return options;
}

function printHelp() {
  console.log("Usage: node skill/generate_report.js [options]");
  console.log("");
  console.log("Options:");
  console.log("  --analysis <path>   Path to analysis JSON (default: analysis.json)");
  console.log("  --policies <path>   Path to policies JSON (default: policies.json)");
  console.log("  --output <path>     Output PPTX path (default: CA_Security_Posture_Report.pptx)");
  console.log("  --theme <path>      Optional theme JSON file");
  console.log("  --help              Show this help");
}

function deepClone(value) {
  return structuredClone(value);
}

function safePath(candidate, projectDir) {
  const resolved = path.isAbsolute(candidate) ? candidate : path.resolve(projectDir, candidate);
  const normalizedResolved = path.resolve(resolved);
  const normalizedProject = path.resolve(projectDir) + path.sep;
  if (!normalizedResolved.startsWith(normalizedProject) && normalizedResolved !== path.resolve(projectDir)) {
    throw new Error(`Path escapes project directory: ${candidate}`);
  }
  return normalizedResolved;
}

function isObject(value) {
  return value && typeof value === "object" && !Array.isArray(value);
}

function deepMerge(base, override) {
  if (!isObject(base) || !isObject(override)) return override;
  const merged = { ...base };
  Object.keys(override).forEach((key) => {
    const next = override[key];
    if (isObject(next) && isObject(base[key])) {
      merged[key] = deepMerge(base[key], next);
      return;
    }
    merged[key] = next;
  });
  return merged;
}

function loadTheme(themePath, projectDir) {
  if (!themePath) return deepClone(defaultTheme);
  const resolved = safePath(themePath, projectDir);
  if (!fs.existsSync(resolved)) {
    throw new Error(`Theme file not found: ${resolved}`);
  }
  if (!resolved.endsWith(".json")) {
    throw new Error(`Theme file must be JSON: ${resolved}`);
  }
  const customTheme = JSON.parse(fs.readFileSync(resolved, "utf8"));
  return deepMerge(deepClone(defaultTheme), customTheme);
}

function ensureFileExists(filePath, friendlyName) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`${friendlyName} not found at ${filePath}`);
  }
}

function validateAnalysis(analysis) {
  const required = ["meta", "executiveSummary", "recommendations"];
  const missing = required.filter((key) => !analysis[key]);
  if (missing.length) {
    throw new Error(`Analysis JSON is missing required sections: ${missing.join(", ")}`);
  }
}

function readJson(filePath, friendlyName) {
  ensureFileExists(filePath, friendlyName);
  const raw = fs.readFileSync(filePath, "utf8");
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error(`Failed to parse ${friendlyName} at ${filePath}: ${err.message}`);
  }
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function hasValue(value) {
  if (Array.isArray(value)) return value.length > 0;
  if (value === null || value === undefined) return false;
  return String(value).trim() !== "";
}

function pickFirst(...values) {
  for (let i = 0; i < values.length; i += 1) {
    if (hasValue(values[i])) return values[i];
  }
  return null;
}

function toArray(value) {
  if (Array.isArray(value)) return value.filter((item) => hasValue(item));
  if (!hasValue(value)) return [];
  if (typeof value === "string") {
    return value
      .split(",")
      .map((item) => item.trim())
      .filter((item) => item);
  }
  return [value];
}

function cleanWhitespace(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function humanizeIdentifier(value) {
  let text = String(value || "").replace(/[_]+/g, " ");
  if (!_skipGuidSanitize) {
    text = text.replace(/\b[a-f0-9]{8}-[a-f0-9-]{27}\b/gi, "Directory object");
  }
  return cleanWhitespace(text);
}

function detectRawGuids(policiesRaw) {
  const guidRe = /\b[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}\b/i;
  const json = JSON.stringify(policiesRaw);
  const matches = json.match(new RegExp(guidRe.source, "gi")) || [];
  // A few GUIDs may appear as well-known app IDs even in clean data; only flag as raw if many are found
  return matches.length > 6;
}

function mapKnownLabel(value) {
  const direct = KNOWN_LABELS[value];
  if (direct) return direct;
  const lowered = String(value || "").toLowerCase();
  if (KNOWN_LABELS[lowered]) return KNOWN_LABELS[lowered];
  return value;
}

function sanitizeText(value, maxLen = 100) {
  if (!hasValue(value)) return "Not configured";
  const text = cleanWhitespace(humanizeIdentifier(mapKnownLabel(value)));
  if (text.length <= maxLen) return text;
  return `${text.slice(0, Math.max(0, maxLen - 1)).trim()}...`;
}

function summarizeList(value, options = {}) {
  const maxItems = options.maxItems || 3;
  const maxLen = options.maxLen || 92;
  const items = toArray(value).map((item) => sanitizeText(item, 42));
  if (!items.length) return "Not configured";
  if (items.length === 1) return sanitizeText(items[0], maxLen);
  if (items.length <= maxItems) return sanitizeText(items.join(", "), maxLen);
  const visible = items.slice(0, maxItems).join(", ");
  return sanitizeText(`${visible}, +${items.length - maxItems} more`, maxLen);
}

function normalizeDate(value) {
  if (!hasValue(value)) return "Not provided";
  const asDate = new Date(value);
  if (Number.isNaN(asDate.valueOf())) return sanitizeText(value, 24);
  return asDate.toISOString().slice(0, 10);
}

function normalizeState(state) {
  if (!hasValue(state)) return "disabled";
  const raw = String(state).toLowerCase();
  if (raw.includes("report") || raw.includes("enabledforreporting")) return "report_only";
  if (raw === "enabled") return "enabled";
  return "disabled";
}

function mapStateLabel(state) {
  if (state === "enabled") return "Enabled";
  if (state === "report_only") return "Report-only";
  return "Disabled";
}

function mapStateTone(state) {
  if (state === "enabled") return "positive";
  if (state === "report_only") return "caution";
  return "critical";
}

function isConfiguredValue(value) {
  return hasValue(value) && String(value).toLowerCase() !== "not configured";
}

function getToneColors(theme, tone) {
  if (tone === "positive") return { strong: theme.palette.positive, soft: theme.palette.positiveSoft };
  if (tone === "caution") return { strong: theme.palette.caution, soft: theme.palette.cautionSoft };
  if (tone === "critical") return { strong: theme.palette.critical, soft: theme.palette.criticalSoft };
  return { strong: theme.palette.neutral, soft: theme.palette.neutralSoft };
}

function getSeverityTone(priority) {
  const raw = String(priority || "").toLowerCase();
  if (raw.includes("high")) return "critical";
  if (raw.includes("medium")) return "caution";
  return "neutral";
}

function formatSignInFrequency(value) {
  if (!hasValue(value)) return null;
  if (typeof value === "string") return sanitizeText(value, 32);
  if (typeof value === "number") return `${value} hours`;
  if (isObject(value) && hasValue(value.value)) {
    const unit = hasValue(value.type) ? String(value.type) : "days";
    return `${value.value} ${unit}`;
  }
  return sanitizeText(JSON.stringify(value), 32);
}

function normalizeGrantControls(policy) {
  const raw = pickFirst(policy.grantControls, policy.conditions && policy.conditions.grantControls, {});
  const controls = toArray(pickFirst(raw.controls, raw.builtInControls)).map((value) => sanitizeText(value, 40));
  const authStrength = sanitizeText(
    pickFirst(raw.authStrength, raw.authenticationStrength && raw.authenticationStrength.displayName, raw.authenticationStrength),
    48
  );
  return {
    operator: sanitizeText(pickFirst(raw.operator, "Not configured"), 20),
    controls,
    authStrength: authStrength === "Not configured" ? null : authStrength,
  };
}

function normalizeSessionControls(policy) {
  const raw = pickFirst(policy.sessionControls, policy.conditions && policy.conditions.sessionControls, {});
  const active = [];
  const signInFrequency = formatSignInFrequency(raw.signInFrequency);
  if (signInFrequency) active.push(`Sign-in frequency (${signInFrequency})`);
  if (hasValue(raw.appEnforcedRestrictions) || hasValue(raw.applicationEnforcedRestrictions)) {
    active.push("App enforced restrictions");
  }
  if (hasValue(raw.cloudAppSecurity)) active.push("Conditional Access App Control");
  if (hasValue(raw.persistentBrowser)) active.push(`Persistent browser (${sanitizeText(raw.persistentBrowser, 20)})`);
  if (hasValue(raw.continuousAccessEvaluation)) active.push("Continuous access evaluation");
  if (hasValue(raw.disableResilienceDefaults)) active.push("Disable resilience defaults");
  if (hasValue(raw.tokenProtection)) active.push("Token protection");
  return {
    signInFrequency,
    active,
  };
}

function normalizeUsers(policy) {
  const source = pickFirst(policy.users, policy.conditions && policy.conditions.users, {});
  const includeUsers = summarizeList(pickFirst(source.includeUsers, source.include, source.users), { maxItems: 2 });
  const excludeUsers = summarizeList(pickFirst(source.excludeUsers, source.exclude), { maxItems: 2, maxLen: 80 });
  const includeGroups = summarizeList(source.includeGroups, { maxItems: 2, maxLen: 80 });
  const excludeGroups = summarizeList(source.excludeGroups, { maxItems: 2, maxLen: 80 });
  const includeRoles = summarizeList(source.includeRoles, { maxItems: 2, maxLen: 80 });
  const excludeRoles = summarizeList(source.excludeRoles, { maxItems: 2, maxLen: 80 });
  return {
    includeUsers,
    excludeUsers,
    includeGroups,
    excludeGroups,
    includeRoles,
    excludeRoles,
  };
}

function normalizeApplications(policy) {
  const source = pickFirst(policy.applications, policy.conditions && policy.conditions.applications, {});
  return {
    include: summarizeList(pickFirst(source.include, source.includeApplications), { maxItems: 2 }),
    exclude: summarizeList(pickFirst(source.exclude, source.excludeApplications), { maxItems: 2 }),
    userActions: summarizeList(source.userActions, { maxItems: 2 }),
    authContext: summarizeList(source.authContext, { maxItems: 2 }),
  };
}

function normalizeConditions(policy) {
  const source = pickFirst(policy.conditions, {});
  const platformsSource = pickFirst(
    source.platformsInclude,
    source.platforms && source.platforms.includePlatforms,
    Array.isArray(source.platforms) || typeof source.platforms === "string" ? source.platforms : null
  );
  const locationsIncludeSource = pickFirst(
    source.locationsInclude,
    source.locations && source.locations.includeLocations,
    Array.isArray(source.locations) || typeof source.locations === "string" ? source.locations : null
  );
  const locationsExcludeSource = pickFirst(
    source.locationsExclude,
    source.locations && source.locations.excludeLocations,
    source.excludeLocations
  );
  const clientAppsSource = pickFirst(
    source.clientApps,
    source.clientAppTypes,
    source.clientApplications
  );
  const riskSource = pickFirst(
    source.signInRisk,
    source.userRisk,
    source.signInRiskLevels,
    source.userRiskLevels
  );
  const devicesSource = pickFirst(
    source.deviceFilter,
    source.authFlow,
    source.authenticationFlows,
    source.devices && source.devices.deviceFilter
  );
  const platforms = summarizeList(
    platformsSource,
    { maxItems: 2, maxLen: 70 }
  );
  const locationsInclude = summarizeList(
    locationsIncludeSource,
    { maxItems: 2, maxLen: 70 }
  );
  const locationsExclude = summarizeList(
    locationsExcludeSource,
    { maxItems: 2, maxLen: 70 }
  );
  const clientApps = summarizeList(
    clientAppsSource,
    { maxItems: 2, maxLen: 70 }
  );
  const risk = summarizeList(
    riskSource,
    { maxItems: 2, maxLen: 70 }
  );
  const devices = summarizeList(
    devicesSource,
    { maxItems: 2, maxLen: 70 }
  );
  return {
    platforms,
    locationsInclude,
    locationsExclude,
    clientApps,
    risk,
    devices,
  };
}

function normalizePolicies(rawPolicies) {
  const list = Array.isArray(rawPolicies)
    ? rawPolicies
    : Array.isArray(rawPolicies && rawPolicies.value)
      ? rawPolicies.value
      : [];

  return list.map((policy, index) => {
    const grantControls = normalizeGrantControls(policy);
    return {
      id: sanitizeText(pickFirst(policy.id, `policy-${index + 1}`), 40),
      name: sanitizeText(pickFirst(policy.name, policy.displayName, `Policy ${index + 1}`), 86),
      state: normalizeState(policy.state),
      lastModified: normalizeDate(pickFirst(policy.lastModified, policy.modifiedDateTime)),
      users: normalizeUsers(policy),
      applications: normalizeApplications(policy),
      conditions: normalizeConditions(policy),
      grantControls,
      sessionControls: normalizeSessionControls(policy),
      action: grantControls.controls.some((item) => item.toLowerCase().includes("block")) ? "Block" : "Grant",
    };
  });
}

function chunk(list, size) {
  if (size <= 0) return [list];
  const pages = [];
  for (let i = 0; i < list.length; i += size) {
    pages.push(list.slice(i, i + size));
  }
  return pages;
}

function inferAssessment(analysis, policies) {
  const assessment = analysis.assessment || {};
  const enabled = policies.filter((policy) => policy.state === "enabled").length;
  const reportOnly = policies.filter((policy) => policy.state === "report_only").length;
  const scoreDerived = clamp(Math.round((((enabled * 1) + (reportOnly * 0.5)) / Math.max(1, policies.length)) * 100), 0, 100);
  const score = Number.isFinite(assessment.score) ? clamp(Math.round(assessment.score), 0, 100) : scoreDerived;
  const level = hasValue(assessment.level)
    ? sanitizeText(assessment.level, 32)
    : score >= 85
      ? "Strong"
      : score >= 70
        ? "Stable"
        : score >= 55
          ? "Needs Improvement"
          : "At Risk";
  return {
    score,
    level,
    verdict: sanitizeText(pickFirst(assessment.verdict, "Posture review complete"), 72),
    prioritySummary: sanitizeText(
      pickFirst(assessment.prioritySummary, "Priorities are grouped in the next section."),
      110
    ),
    criticalGap: hasValue(assessment.criticalGap) ? sanitizeText(assessment.criticalGap, 100) : null,
  };
}

function normalizePriorityItem(item, fallbackPriority) {
  if (typeof item === "string") {
    return {
      title: sanitizeText(item, 94),
      priority: fallbackPriority,
      evidence: null,
    };
  }
  return {
    title: sanitizeText(pickFirst(item.title, item.action, item.text), 94),
    priority: sanitizeText(pickFirst(item.priority, fallbackPriority), 16),
    evidence: hasValue(item.evidence) ? sanitizeText(item.evidence, 94) : null,
  };
}

function collectTopPriorities(analysis) {
  const fromExecutive = toArray(analysis.executiveSummary && analysis.executiveSummary.topPriorities).map((item) =>
    normalizePriorityItem(item, "High")
  );
  if (fromExecutive.length) return fromExecutive;

  const recommendations = analysis.recommendations || {};
  const derived = [];
  toArray(recommendations.high).forEach((item) => derived.push(normalizePriorityItem(item, "High")));
  toArray(recommendations.medium).forEach((item) => derived.push(normalizePriorityItem(item, "Medium")));
  toArray(recommendations.low).forEach((item) => derived.push(normalizePriorityItem(item, "Low")));
  return derived.slice(0, 8);
}

function normalizeRoadmap(analysis) {
  const roadmap = analysis.roadmap || {};
  const nearTerm = toArray(pickFirst(roadmap.nearTerm, roadmap.near_term)).map((item) =>
    normalizePriorityItem(item, "Near-term")
  );
  const midTerm = toArray(pickFirst(roadmap.midTerm, roadmap.mid_term)).map((item) =>
    normalizePriorityItem(item, "Mid-term")
  );
  return { nearTerm, midTerm };
}

function deriveRoadmapFromRecommendations(analysis) {
  const recommendations = analysis.recommendations || {};
  const nearTerm = toArray(recommendations.high).slice(0, 6).map((item) => normalizePriorityItem(item, "Near-term"));
  const midTerm = toArray(recommendations.medium).slice(0, 6).map((item) => normalizePriorityItem(item, "Mid-term"));
  return { nearTerm, midTerm };
}

function buildRoadmap(analysis) {
  const normalized = normalizeRoadmap(analysis);
  if (normalized.nearTerm.length || normalized.midTerm.length) return normalized;
  return deriveRoadmapFromRecommendations(analysis);
}

function resolveLogoPath(logoPath, themeLogoPath, projectDir) {
  const candidate = pickFirst(logoPath, themeLogoPath);
  if (!candidate) return null;
  const resolved = safePath(candidate, projectDir);
  return fs.existsSync(resolved) ? resolved : null;
}

function addDeckBackground(slide, ctx) {
  const bgPath = path.resolve(process.cwd(), "assets", "bg-slide.png");
  if (fs.existsSync(bgPath)) {
    slide.background = { path: bgPath };
  } else {
    slide.background = { color: ctx.theme.palette.brand };
  }
}

function createSlide(ctx, options = {}) {
  const slide = ctx.pres.addSlide();
  ctx.slideNumber += 1;
  slide.background = { color: ctx.theme.palette.canvas };

  if (!options.cover) {
    addDeckBackground(slide, ctx);
    const sectionLabel = options.section || "Report";
    addSlideFrame(slide, ctx, sectionLabel);
  }

  return slide;
}

function addSlideFrame(slide, ctx, sectionLabel) {
  const t = ctx.theme;
  slide.addShape(ctx.pres.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: SLIDE.W,
    h: 0.06,
    fill: { color: t.palette.brandSoft },
    line: { color: t.palette.brandSoft, width: 0 },
  });

  slide.addText(sanitizeText(sectionLabel, 28).toUpperCase(), {
    x: SLIDE.M,
    y: 0.1,
    w: 2.8,
    h: 0.18,
    fontFace: t.typography.body,
    fontSize: 8,
    bold: true,
    color: t.palette.brandSoft,
    charSpacing: 1.2,
    margin: 0,
  });

  slide.addText(`${ctx.tenantName} | ${ctx.reportDate}`, {
    x: SLIDE.W - SLIDE.M - 4.5,
    y: 0.1,
    w: 4.5,
    h: 0.18,
    fontFace: t.typography.body,
    fontSize: 8,
    color: t.palette.textSubtle,
    align: "right",
    margin: 0,
  });

  slide.addShape(ctx.pres.shapes.LINE, {
    x: SLIDE.M,
    y: SLIDE.H - 0.28,
    w: SLIDE.W - (SLIDE.M * 2),
    h: 0,
    line: { color: t.palette.border, width: 0.5 },
  });

  slide.addText(`Slide ${ctx.slideNumber}`, {
    x: SLIDE.W - SLIDE.M - 1.5,
    y: SLIDE.H - 0.24,
    w: 1.5,
    h: 0.14,
    fontFace: t.typography.body,
    fontSize: 7,
    color: t.palette.textSubtle,
    align: "right",
    margin: 0,
  });
}

function addSlideTitle(slide, ctx, text, y = 0.34) {
  slide.addText(sanitizeText(text, 98), {
    x: SLIDE.M,
    y,
    w: SLIDE.W - (SLIDE.M * 2),
    h: 0.35,
    fontFace: ctx.theme.typography.title,
    fontSize: 24,
    bold: true,
    color: ctx.theme.palette.textStrong,
    margin: 0,
  });
}

function addSectionSubtitle(slide, ctx, text, y = 0.73) {
  slide.addText(sanitizeText(text, 128), {
    x: SLIDE.M,
    y,
    w: SLIDE.W - (SLIDE.M * 2),
    h: 0.24,
    fontFace: ctx.theme.typography.body,
    fontSize: 11,
    color: ctx.theme.palette.textMuted,
    margin: 0,
  });
}

function addStatBadge(slide, ctx, options) {
  const tone = options.tone || "neutral";
  const colors = getToneColors(ctx.theme, tone);
  const glass = ctx.theme.effects && ctx.theme.effects.glass ? ctx.theme.effects.glass : { transparency: 0, shadow: null };
  slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    fill: { color: colors.soft, transparency: glass.transparency },
    line: { color: colors.soft, width: 0 },
    shadow: glass.shadow,
    rectRadius: ctx.theme.radii.md,
  });
  slide.addText(String(options.value), {
    x: options.x + 0.1,
    y: options.y + 0.06,
    w: options.w - 0.2,
    h: options.h * 0.55,
    fontFace: ctx.theme.typography.heading,
    fontSize: options.valueSize || 20,
    bold: true,
    color: colors.strong,
    align: "center",
    margin: 0,
  });
  slide.addText(sanitizeText(options.label, 32), {
    x: options.x + 0.1,
    y: options.y + (options.h * 0.58),
    w: options.w - 0.2,
    h: options.h * 0.32,
    fontFace: ctx.theme.typography.body,
    fontSize: 9,
    color: ctx.theme.palette.textMuted,
    align: "center",
    margin: 0,
  });
}

function addCalloutBox(slide, ctx, options) {
  const tone = options.tone || "neutral";
  const colors = getToneColors(ctx.theme, tone);
  const glass = ctx.theme.effects && ctx.theme.effects.glass ? ctx.theme.effects.glass : { transparency: 0, shadow: null };
  slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    fill: { color: colors.soft, transparency: glass.transparency },
    line: { color: colors.strong, width: 0 },
    shadow: glass.shadow,
    rectRadius: ctx.theme.radii.md,
  });
  slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
    x: options.x,
    y: options.y,
    w: options.w,
    h: 0.04,
    fill: { color: colors.strong },
    rectRadius: ctx.theme.radii.md,
  });
  if (hasValue(options.title)) {
    slide.addText(sanitizeText(options.title, 72), {
      x: options.x + 0.12,
      y: options.y + 0.07,
      w: options.w - 0.24,
      h: 0.18,
      fontFace: ctx.theme.typography.body,
      fontSize: 9,
      bold: true,
      color: colors.strong,
      margin: 0,
    });
  }
  slide.addText(sanitizeText(options.text, 190), {
    x: options.x + 0.12,
    y: options.y + (hasValue(options.title) ? 0.26 : 0.1),
    w: options.w - 0.24,
    h: options.h - (hasValue(options.title) ? 0.32 : 0.18),
    fontFace: ctx.theme.typography.body,
    fontSize: 10,
    color: ctx.theme.palette.textBody,
    margin: 0,
    valign: "top",
    shrinkText: true,
  });
}

function addInsightCard(slide, ctx, options) {
  const tone = options.tone || "neutral";
  const colors = getToneColors(ctx.theme, tone);
  const glass = ctx.theme.effects && ctx.theme.effects.glass ? ctx.theme.effects.glass : { transparency: 0, shadow: null };
  slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    fill: { color: ctx.theme.palette.surface, transparency: glass.transparency },
    line: { color: ctx.theme.palette.border, width: 0 },
    shadow: glass.shadow,
    rectRadius: ctx.theme.radii.md,
  });
  slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
    x: options.x,
    y: options.y,
    w: 0.06,
    h: options.h,
    fill: { color: colors.strong },
    line: { color: colors.strong, width: 0 },
    rectRadius: ctx.theme.radii.md,
  });

  slide.addText(sanitizeText(options.title, 72), {
    x: options.x + 0.12,
    y: options.y + 0.07,
    w: options.w - 0.24,
    h: 0.22,
    fontFace: ctx.theme.typography.body,
    fontSize: 10,
    bold: true,
    color: colors.strong,
    margin: 0,
  });

  if (hasValue(options.takeaway)) {
    slide.addText(sanitizeText(options.takeaway, 120), {
      x: options.x + 0.12,
      y: options.y + 0.31,
      w: options.w - 0.24,
      h: 0.24,
      fontFace: ctx.theme.typography.body,
      fontSize: 9,
      color: ctx.theme.palette.textBody,
      margin: 0,
      shrinkText: true,
    });
  }

  if (toArray(options.evidence).length) {
    const lines = toArray(options.evidence).slice(0, 5).map((line) => sanitizeText(line, 82));
    slide.addText(lines.map((line) => ({ text: `- ${line}`, options: { breakLine: true } })), {
      x: options.x + 0.12,
      y: options.y + (hasValue(options.takeaway) ? 0.57 : 0.33),
      w: options.w - 0.24,
      h: options.h - (hasValue(options.takeaway) ? 0.65 : 0.41),
      fontFace: ctx.theme.typography.body,
      fontSize: 8,
      color: ctx.theme.palette.textMuted,
      margin: 0,
      shrinkText: true,
    });
  }
}

function addAppendixDivider(slide, ctx, title, subtitle) {
  const bgPath = path.resolve(process.cwd(), "assets", "bg-hero.png");
  if (fs.existsSync(bgPath)) {
    slide.background = { path: bgPath };
    slide.addShape(ctx.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: SLIDE.W, h: SLIDE.H,
      fill: { color: "000000", transparency: 60 },
      line: { type: "none" }
    });
  } else {
    slide.background = { color: ctx.theme.palette.brand };
  }
  slide.addText(sanitizeText(title, 48), {
    x: SLIDE.M,
    y: 2.2,
    w: SLIDE.W - (SLIDE.M * 2),
    h: 0.5,
    fontFace: ctx.theme.typography.title,
    fontSize: 36,
    bold: true,
    color: ctx.theme.palette.textStrong,
    align: "center",
    margin: 0,
  });
  slide.addText(sanitizeText(subtitle, 110), {
    x: SLIDE.M,
    y: 2.78,
    w: SLIDE.W - (SLIDE.M * 2),
    h: 0.3,
    fontFace: ctx.theme.typography.body,
    fontSize: 12,
    color: ctx.theme.palette.textMuted,
    align: "center",
    margin: 0,
  });
}

function estimateTableHeight(rowCount, rowHeight, minHeight, maxHeight) {
  const computed = ((rowCount + 1) * rowHeight) + 0.14;
  return clamp(computed, minHeight, maxHeight);
}

function addTableWrapper(slide, ctx, options) {
  slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    fill: { color: ctx.theme.palette.surface },
    line: { color: ctx.theme.palette.border, width: 0.75 },
    rectRadius: ctx.theme.radii.sm,
  });

  slide.addTable(options.rows, {
    x: options.x + 0.04,
    y: options.y + 0.04,
    w: options.w - 0.08,
    h: options.h - 0.08,
    colW: options.colW,
    rowH: options.rowH || 0.27,
    border: { pt: 0.3, color: ctx.theme.palette.border },
    autoPage: false,
    valign: "middle",
  });
}

function buildTableRows(ctx, headers, rows) {
  const headerRow = headers.map((header) => ({
    text: sanitizeText(header.label, 36),
    options: {
      fill: { color: ctx.theme.palette.brand },
      color: "FFFFFF",
      bold: true,
      fontFace: ctx.theme.typography.body,
      fontSize: 8,
      align: header.align || "left",
    },
  }));

  const bodyRows = rows.map((row, index) => {
    const fillColor = index % 2 === 0 ? ctx.theme.palette.surface : ctx.theme.palette.surfaceAlt;
    return row.map((cell, idx) => {
      const align = headers[idx].align || "left";
      return {
        text: sanitizeText(hasValue(cell) ? cell : "Not configured", 86),
        options: {
          fill: { color: fillColor },
          color: ctx.theme.palette.textBody,
          fontFace: ctx.theme.typography.body,
          fontSize: 7.5,
          align,
        },
      };
    });
  });

  return [headerRow, ...bodyRows];
}

function addNarrativeHeader(slide, ctx, title, headline, takeaway, sectionTop = 0.34) {
  addSlideTitle(slide, ctx, title, sectionTop);
  addSectionSubtitle(slide, ctx, headline, sectionTop + 0.38);
  addCalloutBox(slide, ctx, {
    x: SLIDE.M,
    y: sectionTop + 0.65,
    w: SLIDE.W - (SLIDE.M * 2),
    h: 0.5,
    title: "Takeaway",
    text: takeaway,
    tone: "neutral",
  });
  return sectionTop + 1.23;
}

function addCoverSlide(ctx, analysis) {
  const slide = createSlide(ctx, { cover: true });
  const t = ctx.theme;

  const bgPath = path.resolve(process.cwd(), "assets", "bg-hero.png");
  if (fs.existsSync(bgPath)) {
    slide.background = { path: bgPath };
    slide.addShape(ctx.pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: SLIDE.W, h: SLIDE.H,
      fill: { color: "000000", transparency: 60 },
      line: { type: "none" }
    });
  } else {
    slide.background = { color: t.palette.brand };
  }

  const reportTitle = pickFirst(
    analysis.meta && analysis.meta.reportTitle,
    t.metadata.title,
    "Conditional Access Security Posture Report"
  );
  const coverTitle = sanitizeText(reportTitle, 78);
  const coverTitleLines = Math.min(4, Math.max(2, Math.ceil(coverTitle.length / 18)));
  const coverTitleY = 1.02;
  const coverTitleH = 0.32 + (coverTitleLines * 0.47);
  const coverMetaY = coverTitleY + coverTitleH + 0.12;

  slide.addText(coverTitle, {
    x: SLIDE.M,
    y: coverTitleY,
    w: 6.5,
    h: coverTitleH,
    fontFace: t.typography.title,
    fontSize: 34,
    bold: true,
    color: "FFFFFF",
    margin: 0,
    valign: "top",
    shrinkText: true,
  });

  slide.addText(`${ctx.tenantName} | ${ctx.reportDate}`, {
    x: SLIDE.M,
    y: coverMetaY,
    w: 6.5,
    h: 0.3,
    fontFace: t.typography.body,
    fontSize: 13,
    color: "FFFFFF",
    margin: 0,
  });

  const badges = [
    { label: "Policies", value: ctx.policyCounts.total, tone: "neutral" },
    { label: "Enabled", value: ctx.policyCounts.enabled, tone: "positive" },
    { label: "Report-only", value: ctx.policyCounts.reportOnly, tone: "caution" },
    { label: "Disabled", value: ctx.policyCounts.disabled, tone: "critical" },
  ];
  badges.forEach((badge, index) => {
    const colors = getToneColors(t, badge.tone);
    const x = SLIDE.M + (index * 1.62);
    slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
      x,
      y: 4.35,
      w: 1.48,
      h: 0.9,
      fill: { color: t.palette.surface, transparency: t.effects && t.effects.glass ? t.effects.glass.transparency : 70 },
      line: { color: t.palette.border, width: 0 },
      shadow: t.effects && t.effects.glass ? t.effects.glass.shadow : null,
      rectRadius: t.radii.md,
    });
    slide.addText(String(badge.value), {
      x,
      y: 4.45,
      w: 1.48,
      h: 0.42,
      fontFace: t.typography.heading,
      fontSize: 22,
      bold: true,
      color: colors.strong,
      align: "center",
      margin: 0,
    });
    slide.addText(badge.label, {
      x,
      y: 4.84,
      w: 1.48,
      h: 0.2,
      fontFace: t.typography.body,
      fontSize: 8,
      color: "FFFFFF",
      align: "center",
      margin: 0,
    });
  });

  if (ctx.logoPath) {
    slide.addImage({
      path: ctx.logoPath,
      x: SLIDE.W - 1.95,
      y: 0.3,
      w: 1.5,
      h: 0.55,
      sizing: { type: "contain", x: SLIDE.W - 1.95, y: 0.3, w: 1.5, h: 0.55 },
    });
  }
}

function buildAgendaItems(ctx) {
  const a = ctx.analysis;
  const counts = ctx.policyCounts;
  const assessment = inferAssessment(a, ctx.policies);

  const strengthCount = toArray(a.executiveSummary && a.executiveSummary.strengths).length;
  const concernCount = toArray(a.executiveSummary && a.executiveSummary.concerns).length;
  const priorityCount = toArray(
    a.executiveSummary && a.executiveSummary.topPriorities
      ? a.executiveSummary.topPriorities
      : a.recommendations
  ).length;
  const roCount = counts.reportOnly;

  const evidenceSections = [];
  if (a.policyLandscape) evidenceSections.push("Policy landscape");
  if (a.geolocationStrategy && a.geolocationStrategy.available) evidenceSections.push("Geolocation strategy");
  if (a.mfaMatrix) evidenceSections.push("MFA matrix");
  if (a.riskPolicies) evidenceSections.push("Risk policies");
  if (a.authStrengths && a.authStrengths.available) evidenceSections.push("Auth strengths");
  if (a.pimCoverage && a.pimCoverage.available) evidenceSections.push("PIM coverage");
  if (roCount > 0) evidenceSections.push("Report-only pipeline");

  return [
    {
      title: "1. Posture Snapshot",
      takeaway: `${assessment.level} posture at ${assessment.score}/100. ${counts.enabled} enforced, ${roCount} report-only, ${counts.disabled} disabled.`,
      evidence: [`${strengthCount} strengths identified`, `${concernCount} concerns flagged`, `Score: ${assessment.score}/100`],
      tone: assessment.score >= 80 ? "positive" : assessment.score >= 65 ? "caution" : "critical",
    },
    {
      title: "2. Immediate Priorities",
      takeaway: `${priorityCount} priority actions identified for the next quarter.`,
      evidence: roCount > 0
        ? [`${priorityCount} priority items`, "90-day roadmap", `${roCount} report-only policies to convert`]
        : [`${priorityCount} priority items`, "90-day roadmap"],
      tone: "critical",
    },
    {
      title: "3. Supporting Analysis",
      takeaway: `${evidenceSections.length} evidence sections covering ${counts.total} policies.`,
      evidence: evidenceSections.slice(0, 5),
      tone: "caution",
    },
    {
      title: "4. Appendix",
      takeaway: `Full matrix and per-policy detail for all ${counts.total} policies.`,
      evidence: ["Policy matrix", "Per-policy breakdown"],
      tone: "neutral",
    },
  ];
}

function addAgendaSlide(ctx) {
  const slide = createSlide(ctx, { section: "Executive Narrative" });
  addSlideTitle(slide, ctx, "Agenda");
  addSectionSubtitle(slide, ctx, `${ctx.tenantName} Conditional Access review \u2014 ${ctx.policyCounts.total} policies assessed.`);

  const agendaItems = buildAgendaItems(ctx);

  const cardW = (SLIDE.W - (SLIDE.M * 2) - 0.25) / 2;
  const cardH = 1.35;
  agendaItems.forEach((item, index) => {
    const row = Math.floor(index / 2);
    const col = index % 2;
    addInsightCard(slide, ctx, {
      x: SLIDE.M + (col * (cardW + 0.25)),
      y: 1.3 + (row * (cardH + 0.2)),
      w: cardW,
      h: cardH,
      title: item.title,
      takeaway: item.takeaway,
      evidence: item.evidence,
      tone: item.tone,
    });
  });
}

function addScorecardSlide(ctx, analysis) {
  const slide = createSlide(ctx, { section: "Executive Narrative" });
  const assessment = inferAssessment(analysis, ctx.policies);
  addSlideTitle(slide, ctx, "Posture Scorecard");
  addSectionSubtitle(slide, ctx, `${ctx.policyCounts.enabled}/${ctx.policyCounts.total} policies enforced \u2014 ${ctx.policyCounts.reportOnly} in report-only, ${ctx.policyCounts.disabled} disabled.`);

  const scoreTone = assessment.score >= 80 ? "positive" : assessment.score >= 65 ? "caution" : "critical";
  addStatBadge(slide, ctx, {
    x: SLIDE.M,
    y: 1.2,
    w: 2.4,
    h: 2.0,
    label: `${assessment.level} posture`,
    value: assessment.score,
    valueSize: 44,
    tone: scoreTone,
  });

  const metrics = [
    { label: "Total policies", value: ctx.policyCounts.total, tone: "neutral" },
    { label: "Enabled", value: ctx.policyCounts.enabled, tone: "positive" },
    { label: "Report-only", value: ctx.policyCounts.reportOnly, tone: "caution" },
    { label: "Disabled", value: ctx.policyCounts.disabled, tone: "critical" },
  ];

  metrics.forEach((item, index) => {
    const col = index % 2;
    const row = Math.floor(index / 2);
    addStatBadge(slide, ctx, {
      x: 3.1 + (col * 3.25),
      y: 1.2 + (row * 1.05),
      w: 3.0,
      h: 0.9,
      label: item.label,
      value: item.value,
      valueSize: 26,
      tone: item.tone,
    });
  });

  addCalloutBox(slide, ctx, {
    x: SLIDE.M,
    y: 3.45,
    w: SLIDE.W - (SLIDE.M * 2),
    h: 1.05,
    title: "Assessment",
    text: `${assessment.verdict}. ${assessment.prioritySummary}${assessment.criticalGap ? ` Critical gap: ${assessment.criticalGap}.` : ""}`,
    tone: scoreTone,
  });
}

function addExecutiveSummarySlide(ctx, analysis) {
  const summary = analysis.executiveSummary || {};
  const strengths = toArray(summary.strengths).slice(0, 6);
  const concerns = toArray(summary.concerns).slice(0, 6);

  const slide = createSlide(ctx, { section: "Executive Narrative" });
  addSlideTitle(slide, ctx, "Executive Summary");
  addSectionSubtitle(slide, ctx, `${strengths.length} strengths and ${concerns.length} concerns identified across ${ctx.policyCounts.total} policies.`);

  addInsightCard(slide, ctx, {
    x: SLIDE.M,
    y: 1.2,
    w: (SLIDE.W - (SLIDE.M * 2) - 0.2) / 2,
    h: 3.0,
    title: "Strengths",
    takeaway: strengths.length ? sanitizeText(strengths[0], 98) : "No strengths were provided in analysis.",
    evidence: strengths,
    tone: "positive",
  });

  addInsightCard(slide, ctx, {
    x: SLIDE.M + ((SLIDE.W - (SLIDE.M * 2) - 0.2) / 2) + 0.2,
    y: 1.2,
    w: (SLIDE.W - (SLIDE.M * 2) - 0.2) / 2,
    h: 3.0,
    title: "Concerns",
    takeaway: concerns.length ? sanitizeText(concerns[0], 98) : "No concerns were provided in analysis.",
    evidence: concerns,
    tone: "critical",
  });
}

function addTopPrioritiesSlides(ctx, analysis) {
  const priorities = collectTopPriorities(analysis);
  if (!priorities.length) return;
  const pages = chunk(priorities, CONTENT_LIMITS.executiveCardsPerSlide);

  pages.forEach((page, pageIndex) => {
    const slide = createSlide(ctx, { section: "Executive Narrative" });
    addSlideTitle(slide, ctx, pageIndex === 0 ? "Top Priorities" : `Top Priorities (${pageIndex + 1}/${pages.length})`);
    addSectionSubtitle(
      slide,
      ctx,
      `${priorities.length} action${priorities.length === 1 ? "" : "s"} ranked by risk reduction impact for ${ctx.tenantName}.`
    );

    page.forEach((item, index) => {
      addInsightCard(slide, ctx, {
        x: SLIDE.M,
        y: 1.2 + (index * 0.9),
        w: SLIDE.W - (SLIDE.M * 2),
        h: 0.8,
        title: item.title,
        takeaway: item.evidence || `Priority: ${item.priority}`,
        evidence: [],
        tone: getSeverityTone(item.priority),
      });
    });
  });
}

function addRoadmapSlides(ctx, analysis) {
  const roadmap = buildRoadmap(analysis);
  if (!roadmap.nearTerm.length && !roadmap.midTerm.length) return;

  const nearPages = chunk(roadmap.nearTerm, CONTENT_LIMITS.roadmapItemsPerColumn);
  const midPages = chunk(roadmap.midTerm, CONTENT_LIMITS.roadmapItemsPerColumn);
  const pageCount = Math.max(nearPages.length, midPages.length);

  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    const slide = createSlide(ctx, { section: "Executive Narrative" });
    addSlideTitle(
      slide,
      ctx,
      pageCount === 1 ? "90-Day Roadmap" : `90-Day Roadmap (${pageIndex + 1}/${pageCount})`
    );
    addSectionSubtitle(
      slide,
      ctx,
      `${roadmap.nearTerm.length} near-term and ${roadmap.midTerm.length} mid-term actions through ${pickFirst(analysis.meta && analysis.meta.nextReview, "next quarter")}.`
    );

    addInsightCard(slide, ctx, {
      x: SLIDE.M,
      y: 1.2,
      w: (SLIDE.W - (SLIDE.M * 2) - 0.25) / 2,
      h: 3.3,
      title: "Near-term actions (0-45 days)",
      takeaway: "Execute first to reduce active exposure and remove stale report-only controls.",
      evidence: (nearPages[pageIndex] || []).map((item) => item.title),
      tone: "critical",
    });

    addInsightCard(slide, ctx, {
      x: SLIDE.M + ((SLIDE.W - (SLIDE.M * 2) - 0.25) / 2) + 0.25,
      y: 1.2,
      w: (SLIDE.W - (SLIDE.M * 2) - 0.25) / 2,
      h: 3.3,
      title: "Mid-term actions (45-90 days)",
      takeaway: "Use these actions to harden role coverage and simplify policy operations.",
      evidence: (midPages[pageIndex] || []).map((item) => item.title),
      tone: "caution",
    });
  }
}

function addEvidenceDividerSlide(ctx) {
  const slide = createSlide(ctx, { section: "Supporting Analysis" });
  addAppendixDivider(slide, ctx, "Supporting Analysis", "Control evidence and policy-level detail that supports the executive narrative.");
}

function addPolicyLandscapeSlides(ctx, analysis) {
  const landscape = analysis.policyLandscape || {};
  const categories = toArray(landscape.categories);
  const categoryPages = chunk(categories, CONTENT_LIMITS.categoryCardsPerSlide);
  const policyCount = Math.max(1, ctx.policyCounts.total);
  const enabledCount = ctx.policyCounts.enabled;
  const reportOnlyCount = ctx.policyCounts.reportOnly;
  const disabledCount = ctx.policyCounts.disabled;

  categoryPages.forEach((page, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const headline = `${enabledCount}/${policyCount} policies are enforced. ${reportOnlyCount} remain in report-only mode.`;
    const takeaway = reportOnlyCount > 0
      ? "Prioritize report-only conversions in parallel with high-severity gaps."
      : "Report-only backlog is clear; focus can shift to control tuning.";
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pageIndex === 0 ? "Policy Landscape" : `Policy Landscape (${pageIndex + 1}/${categoryPages.length})`,
      headline,
      takeaway
    );

    const barData = [
      { label: "Enabled", value: enabledCount, tone: "positive" },
      { label: "Report-only", value: reportOnlyCount, tone: "caution" },
      { label: "Disabled", value: disabledCount, tone: "critical" },
    ];

    barData.forEach((item, index) => {
      const y = startY + (index * 0.37);
      const width = clamp((item.value / Math.max(1, policyCount)) * 4.5, 0.3, 4.5);
      const tone = getToneColors(ctx.theme, item.tone);
      slide.addText(item.label, {
        x: SLIDE.M,
        y,
        w: 1.3,
        h: 0.2,
        fontFace: ctx.theme.typography.body,
        fontSize: 9,
        color: ctx.theme.palette.textMuted,
        margin: 0,
      });
      slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
        x: SLIDE.M + 1.35,
        y: y + 0.02,
        w: width,
        h: 0.17,
        fill: { color: tone.strong },
        line: { color: tone.strong, width: 0.1 },
        rectRadius: ctx.theme.radii.sm,
      });
      slide.addText(String(item.value), {
        x: SLIDE.M + 1.35 + width + 0.08,
        y,
        w: 0.8,
        h: 0.2,
        fontFace: ctx.theme.typography.body,
        fontSize: 9,
        bold: true,
        color: tone.strong,
        margin: 0,
      });
    });

    const cardY = startY + 1.25;
    const cardW = (SLIDE.W - (SLIDE.M * 2) - 0.25) / 3;
    const cardH = 1.2;
    page.forEach((category, index) => {
      const row = Math.floor(index / 3);
      const col = index % 3;
      addInsightCard(slide, ctx, {
        x: SLIDE.M + (col * (cardW + 0.125)),
        y: cardY + (row * (cardH + 0.15)),
        w: cardW,
        h: cardH,
        title: sanitizeText(category.label, 36).replace(/\n/g, " "),
        takeaway: `${category.count || 0} policies`,
        evidence: [],
        tone: "neutral",
      });
    });
  });
}

function addGeolocationSlides(ctx, analysis) {
  const geo = analysis.geolocationStrategy;
  if (!geo || !geo.available) return;
  const layers = toArray(geo.layers);
  if (!layers.length) return;
  const pages = chunk(layers, CONTENT_LIMITS.layerCardsPerSlide);

  pages.forEach((page, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pageIndex === 0 ? "Geolocation Strategy" : `Geolocation Strategy (${pageIndex + 1}/${pages.length})`,
      `${layers.length} geolocation layers are active in the tenant posture.`,
      sanitizeText(pickFirst(geo.note, "Layered controls should align with named locations and report-only exits."), 140)
    );

    const cardW = SLIDE.W - (SLIDE.M * 2);
    const cardH = 1.05;
    page.forEach((layer, index) => {
      const evidence = [];
      if (hasValue(layer.description)) evidence.push(layer.description);
      if (toArray(layer.countries).length) {
        const countrySummary = toArray(layer.countries)
          .slice(0, 4)
          .map((country) => `${sanitizeText(country.name, 24)} (${sanitizeText(country.state, 12)})`)
          .join(", ");
        evidence.push(countrySummary);
      }
      if (hasValue(layer.locationCount)) evidence.push(`${layer.locationCount} trusted locations`);

      addInsightCard(slide, ctx, {
        x: SLIDE.M,
        y: startY + (index * (cardH + 0.16)),
        w: cardW,
        h: cardH,
        title: `Layer ${sanitizeText(layer.id, 4)} - ${sanitizeText(layer.title, 70)}`,
        takeaway: evidence[0] || "No additional detail provided.",
        evidence: evidence.slice(1),
        tone: "neutral",
      });
    });
  });
}

function addMfaMatrixSlides(ctx, analysis) {
  const mfa = analysis.mfaMatrix;
  if (!mfa || !mfa.available) return;
  const rows = toArray(mfa.policies).map((item) => [
    sanitizeText(item.name, 52),
    sanitizeText(item.scope, 42),
    sanitizeText(item.authStrength, 36),
    sanitizeText(item.frequency, 22),
    sanitizeText(item.conditions, 44),
  ]);
  if (!rows.length) return;

  const pages = chunk(rows, CONTENT_LIMITS.tableRows.mfa);
  pages.forEach((pageRows, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pages.length === 1 ? "MFA Enforcement Matrix" : `MFA Enforcement Matrix (${pageIndex + 1}/${pages.length})`,
      `${rows.length} MFA policy entries were analyzed for scope and control strength.`,
      "Confirm report-only MFA policies have transition plans to enforcement."
    );

    const headers = [
      { label: "Policy" },
      { label: "Scope" },
      { label: "Auth strength" },
      { label: "Frequency", align: "center" },
      { label: "Conditions" },
    ];
    const formatted = buildTableRows(ctx, headers, pageRows);
    const tableRowH = 0.28;
    const tableH = estimateTableHeight(pageRows.length, tableRowH, 1.2, 2.95);
    addTableWrapper(slide, ctx, {
      x: SLIDE.M,
      y: startY,
      w: SLIDE.W - (SLIDE.M * 2),
      h: tableH,
      rows: formatted,
      colW: [2.2, 1.7, 1.7, 1.0, 2.4],
      rowH: tableRowH,
    });

    if (pageIndex === 0 && mfa.callout) {
      const calloutY = Math.min(startY + tableH + 0.14, 4.28);
      addCalloutBox(slide, ctx, {
        x: SLIDE.M,
        y: calloutY,
        w: SLIDE.W - (SLIDE.M * 2),
        h: 0.72,
        title: sanitizeText(mfa.callout.title, 52),
        text: sanitizeText(mfa.callout.text, 160),
        tone: "caution",
      });
    }
  });
}

function addRiskPolicySlides(ctx, analysis) {
  const risk = analysis.riskPolicies;
  if (!risk || !risk.available) return;
  const cards = toArray(risk.policies);
  if (!cards.length) return;
  const pages = chunk(cards, CONTENT_LIMITS.riskCardsPerSlide);

  pages.forEach((page, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pages.length === 1 ? "Identity Risk Policies" : `Identity Risk Policies (${pageIndex + 1}/${pages.length})`,
      `${cards.length} risk-focused policies were mapped to user and sign-in conditions.`,
      "High-risk users and high-risk sign-ins should have explicit enforcement and recovery paths."
    );

    const gap = 0.18;
    const cardW = (SLIDE.W - (SLIDE.M * 2) - (gap * (page.length - 1))) / Math.max(1, page.length);
    page.forEach((item, index) => {
      const evidence = [
        `Grant control: ${sanitizeText(item.grantControl, 44)}`,
        `Operator: ${sanitizeText(item.operator, 28)}`,
        `Scope: ${sanitizeText(item.scope, 44)}`,
        `Modified: ${sanitizeText(item.modified, 20)}`,
      ];
      const tone = getSeverityTone(item.title);
      addInsightCard(slide, ctx, {
        x: SLIDE.M + (index * (cardW + gap)),
        y: startY,
        w: cardW,
        h: 2.95,
        title: sanitizeText(item.title, 38),
        takeaway: sanitizeText(item.policyName, 60),
        evidence,
        tone,
      });
    });

    if (pageIndex === 0 && risk.callout) {
      addCalloutBox(slide, ctx, {
        x: SLIDE.M,
        y: 4.45,
        w: SLIDE.W - (SLIDE.M * 2),
        h: 0.65,
        title: "Risk takeaway",
        text: sanitizeText(risk.callout.text, 160),
        tone: getSeverityTone(risk.callout.color),
      });
    }
  });
}

function addAuthStrengthSlides(ctx, analysis) {
  const auth = analysis.authStrengths;
  if (!auth || !auth.available) return;
  const strengths = toArray(auth.strengths);
  if (!strengths.length) return;
  const pages = chunk(strengths, CONTENT_LIMITS.authCardsPerSlide);

  pages.forEach((page, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pages.length === 1 ? "Authentication Strengths" : `Authentication Strengths (${pageIndex + 1}/${pages.length})`,
      `${strengths.length} authentication strengths were evaluated for resistance and deployability.`,
      "Remove phishable factors from admin-facing strengths and map each strength to policy scope."
    );

    const cardW = (SLIDE.W - (SLIDE.M * 2) - 0.18) / 2;
    const cardH = 1.45;
    page.forEach((item, index) => {
      const row = Math.floor(index / 2);
      const col = index % 2;
      const tone = getSeverityTone(item.rating);
      addInsightCard(slide, ctx, {
        x: SLIDE.M + (col * (cardW + 0.18)),
        y: startY + (row * (cardH + 0.16)),
        w: cardW,
        h: cardH,
        title: `${sanitizeText(item.name, 52)} (${sanitizeText(item.rating, 16)})`,
        takeaway: `Type: ${sanitizeText(item.type, 20)} | Methods: ${sanitizeText(item.methods, 55)}`,
        evidence: hasValue(item.note) ? [item.note] : [],
        tone,
      });
    });
  });
}

function addPimCoverageSlides(ctx, analysis) {
  const pim = analysis.pimCoverage;
  if (!pim || !pim.available) return;
  const rows = toArray(pim.roles).map((item) => [
    sanitizeText(item.role, 44),
    sanitizeText(item.eligible, 10),
    sanitizeText(item.active, 10),
    sanitizeText(item.caCoverage, 52),
    sanitizeText(item.direct, 12),
  ]);
  if (!rows.length) return;
  const pages = chunk(rows, CONTENT_LIMITS.tableRows.pim);

  pages.forEach((pageRows, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pages.length === 1 ? "Privileged Access Coverage" : `Privileged Access Coverage (${pageIndex + 1}/${pages.length})`,
      `${rows.length} privileged roles were compared against Conditional Access targeting.`,
      "Roles without direct policy coverage should be prioritized in roadmap planning."
    );

    const headers = [
      { label: "PIM role" },
      { label: "Eligible", align: "center" },
      { label: "Active", align: "center" },
      { label: "CA coverage" },
      { label: "Direct", align: "center" },
    ];
    const tableRowH = 0.27;
    const tableH = estimateTableHeight(pageRows.length, tableRowH, 1.2, 2.95);

    addTableWrapper(slide, ctx, {
      x: SLIDE.M,
      y: startY,
      w: SLIDE.W - (SLIDE.M * 2),
      h: tableH,
      rows: buildTableRows(ctx, headers, pageRows),
      colW: [2.2, 0.75, 0.75, 4.35, 1.0],
      rowH: tableRowH,
    });

    if (pageIndex === 0 && hasValue(pim.note)) {
      const calloutY = Math.min(startY + tableH + 0.14, 4.28);
      addCalloutBox(slide, ctx, {
        x: SLIDE.M,
        y: calloutY,
        w: SLIDE.W - (SLIDE.M * 2),
        h: 0.72,
        title: "Coverage note",
        text: sanitizeText(pim.note, 170),
        tone: "neutral",
      });
    }
  });
}

function addReportOnlyPipelineSlides(ctx, analysis) {
  const pipeline = analysis.reportOnlyPipeline;
  const policies = pipeline ? toArray(pipeline.policies) : [];
  if (!policies.length) return;
  const rows = policies.map((item) => [
    sanitizeText(item.name, 50),
    sanitizeText(item.grantControl, 34),
    sanitizeText(item.targetApps, 36),
    sanitizeText(item.modified, 20),
    sanitizeText(item.priority, 10),
  ]);
  const pages = chunk(rows, CONTENT_LIMITS.tableRows.reportOnly);

  pages.forEach((pageRows, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pages.length === 1 ? "Report-only Pipeline" : `Report-only Pipeline (${pageIndex + 1}/${pages.length})`,
      `${rows.length} report-only entries require staged conversion plans.`,
      "Use priority and staleness to define conversion waves."
    );

    const headers = [
      { label: "Policy" },
      { label: "Grant control" },
      { label: "Target apps" },
      { label: "Modified", align: "center" },
      { label: "Priority", align: "center" },
    ];
    const tableRowH = 0.27;
    const tableH = estimateTableHeight(pageRows.length, tableRowH, 1.2, 2.9);
    addTableWrapper(slide, ctx, {
      x: SLIDE.M,
      y: startY,
      w: SLIDE.W - (SLIDE.M * 2),
      h: tableH,
      rows: buildTableRows(ctx, headers, pageRows),
      colW: [2.45, 1.6, 2.2, 1.2, 0.9],
      rowH: tableRowH,
    });

    if (pageIndex === 0 && pipeline.callout) {
      const calloutY = Math.min(startY + tableH + 0.14, 4.28);
      addCalloutBox(slide, ctx, {
        x: SLIDE.M,
        y: calloutY,
        w: SLIDE.W - (SLIDE.M * 2),
        h: 0.72,
        title: sanitizeText(pipeline.callout.title, 64),
        text: sanitizeText(pipeline.callout.text, 170),
        tone: "critical",
      });
    }
  });
}

function addMsManagedOverlapSlides(ctx, analysis) {
  const overlap = analysis.msManagedOverlap;
  const policies = overlap ? toArray(overlap.policies) : [];
  if (!policies.length) return;
  const rows = policies.map((item) => [
    sanitizeText(item.name, 46),
    sanitizeText(item.description, 56),
    sanitizeText(item.customEquivalent, 46),
    sanitizeText(item.overlap, 12),
  ]);
  const pages = chunk(rows, CONTENT_LIMITS.tableRows.msManaged);

  pages.forEach((pageRows, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pages.length === 1 ? "Microsoft-managed Overlap" : `Microsoft-managed Overlap (${pageIndex + 1}/${pages.length})`,
      `${rows.length} auto-created policies were mapped against custom equivalents.`,
      "Any overlap marked as Gap should move into the high-priority backlog."
    );

    const headers = [
      { label: "MS-managed policy" },
      { label: "What it does" },
      { label: "Custom equivalent" },
      { label: "Overlap", align: "center" },
    ];
    const tableRowH = 0.28;
    const tableH = estimateTableHeight(pageRows.length, tableRowH, 1.2, 2.9);
    addTableWrapper(slide, ctx, {
      x: SLIDE.M,
      y: startY,
      w: SLIDE.W - (SLIDE.M * 2),
      h: tableH,
      rows: buildTableRows(ctx, headers, pageRows),
      colW: [2.55, 2.7, 2.75, 1.1],
      rowH: tableRowH,
    });

    if (pageIndex === 0 && overlap.callout) {
      const calloutY = Math.min(startY + tableH + 0.14, 4.28);
      addCalloutBox(slide, ctx, {
        x: SLIDE.M,
        y: calloutY,
        w: SLIDE.W - (SLIDE.M * 2),
        h: 0.72,
        title: sanitizeText(overlap.callout.title, 64),
        text: sanitizeText(overlap.callout.text, 170),
        tone: "critical",
      });
    }
  });
}

function addRecommendationsSlides(ctx, analysis) {
  const recommendations = analysis.recommendations || {};
  const rows = [];
  toArray(recommendations.high).forEach((item) => rows.push({ priority: "High", text: sanitizeText(item, 94) }));
  toArray(recommendations.medium).forEach((item) => rows.push({ priority: "Medium", text: sanitizeText(item, 94) }));
  toArray(recommendations.low).forEach((item) => rows.push({ priority: "Low", text: sanitizeText(item, 94) }));
  if (!rows.length) return;

  const pages = chunk(rows, 6);
  pages.forEach((page, pageIndex) => {
    const slide = createSlide(ctx, { section: "Supporting Analysis" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pages.length === 1 ? "Recommendation Inventory" : `Recommendation Inventory (${pageIndex + 1}/${pages.length})`,
      `${rows.length} recommendation items were grouped by severity for implementation planning.`,
      "Drive near-term delivery from the High group and keep Medium on a dated transition plan."
    );

    page.forEach((item, index) => {
      const y = startY + (index * 0.58);
      const tone = getSeverityTone(item.priority);
      addInsightCard(slide, ctx, {
        x: SLIDE.M,
        y,
        w: SLIDE.W - (SLIDE.M * 2),
        h: 0.54,
        title: `${item.priority} priority`,
        takeaway: item.text,
        evidence: [],
        tone,
      });
    });
  });
}

function addAppendixDividerSlide(ctx, title, subtitle) {
  const slide = createSlide(ctx, { section: "Appendix" });
  addAppendixDivider(slide, ctx, title, subtitle);
}

function buildMatrixRows(policies) {
  return policies.map((policy, index) => {
    const scopeSummary = summarizeList(
      [
        policy.users.includeUsers !== "Not configured" ? `Users: ${policy.users.includeUsers}` : null,
        policy.users.includeGroups !== "Not configured" ? `Groups: ${policy.users.includeGroups}` : null,
        policy.applications.include !== "Not configured" ? `Apps: ${policy.applications.include}` : null,
      ].filter(Boolean),
      { maxItems: 2, maxLen: 56 }
    );
    return [
      String(index + 1),
      policy.name,
      mapStateLabel(policy.state),
      policy.action,
      scopeSummary,
      policy.lastModified,
    ];
  });
}

function addPolicyMatrixSlides(ctx) {
  const matrixRows = buildMatrixRows(ctx.policies);
  const pages = chunk(matrixRows, CONTENT_LIMITS.tableRows.matrix);

  pages.forEach((pageRows, pageIndex) => {
    const slide = createSlide(ctx, { section: "Appendix" });
    const startY = addNarrativeHeader(
      slide,
      ctx,
      pages.length === 1 ? "Full Policy Matrix" : `Full Policy Matrix (${pageIndex + 1}/${pages.length})`,
      `${ctx.policies.length} policies are listed with state, action, scope summary, and modified date.`,
      "This appendix inventory is optimized for export and audit reference."
    );

    const headers = [
      { label: "#", align: "center" },
      { label: "Policy" },
      { label: "State", align: "center" },
      { label: "Action", align: "center" },
      { label: "Scope summary" },
      { label: "Modified", align: "center" },
    ];
    const tableRowH = 0.24;
    const tableH = estimateTableHeight(pageRows.length, tableRowH, 2.2, 3.5);
    addTableWrapper(slide, ctx, {
      x: SLIDE.M,
      y: startY,
      w: SLIDE.W - (SLIDE.M * 2),
      h: tableH,
      rows: buildTableRows(ctx, headers, pageRows),
      colW: [0.35, 2.7, 0.95, 0.75, 3.3, 1.1],
      rowH: tableRowH,
    });
  });
}

function collectActiveGrantControls(policy) {
  return Array.from(new Set(
    toArray(policy.grantControls.controls)
      .map((value) => {
        const normalized = String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
        if (!normalized || normalized === "notconfigured") return null;
        if (normalized.includes("block")) return "Block access";
        if (normalized.includes("mfa") || normalized.includes("multifactor")) return "Multifactor authentication";
        if (normalized.includes("authenticationstrength") || normalized.includes("authstrength")) return "Authentication strength";
        if (normalized.includes("compliantdevice")) return "Compliant device";
        if (normalized.includes("hybridazureadjoined")) return "Hybrid Azure AD joined";
        if (normalized.includes("approvedclientapp")) return "Approved client app";
        if (normalized.includes("appprotectionpolicy")) return "App protection policy";
        if (normalized.includes("changepassword") || normalized.includes("passwordchange")) return "Change password";
        if (normalized.includes("termsofuse")) return "Terms of use";
        return sanitizeText(value, 36);
      })
      .filter(Boolean)
      .concat(policy.grantControls.authStrength ? ["Authentication strength"] : [])
  ));
}

function collectActiveSessionControls(policy) {
  return Array.from(new Set(
    toArray(policy.sessionControls.active)
      .map((value) => {
        const normalized = String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
        if (!normalized || normalized === "notconfigured") return null;
        if (normalized.includes("appenforcedrestrictions")) return "App enforced restrictions";
        if (normalized.includes("conditionalaccessappcontrol") || normalized.includes("cloudappsecurity")) {
          return "Conditional Access App Control";
        }
        if (normalized.includes("signinfrequency")) return "Sign-in frequency";
        if (normalized.includes("persistentbrowser")) return "Persistent browser session";
        if (normalized.includes("continuousaccessevaluation")) return "Continuous access evaluation";
        if (normalized.includes("disableresiliencedefaults")) return "Disable resilience defaults";
        if (normalized.includes("tokenprotection")) return "Token protection";
        return sanitizeText(value, 36);
      })
      .filter(Boolean)
  ));
}

function buildPolicyUsersLines(policy) {
  const lines = [];
  if (isConfiguredValue(policy.users.includeUsers)) lines.push({ text: `Include users: ${policy.users.includeUsers}`, tone: "positive" });
  if (isConfiguredValue(policy.users.includeGroups)) lines.push({ text: `Include groups: ${policy.users.includeGroups}`, tone: "positive" });
  if (isConfiguredValue(policy.users.includeRoles)) lines.push({ text: `Include roles: ${policy.users.includeRoles}`, tone: "positive" });
  if (isConfiguredValue(policy.users.excludeUsers)) lines.push({ text: `Exclude users: ${policy.users.excludeUsers}`, tone: "critical" });
  if (isConfiguredValue(policy.users.excludeGroups)) lines.push({ text: `Exclude groups: ${policy.users.excludeGroups}`, tone: "critical" });
  if (isConfiguredValue(policy.users.excludeRoles)) lines.push({ text: `Exclude roles: ${policy.users.excludeRoles}`, tone: "critical" });
  if (!lines.length) lines.push({ text: "Not configured", tone: "neutral" });
  return lines;
}

function buildPolicyAppLines(policy) {
  const lines = [];
  if (isConfiguredValue(policy.applications.include)) lines.push({ text: `Include: ${policy.applications.include}`, tone: "positive" });
  if (isConfiguredValue(policy.applications.userActions)) lines.push({ text: `Actions: ${policy.applications.userActions}`, tone: "neutral" });
  if (isConfiguredValue(policy.applications.authContext)) lines.push({ text: `Auth context: ${policy.applications.authContext}`, tone: "neutral" });
  if (isConfiguredValue(policy.applications.exclude)) lines.push({ text: `Exclude: ${policy.applications.exclude}`, tone: "critical" });
  if (!lines.length) lines.push({ text: "Not configured", tone: "neutral" });
  return lines;
}

function addPolicyDetailPanel(slide, ctx, options) {
  const tone = options.tone || "neutral";
  const colors = getToneColors(ctx.theme, tone);
  slide.addShape(ctx.pres.shapes.RECTANGLE, {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    fill: { color: ctx.theme.palette.surface },
    line: { color: ctx.theme.palette.border, width: 0.6 },
  });
  slide.addShape(ctx.pres.shapes.RECTANGLE, {
    x: options.x,
    y: options.y,
    w: 0.03,
    h: options.h,
    fill: { color: colors.strong },
    line: { color: colors.strong, width: 0 },
  });
  slide.addText(sanitizeText(options.title, 38).toUpperCase(), {
    x: options.x + 0.1,
    y: options.y + 0.07,
    w: options.w - 0.16,
    h: 0.2,
    fontFace: ctx.theme.typography.body,
    fontSize: 9,
    bold: true,
    charSpacing: 1.4,
    color: ctx.theme.palette.textMuted,
    margin: 0,
  });

  if (hasValue(options.subtitle)) {
    slide.addText(sanitizeText(options.subtitle, 78), {
      x: options.x + 0.1,
      y: options.y + 0.28,
      w: options.w - 0.16,
      h: 0.2,
      fontFace: ctx.theme.typography.body,
      fontSize: 9,
      bold: true,
      color: ctx.theme.palette.textBody,
      margin: 0,
    });
  }
}

function addPolicyDetailLines(slide, ctx, options) {
  const rowStart = options.y + (hasValue(options.subtitle) ? 0.54 : 0.3);
  const rowH = 0.23;
  const maxRows = Math.max(1, Math.floor((options.h - (rowStart - options.y) - 0.08) / rowH));
  const lines = options.lines.slice(0, maxRows);

  lines.forEach((line, index) => {
    const tone = line.tone || "neutral";
    const colors = getToneColors(ctx.theme, tone);
    slide.addText("- ", {
      x: options.x + 0.1,
      y: rowStart + (index * rowH),
      w: 0.12,
      h: 0.2,
      fontFace: ctx.theme.typography.body,
      fontSize: 9,
      bold: true,
      color: tone === "neutral" ? ctx.theme.palette.textSubtle : colors.strong,
      margin: 0,
    });
    slide.addText(sanitizeText(line.text, 80), {
      x: options.x + 0.2,
      y: rowStart + (index * rowH),
      w: options.w - 0.28,
      h: 0.2,
      fontFace: ctx.theme.typography.body,
      fontSize: 9,
      color: tone === "neutral" ? ctx.theme.palette.textMuted : ctx.theme.palette.textBody,
      margin: 0,
      shrinkText: true,
    });
  });
}

function formatLocationChipValue(policy) {
  if (isConfiguredValue(policy.conditions.locationsInclude) && isConfiguredValue(policy.conditions.locationsExclude)) {
    return `${policy.conditions.locationsInclude} | excl: ${policy.conditions.locationsExclude}`;
  }
  if (isConfiguredValue(policy.conditions.locationsInclude)) return policy.conditions.locationsInclude;
  if (isConfiguredValue(policy.conditions.locationsExclude)) return `Exclude: ${policy.conditions.locationsExclude}`;
  return "Not configured";
}

function addPolicyConditionRow(slide, ctx, policy, y) {
  const chips = [
    { label: "Platforms", value: policy.conditions.platforms },
    { label: "Locations", value: formatLocationChipValue(policy) },
    { label: "Client Apps", value: policy.conditions.clientApps },
    { label: "Risk", value: policy.conditions.risk },
    { label: "Devices", value: policy.conditions.devices },
  ];
  const gap = 0.1;
  const chipW = (SLIDE.W - (SLIDE.M * 2) - (gap * (chips.length - 1))) / chips.length;
  const chipH = 0.44;

  slide.addText("CONDITIONS", {
    x: SLIDE.M,
    y: y - 0.18,
    w: 1.9,
    h: 0.16,
    fontFace: ctx.theme.typography.body,
    fontSize: 8,
    bold: true,
    charSpacing: 2,
    color: ctx.theme.palette.textSubtle,
    margin: 0,
  });

  chips.forEach((chip, index) => {
    const x = SLIDE.M + (index * (chipW + gap));
    const active = isConfiguredValue(chip.value);
    slide.addShape(ctx.pres.shapes.RECTANGLE, {
      x,
      y,
      w: chipW,
      h: chipH,
      fill: { color: ctx.theme.palette.surfaceAlt },
      line: { color: active ? ctx.theme.palette.brandSoft : ctx.theme.palette.border, width: 0.55 },
    });
    if (active) {
      slide.addShape(ctx.pres.shapes.RECTANGLE, {
        x,
        y,
        w: 0.025,
        h: chipH,
        fill: { color: ctx.theme.palette.brandSoft },
        line: { color: ctx.theme.palette.brandSoft, width: 0 },
      });
    }
    slide.addText(chip.label, {
      x: x + 0.08,
      y: y + 0.06,
      w: chipW - 0.12,
      h: 0.16,
      fontFace: ctx.theme.typography.body,
      fontSize: 8,
      bold: true,
      color: ctx.theme.palette.textBody,
      margin: 0,
    });
    slide.addText(sanitizeText(chip.value, 46), {
      x: x + 0.08,
      y: y + 0.24,
      w: chipW - 0.12,
      h: 0.16,
      fontFace: ctx.theme.typography.body,
      fontSize: 7.5,
      color: active ? ctx.theme.palette.textMuted : ctx.theme.palette.textSubtle,
      margin: 0,
      shrinkText: true,
    });
  });
}

function addControlChecklist(slide, ctx, options) {
  const controls = Array.isArray(options.controls) ? options.controls : [];
  const compact = Boolean(options.compact);
  const baseRowH = options.rowH || (compact ? 0.14 : 0.175);
  const minRowH = compact ? 0.11 : 0.13;
  const markerSize = compact ? 7 : 8;
  const textSize = compact ? 8 : 8;
  let subtitle = hasValue(options.subtitle) ? options.subtitle : null;
  let rowStart = options.y + (hasValue(subtitle) ? 0.54 : 0.32);
  let available = Math.max(0.08, options.h - (rowStart - options.y) - 0.08);
  let rowH = controls.length ? Math.min(baseRowH, available / controls.length) : baseRowH;

  if (controls.length && rowH < minRowH && hasValue(subtitle)) {
    subtitle = null;
    rowStart = options.y + 0.32;
    available = Math.max(0.08, options.h - (rowStart - options.y) - 0.08);
    rowH = Math.min(baseRowH, available / controls.length);
  }

  const textH = clamp(rowH, 0.1, 0.18);

  addPolicyDetailPanel(slide, ctx, {
    ...options,
    subtitle,
  });

  controls.forEach((control, index) => {
    const active = options.activeSet.has(control);
    slide.addText(`[${active ? "x" : " "}]`, {
      x: options.x + 0.1,
      y: rowStart + (index * rowH),
      w: 0.2,
      h: textH,
      fontFace: ctx.theme.typography.body,
      fontSize: markerSize,
      bold: true,
      color: active ? ctx.theme.palette.brandSoft : ctx.theme.palette.textSubtle,
      margin: 0,
    });
    slide.addText(sanitizeText(control, 36), {
      x: options.x + 0.3,
      y: rowStart + (index * rowH),
      w: options.w - 0.38,
      h: textH,
      fontFace: ctx.theme.typography.body,
      fontSize: textSize,
      bold: false,
      color: active ? ctx.theme.palette.textBody : ctx.theme.palette.textSubtle,
      margin: 0,
    });
  });
}

function addPolicyDetailSlides(ctx) {
  const sorted = [
    ...ctx.policies.filter((policy) => policy.state === "enabled"),
    ...ctx.policies.filter((policy) => policy.state === "report_only"),
    ...ctx.policies.filter((policy) => policy.state === "disabled"),
  ];

  sorted.forEach((policy, index) => {
    const slide = createSlide(ctx, { section: "Appendix" });
    const shortName = sanitizeText(policy.name.replace(/^Microsoft-managed:\s*/i, ""), 110);
    const stateTone = mapStateTone(policy.state);
    const actionTone = policy.action === "Block" ? "critical" : "positive";
    const actionColors = getToneColors(ctx.theme, actionTone);
    const stateLabel = mapStateLabel(policy.state).toUpperCase();
    const titleY = 0.32;
    const titleH = 0.68;
    const metaY = 1.03;
    const conditionsY = 1.58;

    slide.addText(shortName, {
      x: SLIDE.M,
      y: titleY,
      w: SLIDE.W - (SLIDE.M * 2) - 2.4,
      h: titleH,
      fontFace: ctx.theme.typography.title,
      fontSize: 20,
      bold: true,
      color: ctx.theme.palette.textStrong,
      margin: 0,
      shrinkText: true,
      valign: "top",
    });
    slide.addText(`Modified: ${policy.lastModified}`, {
      x: SLIDE.M,
      y: metaY,
      w: 3.6,
      h: 0.2,
      fontFace: ctx.theme.typography.body,
      fontSize: 9,
      color: ctx.theme.palette.textMuted,
      margin: 0,
    });

    slide.addShape(ctx.pres.shapes.ROUNDED_RECTANGLE, {
      x: SLIDE.W - 2.3,
      y: 0.38,
      w: 1.88,
      h: 0.36,
      fill: { color: getToneColors(ctx.theme, stateTone).soft },
      line: { color: getToneColors(ctx.theme, stateTone).strong, width: 0 },
      rectRadius: ctx.theme.radii.pill,
    });
    slide.addText(stateLabel, {
      x: SLIDE.W - 2.3,
      y: 0.455,
      w: 1.88,
      h: 0.18,
      fontFace: ctx.theme.typography.heading,
      fontSize: 11,
      bold: true,
      color: getToneColors(ctx.theme, stateTone).strong,
      align: "center",
      margin: 0,
    });

    addPolicyConditionRow(slide, ctx, policy, conditionsY);

    const leftX = SLIDE.M;
    const leftY = conditionsY + 0.52;
    const leftW = 2.05;
    const leftH = 2.78;
    const centerX = leftX + leftW + 0.18;
    const centerW = 3.05;
    const rightX = centerX + centerW + 0.14;
    const rightW = SLIDE.W - SLIDE.M - rightX;

    const activeGrantControls = new Set(collectActiveGrantControls(policy));
    activeGrantControls.delete("Block access");
    const grantControls = ALL_GRANT_CONTROLS
      .filter((control) => control !== "Block access")
      .concat(
        Array.from(activeGrantControls).filter((control) => !ALL_GRANT_CONTROLS.includes(control) && control !== "Block access")
      );

    const activeSessionControls = new Set(collectActiveSessionControls(policy));
    const sessionControls = ALL_SESSION_CONTROLS.concat(
      Array.from(activeSessionControls).filter((control) => !ALL_SESSION_CONTROLS.includes(control))
    );

    const grantSubtitle = policy.grantControls.operator;
    const sessionSubtitle = policy.sessionControls.signInFrequency
      ? `Sign-in frequency: ${policy.sessionControls.signInFrequency}`
      : null;
    const grantRowH = 0.175;
    const sessionRowH = 0.14;
    const actionY = leftY + 0.08;
    const grantY = leftY + 0.46;
    const grantH = leftH - 0.46;
    const appsH = 1.08;
    const sessionH = leftH - appsH - 0.08;
    const sessionY = leftY + appsH + 0.08;

    addPolicyDetailPanel(slide, ctx, {
      x: leftX,
      y: leftY,
      w: leftW,
      h: leftH,
      title: "Users",
      tone: "neutral",
    });
    addPolicyDetailLines(slide, ctx, {
      x: leftX,
      y: leftY,
      w: leftW,
      h: leftH,
      lines: buildPolicyUsersLines(policy),
    });

    slide.addText(policy.action === "Block" ? "Block access" : "Grant access", {
      x: centerX,
      y: actionY,
      w: centerW,
      h: 0.26,
      fontFace: ctx.theme.typography.heading,
      fontSize: 22,
      bold: true,
      color: actionColors.strong,
      align: "center",
      margin: 0,
    });
    slide.addShape(ctx.pres.shapes.LINE, {
      x: centerX,
      y: actionY + 0.27,
      w: centerW - 0.12,
      h: 0,
      line: { color: actionColors.strong, width: 2 },
    });
    slide.addShape(ctx.pres.shapes.CHEVRON, {
      x: centerX + centerW - 0.12,
      y: actionY + 0.2,
      w: 0.12,
      h: 0.14,
      fill: { color: actionColors.strong },
      line: { color: actionColors.strong, width: 0 },
    });
    addControlChecklist(slide, ctx, {
      x: centerX,
      y: grantY,
      w: centerW,
      h: grantH,
      title: "Grant Controls",
      subtitle: grantSubtitle,
      controls: grantControls,
      activeSet: activeGrantControls,
      tone: "neutral",
      rowH: grantRowH,
    });

    addPolicyDetailPanel(slide, ctx, {
      x: rightX,
      y: leftY,
      w: rightW,
      h: appsH,
      title: "Apps",
      tone: "neutral",
    });
    addPolicyDetailLines(slide, ctx, {
      x: rightX,
      y: leftY,
      w: rightW,
      h: appsH,
      lines: buildPolicyAppLines(policy),
    });

    addControlChecklist(slide, ctx, {
      x: rightX,
      y: sessionY,
      w: rightW,
      h: sessionH,
      title: "Session Controls",
      subtitle: sessionSubtitle,
      controls: sessionControls,
      activeSet: activeSessionControls,
      tone: "neutral",
      compact: true,
      rowH: sessionRowH,
    });

    slide.addText(`Policy ${index + 1} of ${sorted.length} | ID ${policy.id}`, {
      x: SLIDE.M,
      y: SLIDE.H - 0.24,
      w: 2.8,
      h: 0.14,
      fontFace: ctx.theme.typography.body,
      fontSize: 7,
      color: ctx.theme.palette.textSubtle,
      margin: 0,
    });
  });
}

function createPresentationContext(pres, theme, analysis, policies, projectDir) {
  const meta = analysis.meta || {};
  const enabledCount = meta.enabledCount || policies.filter((p) => p.state === "enabled").length;
  const reportOnlyCount = meta.reportOnlyCount || policies.filter((p) => p.state === "report_only").length;
  const disabledCount = meta.disabledCount || policies.filter((p) => p.state === "disabled").length;
  return {
    pres,
    theme,
    analysis,
    policies,
    projectDir,
    slideNumber: 0,
    tenantName: sanitizeText(pickFirst(meta.clientName, "Tenant"), 34),
    reportDate: sanitizeText(pickFirst(meta.date, new Date().toISOString().slice(0, 10)), 28),
    logoPath: resolveLogoPath(meta.logoPath, theme.metadata.logoPath, projectDir),
    policyCounts: {
      total: meta.policyCount || policies.length,
      enabled: enabledCount,
      reportOnly: reportOnlyCount,
      disabled: disabledCount,
    },
  };
}

function buildDeck(ctx) {
  addCoverSlide(ctx, ctx.analysis);
  addAgendaSlide(ctx);
  addScorecardSlide(ctx, ctx.analysis);
  addExecutiveSummarySlide(ctx, ctx.analysis);
  addTopPrioritiesSlides(ctx, ctx.analysis);
  addRoadmapSlides(ctx, ctx.analysis);

  addEvidenceDividerSlide(ctx);
  addPolicyLandscapeSlides(ctx, ctx.analysis);
  addGeolocationSlides(ctx, ctx.analysis);
  addMfaMatrixSlides(ctx, ctx.analysis);
  addRiskPolicySlides(ctx, ctx.analysis);
  addAuthStrengthSlides(ctx, ctx.analysis);
  addPimCoverageSlides(ctx, ctx.analysis);
  addReportOnlyPipelineSlides(ctx, ctx.analysis);
  addMsManagedOverlapSlides(ctx, ctx.analysis);
  addRecommendationsSlides(ctx, ctx.analysis);

  addAppendixDividerSlide(ctx, "Appendix", "Policy inventory and detail are separated from the executive narrative.");
  addPolicyMatrixSlides(ctx);
  addPolicyDetailSlides(ctx);
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    printHelp();
    return;
  }

  const projectDir = path.resolve(__dirname, "..");
  const analysisPath = path.resolve(projectDir, options.analysis || "analysis.json");
  const policiesPath = path.resolve(projectDir, options.policies || "policies.json");
  const outputPath = path.resolve(projectDir, options.output || "CA_Security_Posture_Report.pptx");

  const analysis = readJson(analysisPath, "analysis JSON");
  validateAnalysis(analysis);
  const policiesRaw = readJson(policiesPath, "policies JSON");
  _skipGuidSanitize = !detectRawGuids(policiesRaw);
  const policies = normalizePolicies(policiesRaw);
  if (!policies.length) throw new Error("No policies found in policies input.");

  const theme = loadTheme(options.theme, projectDir);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "CA Documenter";
  pres.title = sanitizeText(pickFirst(theme.metadata.title, "Conditional Access Security Posture Report"), 64);
  pres.subject = "Conditional Access posture analysis";
  pres.company = sanitizeText(pickFirst(analysis.meta && analysis.meta.clientName, "CA Documenter"), 48);

  const ctx = createPresentationContext(pres, theme, analysis, policies, projectDir);
  buildDeck(ctx);

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await pres.writeFile({ fileName: outputPath });
  console.log(`Generated ${outputPath}`);
  console.log(`Slides generated: ${ctx.slideNumber}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
