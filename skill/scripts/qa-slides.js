"use strict";

const fs = require("fs");
const path = require("path");

const ROOT_DIR = path.resolve(__dirname, "..", "..");
const ANALYSIS_PATH = path.resolve(ROOT_DIR, "skill/examples/analysis-example.json");
const POLICIES_PATH = path.resolve(ROOT_DIR, "skill/examples/policies-example.json");

function readJson(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Missing file: ${filePath}`);
  }
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function hasArray(value) {
  return Array.isArray(value) && value.length > 0;
}

function check(condition, label, failures) {
  if (!condition) failures.push(label);
}

function warn(condition, label, warnings) {
  if (!condition) warnings.push(label);
}

function maxLength(list, limit) {
  return (list || []).every((item) => String(typeof item === "string" ? item : item.title || item.text || "").length <= limit);
}

function runChecks() {
  const failures = [];
  const warnings = [];
  const analysis = readJson(ANALYSIS_PATH);
  const policies = readJson(POLICIES_PATH);

  check(analysis.meta, "meta section is required", failures);
  check(analysis.executiveSummary, "executiveSummary section is required", failures);
  check(analysis.recommendations, "recommendations section is required", failures);

  warn(hasArray(analysis.executiveSummary && analysis.executiveSummary.strengths), "executiveSummary.strengths is empty", warnings);
  warn(hasArray(analysis.executiveSummary && analysis.executiveSummary.concerns), "executiveSummary.concerns is empty", warnings);
  warn(maxLength((analysis.executiveSummary && analysis.executiveSummary.strengths) || [], 120), "strength item exceeds 120 chars", warnings);
  warn(maxLength((analysis.executiveSummary && analysis.executiveSummary.concerns) || [], 120), "concern item exceeds 120 chars", warnings);

  const topPriorities = (analysis.executiveSummary && analysis.executiveSummary.topPriorities) || [];
  warn(topPriorities.length <= 8, "topPriorities contains more than 8 items", warnings);

  const roadmap = analysis.roadmap || {};
  warn(((roadmap.nearTerm || []).length <= 10), "roadmap.nearTerm contains more than 10 items", warnings);
  warn(((roadmap.midTerm || []).length <= 10), "roadmap.midTerm contains more than 10 items", warnings);

  const mfaCount = ((analysis.mfaMatrix && analysis.mfaMatrix.policies) || []).length;
  warn(mfaCount <= 25, "mfaMatrix.policies contains more than 25 rows", warnings);

  const pimCount = ((analysis.pimCoverage && analysis.pimCoverage.roles) || []).length;
  warn(pimCount <= 30, "pimCoverage.roles contains more than 30 rows", warnings);

  const reportOnlyCount = ((analysis.reportOnlyPipeline && analysis.reportOnlyPipeline.policies) || []).length;
  warn(reportOnlyCount <= 30, "reportOnlyPipeline.policies contains more than 30 rows", warnings);

  const msCount = ((analysis.msManagedOverlap && analysis.msManagedOverlap.policies) || []).length;
  warn(msCount <= 30, "msManagedOverlap.policies contains more than 30 rows", warnings);

  const recommendationCount =
    ((analysis.recommendations && analysis.recommendations.high) || []).length +
    ((analysis.recommendations && analysis.recommendations.medium) || []).length +
    ((analysis.recommendations && analysis.recommendations.low) || []).length;
  warn(recommendationCount <= 18, "recommendation count exceeds 18 and may cause crowded slides", warnings);
  warn(
    maxLength((analysis.recommendations && analysis.recommendations.high) || [], 96) &&
      maxLength((analysis.recommendations && analysis.recommendations.medium) || [], 96) &&
      maxLength((analysis.recommendations && analysis.recommendations.low) || [], 96),
    "one or more recommendation strings exceed 96 chars",
    warnings
  );

  const longPolicyNames = (policies || []).filter((policy) => String(policy.name || "").length > 72);
  warn(longPolicyNames.length === 0, "one or more policy names exceed 72 chars and may wrap in detail headers", warnings);

  if (failures.length) {
    console.error("QA FAILED");
    failures.forEach((item) => console.error(`- ${item}`));
    process.exit(1);
  }

  console.log("QA PASSED");
  if (warnings.length) {
    console.log("Warnings:");
    warnings.forEach((item) => console.log(`- ${item}`));
  } else {
    console.log("No warnings.");
  }
}

try {
  runChecks();
} catch (error) {
  console.error(`QA FAILED: ${error.message}`);
  process.exit(1);
}
