"use strict";

module.exports = {
  metadata: {
    title: "Conditional Access Security Posture Report",
    subtitle: "Client-facing executive briefing",
    logoPath: null,
  },
  palette: {
    canvas: "0A0F1A", // Deep Obsidian Black
    surface: "161E31", // Premium Dark Blue-Grey
    surfaceAlt: "101626", // Slightly darker surface
    surfaceStrong: "212D45",
    border: "293753",
    brand: "0A0F1A", // Matches canvas
    brandSoft: "00F0FF", // Neon Cyan Accent
    brandAccent: "8533FF", // Neon Violet Accent
    textStrong: "FFFFFF",
    textBody: "DFE6F5",
    textMuted: "94A3B8",
    textSubtle: "64748B",
    positive: "10B981", // Rich Emerald
    positiveSoft: "064E3B",
    caution: "F59E0B", // Vibrant Amber
    cautionSoft: "78350F",
    critical: "EF4444", // Crisp Red
    criticalSoft: "7F1D1D",
    neutral: "94A3B8",
    neutralSoft: "1E293B",
  },
  typography: {
    title: "Inter",
    heading: "Inter",
    body: "Inter",
    mono: "Consolas",
  },
  spacing: {
    xs: 0.05,
    sm: 0.1,
    md: 0.16,
    lg: 0.24,
    xl: 0.36,
  },
  radii: {
    sm: 0.04,
    md: 0.08,
    lg: 0.12,
    pill: 0.2,
  },
  effects: {
    glass: {
      transparency: 70,
      shadow: { type: "outer", color: "000000", opacity: 0.45, blur: 15, offset: 4, angle: 90 }
    },
    shadowSm: { type: "outer", color: "000000", opacity: 0.3, blur: 8, offset: 2, angle: 90 },
    shadowLg: { type: "outer", color: "000000", opacity: 0.6, blur: 20, offset: 8, angle: 90 }
  },
  accents: {
    severity: {
      high: "EF4444",
      medium: "F59E0B",
      low: "3B82F6",
    },
    state: {
      enabled: "10B981",
      report_only: "F59E0B",
      disabled: "EF4444",
    },
  },
};
