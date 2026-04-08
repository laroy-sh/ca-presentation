# Analysis JSON Schema

This document defines the structure of `analysis.json` — the contract between Claude's policy analysis and the PptxGenJS report template.

## Top-Level Structure

```json
{
  "meta": { ... },
  "executiveSummary": { ... },
  "policyLandscape": { ... },
  "geolocationStrategy": { ... },
  "mfaMatrix": { ... },
  "riskPolicies": { ... },
  "authStrengths": { ... },
  "pimCoverage": { ... },
  "reportOnlyPipeline": { ... },
  "msManagedOverlap": { ... },
  "recommendations": { ... },
  "assessment": { ... }
}
```

## Section Details

### meta

```json
{
  "date": "April 8, 2026",
  "nextReview": "Q3 2026",
  "policyCount": 33,
  "enabledCount": 21,
  "reportOnlyCount": 7,
  "disabledCount": 5,
  "statsFooter": "70 trusted locations  •  19 PIM roles  •  6 auth strengths"
}
```

- `date`: generation date (today)
- `nextReview`: suggested next review (current quarter + 1)
- `statsFooter`: one-line summary of supplementary data. Omit or set to `null` if no extra data provided.

### executiveSummary

```json
{
  "strengths": [
    "Legacy authentication blocked for all users",
    "Risk-based policies cover all levels (Low-High)"
  ],
  "concerns": [
    "Custom MFA allows SMS - phishable factor",
    "7 policies stuck in report-only mode"
  ]
}
```

- 3-6 items per list. Lead with the most impactful.
- Keep each string under 60 characters for slide readability.

### policyLandscape

```json
{
  "categories": [
    { "label": "Geolocation\nBlocks", "count": 10, "colorKey": "teal" },
    { "label": "MFA\nEnforcement", "count": 7, "colorKey": "green" },
    { "label": "Risk-Based\nPolicies", "count": 3, "colorKey": "amber" },
    { "label": "Access\nBlocking", "count": 3, "colorKey": "red" },
    { "label": "Device &\nSession", "count": 2, "colorKey": "tealDark" },
    { "label": "Auth\nRegistration", "count": 1, "colorKey": "muted" },
    { "label": "Microsoft-\nManaged", "count": 5, "colorKey": "mutedDark" }
  ]
}
```

- `colorKey`: one of `teal`, `green`, `amber`, `red`, `tealDark`, `muted`, `mutedDark`
- Maximum 7 categories (more won't fit on the slide)
- Labels use `\n` for two-line display in the cards
- Categories should sum to the total policy count

### geolocationStrategy

```json
{
  "available": true,
  "layers": [
    {
      "id": 1,
      "title": "Catch-All Block",
      "description": "\"Geolocation - Main policy\" blocks 239 countries. Only CA, US, IN, NL, PH, CR, UK, GR, BR, CO, IR allowed.",
      "accentColor": "teal"
    },
    {
      "id": 2,
      "title": "Country-Specific Blocks",
      "countries": [
        { "name": "Colombia (CO)", "state": "Enabled" },
        { "name": "US", "state": "Report-Only" }
      ],
      "accentColor": "amber"
    },
    {
      "id": 3,
      "title": "Trusted Office Networks (MFA Exclusion Zones)",
      "locationCount": 70,
      "description": "70 named locations across Canada and UK.",
      "accentColor": "green"
    }
  ],
  "note": "US geo block remains in report-only since Sep 2024 (18+ months)."
}
```

- Set `available: false` if no geolocation policies exist or insufficient data to analyze.
- Layer 1 is always the broadest block. Layer 2 is country-specific. Layer 3 is trusted networks.
- If named location data is not provided, Layer 3 can be omitted or show count from policy conditions.

### mfaMatrix

```json
{
  "available": true,
  "policies": [
    {
      "name": "MFA - Standard to all users",
      "scope": "All users (10+ group excl.)",
      "authStrength": "Built-in MFA",
      "frequency": "15 days",
      "conditions": "Trusted IPs + 4 offices excl."
    }
  ],
  "callout": {
    "title": "Passwordless MFA for Admins",
    "text": "Allows deviceBasedPush only. Strong but not phishing-resistant."
  }
}
```

- Include only MFA-related policies (grant control includes MFA or auth strength).
- `callout`: optional highlight box. Set to `null` if nothing noteworthy.
- Maximum 7 rows (table space limit).

### riskPolicies

```json
{
  "available": true,
  "policies": [
    {
      "title": "High User Risk",
      "policyName": "Azure Risky Users - High",
      "grantControl": "MFA + Password Change",
      "operator": "Require ALL",
      "scope": "All users",
      "modified": "2024-11-19",
      "accentColor": "red"
    }
  ],
  "callout": {
    "text": "Full Risk Coverage - All user risk and sign-in risk levels protected.",
    "color": "green"
  }
}
```

- Set `available: false` if no risk-based policies exist.
- Maximum 3 cards (slide layout limit).
- `accentColor`: `red` for high risk, `amber` for medium/low, `teal` for sign-in risk.
- `callout.color`: `green` if full coverage, `red` if gaps exist.

### authStrengths

```json
{
  "available": true,
  "strengths": [
    {
      "name": "Passwordless MFA",
      "rating": "Strong",
      "ratingColor": "green",
      "type": "Custom",
      "methods": "deviceBasedPush",
      "note": "Used by admin MFA + PIM activation"
    }
  ]
}
```

- Set `available: false` if no authentication strength data is provided or policies don't use auth strengths.
- `rating`: "Weak", "Strong", or "Strongest"
- `ratingColor`: `red` for Weak, `green` for Strong, `teal` for Strongest
- Maximum 4 cards (slide layout limit).

### pimCoverage

```json
{
  "available": true,
  "roles": [
    {
      "role": "Global Administrator",
      "eligible": "3",
      "active": "2",
      "caCoverage": "MFA - All Admins, Passwordless, PIM",
      "direct": "Yes"
    }
  ],
  "note": "\"Partial\" = covered only via PIM auth context, not by role-targeted CA policy."
}
```

- Set `available: false` if no PIM role assignment data is provided.
- `direct`: "Yes", "Partial", or "No"
- Maximum 10 rows (table space limit).

### reportOnlyPipeline

```json
{
  "policies": [
    {
      "name": "Token Protection Policy",
      "grantControl": "Token protection",
      "targetApps": "Exchange + SharePoint",
      "modified": "2025-08-05",
      "priority": "High"
    }
  ],
  "callout": {
    "title": "Action Required",
    "text": "Top 3 priorities have been in report-only since August 2025 (8+ months)."
  }
}
```

- Always populated if report-only policies exist.
- `priority`: "High", "Medium", or "Low" — based on staleness and security impact.
- Maximum 7 rows (table space limit).

### msManagedOverlap

```json
{
  "policies": [
    {
      "name": "MFA for admins at Admin Portals",
      "description": "MFA for 10+ admin roles at Admin Portals",
      "customEquivalent": "MFA - All Administrators",
      "overlap": "Full"
    }
  ],
  "callout": {
    "title": "Critical Gap",
    "text": "No custom policy enforces phishing-resistant MFA for admin roles."
  }
}
```

- Always populated if Microsoft-managed policies exist (displayName starts with "Microsoft-managed:").
- `overlap`: "Full", "Partial", or "Gap"
- `callout`: set to `null` if no critical gaps found.

### recommendations

```json
{
  "high": [
    "Enforce phishing-resistant MFA for admin roles",
    "Remove SMS from Custom MFA authentication strength"
  ],
  "medium": [
    "Block authentication transfer for admin roles",
    "Eliminate standing Global Admin assignments"
  ],
  "low": [
    "Enforce US geolocation block (18+ months in report-only)"
  ]
}
```

- Each array: 1-4 items. Keep strings under 60 characters.
- Empty arrays are allowed (e.g., no low-priority items).

### assessment

```json
{
  "verdict": "GOOD - with critical improvements needed",
  "verdictColor": "amber",
  "prioritySummary": "3 high-priority items  •  3 medium-priority items  •  2 low-priority items",
  "criticalGap": "No phishing-resistant MFA enforced for privileged roles",
  "dataSources": "CA Policies, PIM, Named Locations, Auth Strength, Role Definitions"
}
```

- `verdictColor`: `green` (strong), `amber` (good with issues), `red` (critical gaps)
- `criticalGap`: single most important finding. Set to `null` if none.
- `dataSources`: comma-separated list of data sources used in the analysis.
