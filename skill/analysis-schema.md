# Analysis JSON Schema

`analysis.json` is the presentation contract consumed by `skill/generate_report.js`.

The generator now follows an executive-first narrative and supports optional branding + roadmap fields while remaining backward compatible with older analysis payloads.

## Top-Level Shape

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
  "roadmap": { ... },
  "assessment": { ... }
}
```

## Core Sections

### `meta`

```json
{
  "date": "April 8, 2026",
  "nextReview": "Q3 2026",
  "policyCount": 33,
  "enabledCount": 21,
  "reportOnlyCount": 7,
  "disabledCount": 5,
  "clientName": "Contoso",
  "logoPath": "docs/branding/contoso-logo.png",
  "statsFooter": "70 trusted locations | 19 PIM roles | 6 auth strengths"
}
```

- `clientName` is optional and used in header/footer context.
- `logoPath` is optional and can be relative to repo root.
- If omitted, generator falls back to `"Tenant"` and no logo.

### `executiveSummary`

```json
{
  "strengths": [
    "Legacy authentication blocked for all users"
  ],
  "concerns": [
    "Phishing-resistant MFA is not enforced for admins"
  ],
  "topPriorities": [
    {
      "title": "Enforce phishing-resistant MFA for admin roles",
      "priority": "High",
      "evidence": "No equivalent custom control currently enforced"
    },
    "Move stale report-only controls to enforcement"
  ]
}
```

- `topPriorities` is optional.
- Item shape supports either string or object.
- If `topPriorities` is missing, generator derives priorities from `recommendations`.

### `policyLandscape`

```json
{
  "categories": [
    { "label": "MFA Enforcement", "count": 7, "colorKey": "green" }
  ]
}
```

- Category count is adaptive; generator paginates cards when needed.
- `colorKey` remains informational for analysis output.

### `geolocationStrategy`

```json
{
  "available": true,
  "layers": [
    {
      "id": 1,
      "title": "Catch-all geo block",
      "description": "Blocks non-approved countries",
      "accentColor": "teal"
    }
  ],
  "note": "One country policy remains report-only"
}
```

- Set `available: false` when geolocation detail is not present.
- Layers are rendered as evidence cards and continue across slides when needed.

### `mfaMatrix`

```json
{
  "available": true,
  "policies": [
    {
      "name": "MFA - All users",
      "scope": "All users",
      "authStrength": "Built-in MFA",
      "frequency": "14 days",
      "conditions": "Trusted locations excluded"
    }
  ],
  "callout": {
    "title": "Admin MFA posture",
    "text": "Admin MFA uses standard strength instead of phishing-resistant methods."
  }
}
```

- Generator paginates rows to avoid overflow.

### `riskPolicies`

```json
{
  "available": true,
  "policies": [
    {
      "title": "High user risk",
      "policyName": "Risky Users - High",
      "grantControl": "MFA + Password Change",
      "operator": "Require ALL",
      "scope": "All users",
      "modified": "2024-11-19",
      "accentColor": "red"
    }
  ],
  "callout": {
    "text": "High-risk and sign-in risk coverage is complete.",
    "color": "green"
  }
}
```

### `authStrengths`

```json
{
  "available": true,
  "strengths": [
    {
      "name": "Passwordless MFA",
      "rating": "Strongest",
      "ratingColor": "teal",
      "type": "Custom",
      "methods": "fido2, windowsHelloForBusiness",
      "note": "Used by privileged access policies"
    }
  ]
}
```

### `pimCoverage`

```json
{
  "available": true,
  "roles": [
    {
      "role": "Global Administrator",
      "eligible": "3",
      "active": "1",
      "caCoverage": "Admin MFA policy",
      "direct": "Yes"
    }
  ],
  "note": "Partial means coverage is only via auth context"
}
```

### `reportOnlyPipeline`

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
    "title": "Action required",
    "text": "Security-critical report-only items are stale."
  }
}
```

### `msManagedOverlap`

```json
{
  "policies": [
    {
      "name": "Phishing-resistant MFA for admins",
      "description": "WHfB/FIDO2/x509 requirement",
      "customEquivalent": "No custom equivalent",
      "overlap": "Gap"
    }
  ],
  "callout": {
    "title": "Critical gap",
    "text": "No custom policy enforces phishing-resistant admin MFA."
  }
}
```

### `recommendations`

```json
{
  "high": ["Enforce phishing-resistant admin MFA"],
  "medium": ["Reduce stale report-only controls"],
  "low": ["Clean up legacy exceptions"]
}
```

## New Executive Roadmap Contract

### `roadmap` (optional)

```json
{
  "nearTerm": [
    {
      "title": "Enable token protection controls",
      "priority": "Near-term",
      "evidence": "High priority report-only finding"
    }
  ],
  "midTerm": [
    "Align privileged role targeting with direct CA coverage"
  ]
}
```

- `nearTerm` and `midTerm` support string or object entries.
- If `roadmap` is missing, generator derives roadmap columns from recommendations.

## Assessment Contract

### `assessment`

```json
{
  "score": 72,
  "level": "Stable",
  "verdict": "GOOD - improvements required",
  "verdictColor": "amber",
  "prioritySummary": "3 high | 2 medium | 1 low",
  "criticalGap": "No phishing-resistant MFA for privileged roles",
  "dataSources": "CA Policies, Named Locations, PIM, Auth Strengths"
}
```

- `score` and `level` are optional.
- If absent, generator derives score from policy state distribution.

## Backward Compatibility Rules

- Older payloads without `topPriorities`, `roadmap`, `score`, `level`, `clientName`, or `logoPath` still render.
- Optional sections (`geolocationStrategy`, `mfaMatrix`, `riskPolicies`, `authStrengths`, `pimCoverage`) are skipped when unavailable.
- Generator enforces adaptive max-content rules and paginates tables/cards to avoid overflow.
