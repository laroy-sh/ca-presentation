# CA Documenter Skill

Generate a professional PowerPoint security posture report from Microsoft Entra Conditional Access policies.

## Skill Structure

```
skill/
├── SKILL.md                  — Skill definition + full analysis rubric
├── analysis-schema.md        — JSON schema for analysis.json
├── generate_report.js        — PptxGenJS template (zero hardcoded content)
├── README.md                 — This file
└── examples/
    └── analysis-example.json — Sanitized reference example
```

## How It Works

1. **User provides** CA policy JSON (required) + optional named locations, auth strengths, PIM data
2. **Claude reads** `skill/SKILL.md` and analyzes the policies using the rubric (classification rules, strengths/concerns detection, recommendation prioritization)
3. **Claude writes** `analysis.json` — all slide content, structured per the schema in `analysis-schema.md`
4. **Template runs**: `node skill/generate_report.js` reads `analysis.json` + `policies.json` and produces the PPTX

```
User provides:                    Claude analyzes:              Template renders:
┌──────────────────────┐         ┌──────────────────┐         ┌──────────────────────────────┐
│ CA policy JSON       │────────>│ Classify policies │────────>│ analysis.json                │
│ (required)           │         │ Detect strengths  │         │ (all slide content)          │
├──────────────────────┤         │ Detect concerns   │         └──────────┬───────────────────┘
│ Named Locations JSON │         │ Assess MFA        │                    │
│ (optional)           │────────>│ Assess risk       │                    ▼
├──────────────────────┤         │ Find overlaps     │         ┌──────────────────────────────┐
│ Auth Strengths JSON  │         │ Prioritize recs   │         │ node skill/generate_report.js│
│ (optional)           │────────>│ Write verdict     │         │ reads analysis.json          │
├──────────────────────┤         └──────────────────┘         │ reads policies.json          │
│ PIM Roles JSON       │                                       └──────────┬───────────────────┘
│ (optional)           │                                                  │
└──────────────────────┘                                                  ▼
                                                              ┌──────────────────────────────┐
                                                              │ CA_Security_Posture_Report.pptx│
                                                              └──────────────────────────────┘
```

## Inputs

### Required

**CA Policy JSON** — one of:
- Microsoft Graph API envelope: `{ "@odata.context": "...", "value": [ ...policies ] }`
- Bare array of policy objects: `[ ...policies ]`
- Already-parsed `policies.json` from a previous run

Get it via PowerShell (requires `Directory.Read.All` + `Policy.Read.All`):

```powershell
Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies' -OutputType Json | Set-Clipboard
```

### Optional (enrichment data)

| Data | Enables | How to Get |
|------|---------|------------|
| Named Locations JSON | Geolocation Strategy slide with country/location details | `Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations'` |
| Authentication Strengths JSON | Auth Strengths slide with method details and ratings | `Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/authenticationStrength/policies'` |
| PIM Role Assignments JSON | PIM Coverage slide with role/assignment cross-reference | `Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments'` |

More input data = more analytical slides in the report. The skill works with just CA policies alone.

## Output

A 30-50 slide PPTX with dark theme, structured as:

### Section A: Executive Overview (always generated)
1. Title Slide — policy counts, generation date
2. Executive Summary — strengths and concerns side-by-side
3. Policy Landscape — state distribution bars + category breakdown cards

### Section B: Deep Analysis (conditional on available data)
4. Geolocation Strategy — layered defense-in-depth view *(requires geo policies)*
5. MFA Enforcement Matrix — policy/scope/strength/frequency table *(requires MFA policies)*
6. Identity Protection & Risk-Based Policies — risk coverage cards *(requires risk policies)*
7. Authentication Strength Policies — rated strength cards *(requires auth strength data)*
8. Privileged Access / PIM Role Coverage — cross-reference table *(requires PIM data)*
9. Report-Only Pipeline — priority-ranked policies under evaluation
10. Microsoft-Managed Policies Overlap — gap analysis
11. Security Gaps & Recommendations — HIGH/MEDIUM/LOW priority

### Section C: Full Policy Matrix (always generated)
12-13. Complete policy list with state, action, and last modified date

### Section D: Per-Policy Detail (always generated)
14-49. One slide per policy showing conditions, users, apps, grant/session controls

### Section E: Closing (always generated)
50. Security Posture Assessment — verdict, priority summary, critical gap

## Conditional Slides

If only CA policies are provided (no extra data), the report generates with: Title, Executive Summary, Policy Landscape, Report-Only Pipeline, MS-Managed Overlap, Recommendations, Full Policy Matrix, Per-Policy Details, and Closing.

The following slides are included only when their data source is available:

| Slide | Requires |
|-------|----------|
| Geolocation Strategy | Geolocation policies in the CA export |
| MFA Enforcement Matrix | MFA-related policies in the CA export |
| Risk-Based Policies | Risk-based policies in the CA export |
| Auth Strength Policies | Authentication Strengths JSON (optional input) |
| PIM Role Coverage | PIM Role Assignments JSON (optional input) |

## Sanitization

- **Zero customer data** in the template (`generate_report.js`) — all content comes from `analysis.json`
- **Example file** (`examples/analysis-example.json`) uses generic/anonymized data
- **SKILL.md** contains only analysis rules and detection logic, no customer specifics
- `analysis.json` in the project root is regenerated each run and is the only file with real customer data

## Dependencies

```bash
npm install pptxgenjs
```

For visual QA (optional):
- LibreOffice (`soffice`) — PPTX to PDF conversion
- Poppler (`pdftoppm`) — PDF to JPEG slides

## Quick Start

```bash
# 1. Place your CA policy data as policies.json in the project root
# 2. Ask Claude to analyze and generate the report
# 3. Or run manually after creating analysis.json:
node skill/generate_report.js
```
