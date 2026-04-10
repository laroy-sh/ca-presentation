# CA Documenter Skill

Generate an executive-first, client-facing PowerPoint report from Microsoft Entra Conditional Access policy data.

## Trigger

Use this skill when:
- The user provides CA policy JSON (Graph envelope, array, or existing `policies.json`)
- The user asks to document CA posture, generate a CA report, or produce executive security slides

## Inputs

### Required
- **CA Policy JSON**:
  - Graph envelope (`{ "value": [ ... ] }`), or
  - bare policy array, or
  - existing `policies.json`

### Optional enrichment
- **Named Locations JSON** for geolocation narrative
- **Authentication Strength JSON** for auth strength analysis
- **PIM Role Assignments JSON** for privileged role coverage analysis

## Workflow

### 1. Parse and normalize policies

- If input is Graph envelope, use `.value`
- Normalize policy state:
  - `enabled` -> `enabled`
  - `enabledForReportingButNotEnforced` -> `report_only`
  - `disabled` -> `disabled`

Report counts in this format:
`Found N policies (X enabled, Y report-only, Z disabled).`

### 2. Produce `analysis.json` using `skill/analysis-schema.md`

The report now expects an executive narrative before appendix content. Keep sections concise and slide-safe.

#### Executive summary requirements

- Populate `executiveSummary.strengths` and `executiveSummary.concerns` (3-6 each).
- Add optional `executiveSummary.topPriorities`:
  - string or object entries (`title`, optional `priority`, optional `evidence`)
  - keep phrasing action-oriented and executive-ready

#### Roadmap requirements

- Add optional `roadmap` with:
  - `nearTerm`
  - `midTerm`
- Group recommendations by urgency and implementation effort assumptions.
- Use short action statements suitable for slide cards.

#### Assessment requirements

- Continue `assessment.verdict`, `prioritySummary`, and `criticalGap`.
- Add optional:
  - `assessment.score` (0-100)
  - `assessment.level` (for example: `Strong`, `Stable`, `Needs Improvement`)

#### Branding metadata

- Optional:
  - `meta.clientName`
  - `meta.logoPath`

If these fields are absent, generator fallback behavior is automatic.

### 3. Write artifacts to project root

- `analysis.json`
- `policies.json`

### 4. Generate deck

```bash
node skill/generate_report.js
```

Optional explicit paths:

```bash
node skill/generate_report.js \
  --analysis analysis.json \
  --policies policies.json \
  --output CA_Security_Posture_Report.pptx
```

### 5. Visual QA

Recommended flow:

```bash
npm run qa:render
npm run qa:slides
```

`qa:render` generates sample PPTX and converts to PDF/JPG previews when local tools exist.

`qa:slides` runs deterministic preflight checks:
- bounded table row counts
- bounded executive card counts
- bounded text lengths for slide-critical fields

## Analysis guidance (unchanged logic, updated output style)

- Keep classification and detection logic from prior version:
  - geolocation, MFA, risk, auth strengths, PIM, report-only, Microsoft-managed overlap
- Preserve recommendation severity grouping: `high`, `medium`, `low`
- Preserve existing optional section toggles (`available: true|false`)
- Prefer concise phrasing and avoid verbose raw identifiers in output fields

## Well-known label resolution

Resolve these before writing final text values where possible:

- `00000003-0000-0000-c000-000000000000` -> `Microsoft Graph`
- `00000002-0000-0ff1-ce00-000000000000` -> `Office 365 Exchange Online`
- `00000003-0000-0ff1-ce00-000000000000` -> `Microsoft SharePoint Online`
- `cc15fd57-2c6c-4117-a88c-83b1d56b4bbe` -> `Microsoft Teams Services`
- `Office365` -> `Office 365`
- `MicrosoftAdminPortals` -> `Microsoft Admin Portals`

## PptxGenJS safety rules

1. No `#` prefix in color hex values.
2. Do not encode transparency in hex color strings.
3. Avoid reusing mutable style objects across unrelated elements.
4. Set `margin: 0` for text where precise alignment is required.
