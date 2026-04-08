# CA Documenter Skill

Generate a professional PowerPoint security posture report from Microsoft Entra Conditional Access policies.

## Trigger

Use this skill when:
- The user provides CA policy JSON (Graph API export or `policies.json`)
- The user asks to "document CA policies", "generate a CA report", or "create a security posture report"

## Inputs

### Required
- **CA Policy JSON** — either:
  - A Microsoft Graph API envelope: `{ "@odata.context": "...", "value": [ ...policies ] }`
  - A bare array of policy objects: `[ ...policies ]`
  - An already-parsed `policies.json` (from a previous run of `extract_data.py`)

### Optional (enrichment data)
- **Named Locations JSON** — enables the Geolocation Strategy slide with country/location details
- **Authentication Strengths JSON** — enables the Auth Strengths slide with method details and ratings
- **PIM Role Assignments JSON** — enables the PIM Coverage slide with role/assignment cross-reference

More input data = more analytical slides in the report. The skill works with just CA policies alone.

## Workflow

### Step 1: Parse and validate input

If the input is a Graph API envelope, extract policies from `.value`. Normalize the state field:
- `"enabled"` stays `"enabled"`
- `"enabledForReportingButNotEnforced"` becomes `"report_only"`
- `"disabled"` stays `"disabled"`

If the input is already a `policies.json` (array of objects with `name`, `state`, `users`, `grantControls`, etc.), use it directly.

Report: "Found N policies (X enabled, Y report-only, Z disabled)"

### Step 2: Analyze policies

Analyze the policies and produce `analysis.json` following the schema in `skill/analysis-schema.md`. Read that file for the full JSON structure.

#### Policy Classification

Classify each policy into exactly one category:

| Category | Detection Rule |
|----------|---------------|
| Geolocation Blocks | `displayName` contains "Geolocation" or "Geo", OR `conditions.locations` targets specific countries |
| MFA Enforcement | `grantControls` contains MFA or authentication strength (not geo-related, not risk-related) |
| Risk-Based Policies | `conditions.userRiskLevels` or `conditions.signInRiskLevels` is non-empty |
| Access Blocking | `grantControls` contains "block" AND not geo-related AND not legacy-auth-related |
| Device & Session | `sessionControls` configured OR `conditions.platforms` configured OR device registration actions |
| Auth Registration | `applications.includeUserActions` contains "registerOrJoinDevices" or "registerSecurityInformation" |
| Microsoft-Managed | `displayName` starts with "Microsoft-managed:" |

If a policy fits multiple categories, prefer: Risk-Based > Microsoft-Managed > Geolocation > MFA > Auth Registration > Device & Session > Access Blocking.

#### Strengths Detection

Check for each and include if true:
- Legacy authentication blocked? (policy with block + clientApps containing "exchangeActiveSync" or "other")
- Device code flow blocked? (policy blocking "deviceCodeFlow" auth flow)
- Risk-based coverage complete? (policies covering low, medium, AND high user risk + high sign-in risk)
- Geo-blocking present? (any geolocation block policies enabled)
- MFA broadly enforced? (MFA policy targeting "All users")
- Trusted location exclusions? (named locations used as MFA exclusion zones)

#### Concerns Detection

Check for each and include if true:
- SMS in any authentication strength? (phishable factor — always flag)
- Report-only policies stale? (any report-only policy with `modifiedDateTime` > 6 months ago)
- Phishing-resistant MFA not enforced? (no policy requires WHfB, FIDO2, or x509 for admin roles)
- Standing privileged assignments? (if PIM data shows active > 0 for Global Admin — should use eligible-only)
- MS-managed policies disabled without custom equivalent? (overlap = "Gap")
- Device code flow block too narrow? (block targets specific groups instead of all users)

#### MFA Matrix

For each MFA-related policy, extract:
- `name`: policy display name
- `scope`: who it targets (resolve to human-readable: "All users", "Admin roles (N roles)", group count)
- `authStrength`: which authentication strength is required
- `frequency`: sign-in frequency if configured
- `conditions`: key conditions (excluded locations, specific apps, etc.)

Callout: flag if admin MFA uses standard strength instead of phishing-resistant.

#### Risk Policies

For each risk-based policy, extract title, policy name, grant control, operator, scope, modified date. Use accent colors: `red` for high risk, `amber` for medium/low, `teal` for sign-in risk. Callout should note whether coverage is complete or has gaps.

#### Geolocation Strategy (if data available)

Build a layered view:
- **Layer 1**: The broadest geo-block (catch-all). Note how many countries are blocked and which are allowed.
- **Layer 2**: Country-specific block policies with their state (enabled/report-only).
- **Layer 3**: Trusted network locations used as MFA exclusion zones (if named location data provided).

#### Auth Strengths (if data available)

For each authentication strength policy, assess:
- `rating`: "Weak" if SMS/phone included, "Strongest" if only WHfB/FIDO2/x509, "Strong" otherwise
- Flag any strength that includes phishable factors

#### PIM Coverage (if data available)

Cross-reference PIM role assignments with CA policy targeting:
- `direct`: "Yes" if a CA policy explicitly targets that admin role
- `direct`: "Partial" if only covered via PIM auth context (c1)
- `direct`: "No" if no CA policy covers the role

#### Report-Only Pipeline

For each report-only policy, assess priority:
- **High**: Security-critical (token protection, auth transfer blocking, phishing-resistant MFA) AND stale (> 6 months)
- **Medium**: Moderately important (admin portal blocks, geo-blocks) OR moderately stale (3-6 months)
- **Low**: Low risk (BYOD pilots, vendor access) OR fresh (< 3 months)

#### MS-Managed Overlap

For each Microsoft-managed policy (displayName starts with "Microsoft-managed:"):
- Find the closest custom equivalent by comparing what the policy does
- Rate overlap: "Full" (custom policy covers same scope), "Partial" (custom covers some), "Gap" (no custom equivalent)
- Critical gap: flag if "Require phishing-resistant multifactor authentication for admins" has no custom equivalent

#### Recommendations

Prioritize findings into HIGH / MEDIUM / LOW:
- **HIGH**: Missing phishing-resistant MFA for admins, phishable factors (SMS) in auth strengths, critical security features stuck in report-only
- **MEDIUM**: Stale report-only policies (6-12 months), uncovered privileged roles, standing admin assignments
- **LOW**: Minor scope improvements, long-stale low-risk report-only policies

#### Assessment Verdict

- `"green"`: No high-priority items, <= 2 medium items
- `"amber"`: Any high-priority items OR > 2 medium items
- `"red"`: 3+ high-priority items OR critical gap with no remediation path

### Step 3: Write analysis.json

Write `analysis.json` to the project root directory. Verify it matches the schema in `skill/analysis-schema.md`.

### Step 4: Generate the report

```bash
cd "<project-dir>" && node skill/generate_report.js
```

This reads `analysis.json` + `policies.json` and produces `CA_Security_Posture_Report.pptx`.

### Step 5: Visual QA

Convert to images and inspect:

```bash
soffice --headless --convert-to pdf CA_Security_Posture_Report.pptx
pdftoppm -jpeg -r 150 CA_Security_Posture_Report.pdf qa-slides/slide
```

Use a subagent to visually inspect the analytical slides (typically slides 1-11). Look for:
- Overlapping text elements
- Content cut off at slide edges
- Tables extending beyond slide boundaries
- Low-contrast text
- Uneven spacing

Fix any issues by adjusting `analysis.json` content (e.g., shortening long strings) and regenerating.

## Well-Known GUIDs

Always resolve these in policy data:

```
"All" / "all"                                  -> "All users" / "All cloud apps"
"none" / "None"                                -> "None"
"00000003-0000-0000-c000-000000000000"         -> Microsoft Graph
"00000002-0000-0ff1-ce00-000000000000"         -> Office 365 Exchange Online
"00000003-0000-0ff1-ce00-000000000000"         -> Microsoft SharePoint Online
"cc15fd57-2c6c-4117-a88c-83b1d56b4bbe"         -> Microsoft Teams Services
"Office365"                                     -> Office 365
"MicrosoftAdminPortals"                         -> Microsoft Admin Portals
```

## PptxGenJS Rules

1. NEVER use `#` prefix on hex colors — causes file corruption
2. NEVER encode opacity in the hex string — use `transparency` property
3. NEVER reuse option objects across addShape/addText calls
4. Set `margin: 0` on text boxes for precise alignment with shapes

## Files

| File | Purpose |
|------|---------|
| `skill/SKILL.md` | This file — skill definition and analysis rubric |
| `skill/analysis-schema.md` | JSON schema for analysis.json |
| `skill/generate_report.js` | PptxGenJS template (reads analysis.json + policies.json) |
| `skill/examples/analysis-example.json` | Sanitized example of analysis.json |
