# CA Documenter Presentation Upgrade Plan

## Summary
Reposition CA Documenter as a client-facing reporting product, not just an internal PPTX generator. The implementation should deliver three outcomes in one pass:

- Redesign the generated deck into an executive-first narrative with a clearly separated appendix for policy inventory/detail.
- Introduce a configurable theme/branding layer so the deck can be re-skinned without rewriting slide code.
- Productize the repository with public-facing documentation, runnable scripts, and a sanitized sample output bundle.

The plan assumes the current analysis model remains the source of truth for findings, while the presentation layer and repo packaging are upgraded around it.

## Key Changes

### 1. Reframe the deck around an executive narrative
Implement a new slide sequence in `skill/generate_report.js`:

- Replace the current front section with:
  - Cover slide
  - Agenda / report structure slide
  - Posture scorecard / status overview slide
  - Executive summary slide
  - Top priorities slide
  - Optional 90-day roadmap slide derived from recommendation priority
- Keep the existing analytical sections, but treat them as supporting evidence rather than the core story.
- Insert a clear appendix divider before:
  - Full Policy Matrix
  - Per-policy detail slides
- Preserve all current conditional content areas, but normalize them into a consistent "headline + takeaway + evidence" format.
- Reduce dense content on detail-heavy slides:
  - Show only active controls where possible.
  - Convert long plain-text values into short chips/badges/labels.
  - Prefer summarized scope text over raw verbose strings.
- Add repeated context on major analytical slides:
  - tenant/report date
  - section label
  - slide number or progress indicator
- Remove emoji-based decorative elements and replace them with shapes/icons that render consistently in PowerPoint.

### 2. Introduce a themeable presentation system
Refactor the current hardcoded styling in `skill/generate_report.js` into a configurable theme layer:

- Add a single theme object or external theme config containing:
  - palette
  - typography
  - spacing scale
  - border radii
  - accent mappings for severity/state
  - optional logo/title metadata
- Replace direct use of `Calibri` and `Calibri Light` with theme-driven font tokens.
- Standardize component primitives:
  - slide title
  - section subtitle
  - stat badge
  - callout box
  - insight card
  - appendix divider
  - table wrapper
- Define one default premium client-facing theme as the shipped default.
- Support future alternate themes without code branching across slide functions.
- Keep color semantics stable:
  - green = positive/enforced
  - amber = caution/report-only
  - red = critical/gap
  - neutral = informational/background

### 3. Make the layout adaptive instead of fixed-size brittle
Refactor slide composition rules so realistic tenant data does not overflow or become unreadable:

- Add helper logic for content-length-aware layout decisions:
  - split long tables across multiple slides
  - trim card counts per slide and continue onto sibling slides
  - cap visible text and route overflow to continuation or appendix
- Replace fixed `autoPage: false` table assumptions with deterministic pagination helpers.
- Add explicit max-content rules per slide type:
  - max rows
  - max cards
  - max line length before truncation or continuation
- Use a shared text sanitization/normalization layer before rendering:
  - humanize long identifiers
  - shorten repeated phrases
  - resolve GUIDs and known app labels before slide placement
- Add a slide footer/header utility so appendix slides and analytical slides remain visually consistent.

### 4. Expand the analysis contract only where needed for presentation
Update the analysis schema and skill instructions only for presentation-critical fields. Extend `skill/analysis-schema.md` and `skill/SKILL.md` to support the new narrative deck:

- Add executive-facing fields:
  - `assessment.score` or `assessment.level`
  - `executiveSummary.topPriorities`
  - `roadmap` grouped into near-term / mid-term actions
  - `meta.clientName` optional
  - `meta.logoPath` or theme-driven logo reference optional
- Keep existing findings/risk/recommendation structures unless a new slide needs explicit data.
- Update the skill instructions so the analysis step produces:
  - concise slide-safe summaries
  - recommendation phrasing suitable for executive audiences
  - roadmap grouping rules derived from priority and effort assumptions
- Preserve backwards compatibility where possible:
  - if new optional fields are missing, the generator should fall back cleanly
  - no requirement to rewrite older example data immediately beyond minimum compatibility updates

### 5. Productize the repository
Turn the repo into a presentable project root rather than a skill folder with supporting files:

- Add a root `README.md` that includes:
  - one-sentence product pitch
  - who it is for
  - what inputs it accepts
  - what output it generates
  - 3-5 sample slide thumbnails
  - quick start
  - sample output references
- Keep the skill-level README, but make the root README the main landing page.
- Expand `package.json` with:
  - `name`
  - `version`
  - `description`
  - `private`
  - `scripts` for `generate`, `generate:example`, `qa:render`, and `qa:slides`
- Add a sanitized sample bundle generated from `skill/examples/analysis-example.json`:
  - sample PPTX
  - sample PDF if generation is available in project workflow
  - slide thumbnails or preview images committed under a predictable `examples/` or `docs/` path
- Document the sample generation path so a maintainer can refresh showcase assets intentionally.

### 6. Add repeatable visual QA
Convert the current manual QA guidance into a defined workflow:

- Add a script or documented command chain to:
  - generate the sample deck
  - convert to PDF/images when local tooling exists
  - output slide previews to a known directory
- Define acceptance checks for presentation:
  - no cut-off titles
  - no table overflow
  - no text overlapping callouts
  - minimum readable text sizes for non-appendix slides
  - consistent footer/header placement
  - no emoji/glyph fallback artifacts
- Keep QA non-destructive and optional when external tools such as LibreOffice are unavailable, but make the "best path" explicit.

## Public Interfaces / Contract Changes
The implementation should treat these as the only planned interface-level changes:

- `analysis.json` gains optional presentation fields for scorecard, top priorities, roadmap, and optional client branding metadata.
- `package.json` gains runnable scripts for generation and QA.
- The repo gains a root `README.md` as the primary public entrypoint.
- The generator continues to consume `analysis.json` + `policies.json`; no change to the core input model for CA policy data.

## Test Plan
Implementation is complete only if these scenarios pass:

- Generate a deck from the sanitized example data and verify the new executive-first structure appears before appendix content.
- Generate with minimal data:
  - policies only
  - no optional enrichment
  - no new optional presentation fields
  - deck still renders cleanly with sensible fallbacks.
- Generate with "full" data shape:
  - geolocation
  - MFA
  - risk
  - auth strengths
  - PIM
  - all optional sections render without overlap.
- Validate adaptive overflow behavior:
  - many categories
  - long policy names
  - long scope strings
  - many report-only items
  - many appendix rows
- Confirm the sample bundle can be regenerated from sanitized inputs and matches the documented workflow.
- Validate repo presentation:
  - root README stands alone as product documentation
  - screenshots/thumbnails exist and correspond to the new deck
  - `npm` scripts cover normal usage and showcase generation.

## Assumptions and Defaults
- Primary audience is client-facing stakeholders, not internal operators.
- The redesign should be substantial but not a ground-up product rewrite.
- Detailed policy inventory remains in the report, but behind a clear appendix boundary.
- A themeable system is required; the shipped default should feel premium and client-ready.
- The repo should include a full sanitized sample bundle, not just copy-only documentation.
- Existing CA analysis logic is preserved unless a new executive-facing slide requires small schema additions.
