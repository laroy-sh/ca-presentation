# CA Presentation Skill

This folder contains the analysis contract and PowerPoint generator used by CA Presentation.

## Files

```text
skill/
├── SKILL.md
├── analysis-schema.md
├── theme.default.js
├── generate_report.js
├── scripts/
│   ├── qa-render.sh
│   └── qa-slides.js
└── examples/
    ├── analysis-example.json
    └── policies-example.json
```

## Generator

`generate_report.js` renders an executive-first deck with:
- cover, agenda, scorecard, executive summary, priorities, and roadmap
- supporting analysis slides
- appendix divider, full policy matrix, and per-policy detail

It supports optional CLI arguments:

```bash
node skill/generate_report.js \
  --analysis analysis.json \
  --policies policies.json \
  --output CA_Security_Posture_Report.pptx \
  --theme skill/theme.default.js
```

## Theme System

`theme.default.js` defines:
- palette
- typography
- spacing/radius tokens
- semantic accent mappings (`positive`, `caution`, `critical`, `neutral`)

Provide `--theme` to load an alternate JS/JSON theme file.

## QA Utilities

- `npm run qa:render`:
  - generates example PPTX
  - optionally converts PPTX -> PDF -> JPG previews if local tools exist
- `npm run qa:slides`:
  - runs deterministic checks against sample analysis content limits

## Example Inputs

`skill/examples/analysis-example.json` and `skill/examples/policies-example.json` are sanitized and safe to commit. Use them to regenerate sample outputs.
