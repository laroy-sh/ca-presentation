#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
OUT_DIR="$ROOT_DIR/examples/sample"
PREVIEW_DIR="$ROOT_DIR/docs/previews"
PPTX_PATH="$OUT_DIR/CA_Security_Posture_Report.sample.pptx"
PDF_PATH="$OUT_DIR/CA_Security_Posture_Report.sample.pdf"

mkdir -p "$OUT_DIR"
mkdir -p "$PREVIEW_DIR"

echo "[qa:render] Generating sample PPTX..."
node "$ROOT_DIR/skill/generate_report.js" \
  --analysis "$ROOT_DIR/skill/examples/analysis-example.json" \
  --policies "$ROOT_DIR/skill/examples/policies-example.json" \
  --output "$OUT_DIR/CA_Security_Posture_Report.sample.pptx"

if command -v soffice >/dev/null 2>&1; then
  echo "[qa:render] Converting PPTX to PDF..."
  soffice --headless --convert-to pdf --outdir "$OUT_DIR" "$PPTX_PATH" >/dev/null
  if [[ ! -f "$PDF_PATH" ]]; then
    echo "[qa:render] PDF conversion did not produce expected output path."
  fi
else
  echo "[qa:render] Skipping PDF conversion (soffice not found)."
fi

if command -v pdftoppm >/dev/null 2>&1 && [[ -f "$PDF_PATH" ]]; then
  echo "[qa:render] Rendering JPG slide previews..."
  rm -f "$PREVIEW_DIR"/slide-*.jpg
  pdftoppm -jpeg -r 140 "$PDF_PATH" "$PREVIEW_DIR/slide" >/dev/null
else
  echo "[qa:render] Skipping JPG conversion (pdftoppm missing or PDF unavailable)."
fi

echo "[qa:render] Done."
