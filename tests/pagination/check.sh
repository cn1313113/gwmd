#!/bin/bash
# 分页一致性检查：预览（确定性行模型）vs docx（LibreOffice 渲染 PDF）
# 用法: bash tests/pagination/check.sh [样例.gwmd ...]  默认: 3 个内置样例
set -e
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
DEFAULT="$ROOT/tests/cases.gwmd $ROOT/tests/cases_warnings.gwmd"
# SAMPLE 不在此处（需要 app.js 提取），见 README
FILES="${@:-$DEFAULT}"
fail=0
for gwmd in $FILES; do
  [ -f "$gwmd" ] || { echo "!! 文件不存在: $gwmd"; exit 2; }
  tmpd=$(mktemp -d)
  sudo PYTHONDONTWRITEBYTECODE=1 python3 "$ROOT/gwmd2docx.py" "$gwmd" "$tmpd/out.docx" >/dev/null 2>&1
  soffice --headless --convert-to pdf "$tmpd/out.docx" --outdir "$tmpd" >/dev/null 2>&1
  docx_pages=$(pdfinfo "$tmpd/out.pdf" | awk '/^Pages/{print $2}')
  prev=$(node "$ROOT/tests/pagination/preview.js" "$gwmd" | cut -d' ' -f1)
  pv=${prev#PREVIEW_PAGES:}
  if [ "$docx_pages" = "$pv" ]; then
    echo "PASS $(basename "$gwmd"): docx=$docx_pages 页 / 预览=$pv 页"
  else
    echo "FAIL $(basename "$gwmd"): docx=$docx_pages 页 / 预览=$pv 页"
    fail=1
  fi
  rm -rf "$tmpd"
done
exit $fail
