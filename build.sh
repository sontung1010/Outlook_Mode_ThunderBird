#!/bin/bash
# Build outlook-mode.xpi from the outlook-mode/ source folder.
# Usage: ./build.sh            -> outlook-mode.xpi (LICENSE is added to the package)
set -euo pipefail
cd "$(dirname "$0")"
SRC=outlook-mode
OUT="$PWD/outlook-mode.xpi"

python3 - "$SRC/manifest.json" <<'PY'
import json, sys
json.load(open(sys.argv[1]))  # fail early on a broken manifest
PY

rm -f "$OUT"
(cd "$SRC" && zip -q -r -X "$OUT" . -x '*.DS_Store' -x '*~')
zip -q -X "$OUT" LICENSE
VERSION=$(python3 -c "import json;print(json.load(open('$SRC/manifest.json'))['version'])")
echo "Built $OUT (version $VERSION)"
