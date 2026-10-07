#!/usr/bin/env bash
# Usage: ./build.sh https://user.github.io/repo
set -euo pipefail
BASE="${1%/}"
ORIGIN="$(echo "$BASE" | sed -E 's#^(https://[^/]+).*#\1#')"
sed -e "s#{{BASE}}#${BASE}#g" -e "s#{{ORIGIN}}#${ORIGIN}#g" manifest.template.xml > manifest.xml
echo "manifest.xml built for ${BASE}"
