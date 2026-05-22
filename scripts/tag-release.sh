#!/usr/bin/env bash
# Create annotated git tag for a ppt-academizer release (local only; no push).
set -euo pipefail
VER="${1:?Usage: tag-release.sh 1.0.0}"
TAG="ppt-academizer-v${VER}"
ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
RELEASE="apps/ppt-academizer/releases/v${VER}-RELEASE.md"

if [ ! -f "$ROOT/$RELEASE" ]; then
  echo "Missing $ROOT/$RELEASE" >&2
  exit 1
fi

cd "$ROOT"
if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "No git repo at $ROOT — init first:" >&2
  echo "  cd $ROOT && git init && git add … && git commit -m 'baseline'" >&2
  exit 1
fi

git tag -a "$TAG" -m "ppt-academizer ${VER} — see ${RELEASE}"
HASH="$(git rev-parse "$TAG")"
echo "Tagged $TAG at $HASH"
echo "Update Git section in $RELEASE with: $HASH"
