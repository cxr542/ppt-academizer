#!/usr/bin/env bash
# Start ppt-academizer web UI (http://127.0.0.1:8765)
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"

_resolve_engine() {
  if [ -n "${PPT_ENGINE_ROOT:-}" ]; then
    echo "$PPT_ENGINE_ROOT"
    return
  fi
  if [ -n "${PPT_TEST_ROOT:-}" ]; then
    echo "$PPT_TEST_ROOT"
    return
  fi
  if [ -f "$ROOT/engine/scripts/academy_deck_build_lib.py" ]; then
    echo "$ROOT/engine"
    return
  fi
  echo "$ROOT/../../cursorstudy/experiments/ppt-test"
}

ENGINE="$(_resolve_engine)"
if [ ! -f "$ENGINE/scripts/academy_deck_build_lib.py" ]; then
  echo "Engine missing. Run: cd $ROOT && python scripts/sync_engine_from_ppt_test.py" >&2
  exit 1
fi

if [ -x "$ROOT/.venv/bin/python" ]; then
  VENV="$ROOT/.venv"
else
  echo "Run: cd $ROOT && python3 -m venv .venv && .venv/bin/pip install -r requirements.txt" >&2
  exit 1
fi

export PYTHONPATH="$ENGINE:$ROOT${PYTHONPATH:+:$PYTHONPATH}"
PORT="${PORT:-8765}"

# Server defaults: no Mac PowerPoint popup; bind localhost. Set PPT_ACADEMIZER_DEV=1 for --reload + 0.0.0.0
export PPT_ACADEMIZER_SKIP_PP_REPAIR="${PPT_ACADEMIZER_SKIP_PP_REPAIR:-1}"

HOST="127.0.0.1"
if [ "${PPT_ACADEMIZER_DEV:-}" = "1" ]; then
  HOST="0.0.0.0"
  echo "Dev mode: reload on, host $HOST (set PPT_ACADEMIZER_DEV=0 for stable server)" >&2
  exec "$VENV/bin/python" -m uvicorn api.main:app --host "$HOST" --port "$PORT" --reload --app-dir "$ROOT"
fi

echo "Server mode: PP repair skipped, host $HOST (PPT_ACADEMIZER_DEV=1 enables reload)" >&2
exec "$VENV/bin/python" -m uvicorn api.main:app --host "$HOST" --port "$PORT" --app-dir "$ROOT"
