#!/usr/bin/env bash
# Self-verification gate for payroll_app. Exit 0 = safe to push.
# Heartbeat, focus-task, and CI all run this. Add a run_step per tool you adopt;
# steps no-op cleanly until their tool exists.
set -uo pipefail
cd "$(dirname "$0")"
PASS=0; FAIL=0; SKIP=0
log_pass(){ echo "  ✓ $1"; PASS=$((PASS+1)); }
log_fail(){ echo "  ✗ $1"; FAIL=$((FAIL+1)); }
log_skip(){ echo "  · $1 [skip: $2]"; SKIP=$((SKIP+1)); }
run_step(){ local n="$1" det="$2"; shift 2
  if eval "$det" >/dev/null 2>&1; then "$@" && log_pass "$n" || { log_fail "$n"; return 1; }
  else log_skip "$n" "not yet introduced"; fi; }

echo "── check.sh ──"

# Backlog integrity (always available).
run_step "task validate" '[ -f tools/task ] && command -v node' node tools/task validate

# Python tests — the real gate for this repo.
run_step "pytest" '[ -d tests ] && command -v python3' bash -c 'python3 -m pytest -q'

echo "── ${PASS} passed, ${FAIL} failed, ${SKIP} skipped ──"
[ "$FAIL" -gt 0 ] && { echo "FAILED — do not commit."; exit 1; }
echo "OK — safe to commit."
