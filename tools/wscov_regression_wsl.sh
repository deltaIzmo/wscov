#!/usr/bin/env bash

set -u
set -o pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RUNNER="$ROOT_DIR/tools/wscov_wsl.sh"

usage() {
  cat <<EOF
Usage: tools/wscov_regression_wsl.sh

Runs the built-in sample regressions from WSL by delegating to
tools/wscov_wsl.sh for each suite.
EOF
}

log() {
  printf '%s\n' "[wscov-regression-wsl] $*"
}

fail() {
  printf '%s\n' "[wscov-regression-wsl] ERROR: $*" >&2
  exit 2
}

run_suite() {
  local suite_name="$1"
  local input_wsc="$2"
  local tests_dir="$3"
  local out_dir="$4"
  local component_id="$5"

  log "regression: $suite_name"
  "$RUNNER" "$input_wsc" "$tests_dir" "$out_dir" "$component_id"
}

if [ "$#" -ne 0 ]; then
  usage >&2
  exit 2
fi

[ -x "$RUNNER" ] || fail "Expected executable runner not found: $RUNNER"

run_suite \
  "calculator" \
  "$ROOT_DIR/samples/sut/Calculator.wsc" \
  "$ROOT_DIR/samples/tests/calculator" \
  "$ROOT_DIR/out/wsl/regression/calculator" \
  "Calculator" || exit $?

run_suite \
  "branchy" \
  "$ROOT_DIR/samples/sut/Branchy.wsc" \
  "$ROOT_DIR/samples/tests/branchy" \
  "$ROOT_DIR/out/wsl/regression/branchy" \
  "Branchy" || exit $?

log "regression passed"
exit 0
