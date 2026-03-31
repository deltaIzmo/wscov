#!/usr/bin/env bash

set -u
set -o pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

usage() {
  cat <<EOF
Usage: tools/wscov_wsl.sh <input.wsc> <testsDir> <outDir> [componentId]

Runs WSCOV from WSL by copying the minimum required files into a Windows-local
workspace, executing tools\\wscov_apply_target.cmd there, and copying outputs
back into the requested WSL outDir.

Environment:
  WSCOV_WINDOWS_WORKROOT  Optional managed Windows workspace root.
                          Accepts either a Windows path (C:\\...) or a WSL path
                          (/mnt/c/...). Default: %USERPROFILE%\\wscov-work\\wscov
EOF
}

log() {
  printf '%s\n' "[wscov-wsl] $*"
}

warn() {
  printf '%s\n' "[wscov-wsl] WARNING: $*" >&2
}

fail() {
  printf '%s\n' "[wscov-wsl] ERROR: $*" >&2
  exit 2
}

report_error() {
  printf '%s\n' "[wscov-wsl] ERROR: $*" >&2
  return 1
}

require_command() {
  local command_name="$1"
  if ! command -v "$command_name" >/dev/null 2>&1; then
    fail "Required command not found: $command_name"
  fi
}

trim_cr() {
  printf '%s' "$1" | tr -d '\r'
}

escape_for_bash_double_quotes() {
  local value="$1"
  value="${value//\\/\\\\}"
  value="${value//\"/\\\"}"
  value="${value//\$/\\\$}"
  value="${value//\`/\\\`}"
  printf '%s' "$value"
}

to_abs_existing_file() {
  local raw_path="$1"
  local abs_path

  abs_path="$(realpath "$raw_path" 2>/dev/null)" || { report_error "Input WSC not found: $raw_path"; return 1; }
  [ -f "$abs_path" ] || { report_error "Input WSC is not a file: $raw_path"; return 1; }
  printf '%s' "$abs_path"
}

to_abs_existing_dir() {
  local raw_path="$1"
  local abs_path

  abs_path="$(realpath "$raw_path" 2>/dev/null)" || { report_error "Tests directory not found: $raw_path"; return 1; }
  [ -d "$abs_path" ] || { report_error "Tests directory is not a folder: $raw_path"; return 1; }
  printf '%s' "$abs_path"
}

prepare_out_dir() {
  local raw_path="$1"
  local abs_path

  mkdir -p "$raw_path" || { report_error "Failed to create outDir: $raw_path"; return 1; }
  abs_path="$(realpath "$raw_path" 2>/dev/null)" || { report_error "Failed to resolve outDir: $raw_path"; return 1; }
  [ -d "$abs_path" ] || { report_error "outDir is not a folder: $raw_path"; return 1; }
  printf '%s' "$abs_path"
}

to_windows_path() {
  local wsl_path="$1"
  local windows_path

  windows_path="$(wslpath -w "$wsl_path" 2>/dev/null)" || { report_error "Failed to convert WSL path to Windows path: $wsl_path"; return 1; }
  trim_cr "$windows_path"
}

to_wsl_path() {
  local windows_path="$1"
  local wsl_path

  wsl_path="$(wslpath -u "$windows_path" 2>/dev/null)" || { report_error "Failed to convert Windows path to WSL path: $windows_path"; return 1; }
  trim_cr "$wsl_path"
}

default_windows_workroot() {
  local userprofile=""
  local userprofile_wsl=""
  local fallback_win=""

  if userprofile="$(cmd.exe /d /c echo %USERPROFILE% 2>/dev/null)"; then
    userprofile="$(trim_cr "$userprofile")"
    if [ -n "$userprofile" ]; then
      printf '%s\\wscov-work\\wscov' "$userprofile"
      return 0
    fi
  fi

  userprofile_wsl="/mnt/c/Users/$USER"
  if [ -d "$userprofile_wsl" ]; then
    fallback_win="$(to_windows_path "$userprofile_wsl")" || return 1
    printf '%s\\wscov-work\\wscov' "$fallback_win"
    return 0
  fi

  report_error "Failed to determine the default Windows workspace. Set WSCOV_WINDOWS_WORKROOT explicitly."
  return 1
}

resolve_windows_workroot() {
  local raw_value="$1"
  local normalized_wsl=""
  local normalized_win=""

  if [ -z "$raw_value" ]; then
    raw_value="$(default_windows_workroot)" || return 1
  fi

  case "$raw_value" in
    *://*)
      report_error "URL-style work roots are not supported: $raw_value"
      return 1
      ;;
    \\\\*)
      report_error "UNC work roots are not supported: $raw_value"
      return 1
      ;;
    /mnt/*)
      normalized_wsl="$(realpath -m "$raw_value")"
      normalized_win="$(to_windows_path "$normalized_wsl")" || return 1
      ;;
    /?*)
      report_error "WSL work roots must live under /mnt/<drive>/: $raw_value"
      return 1
      ;;
    [A-Za-z]:\\*|[A-Za-z]:/*)
      normalized_win="$raw_value"
      normalized_wsl="$(to_wsl_path "$normalized_win")" || return 1
      normalized_wsl="$(realpath -m "$normalized_wsl")"
      normalized_win="$(to_windows_path "$normalized_wsl")" || return 1
      ;;
    *)
      report_error "Work root must be an absolute Windows path or an absolute /mnt path: $raw_value"
      return 1
      ;;
  esac

  case "$normalized_wsl" in
    /mnt/[A-Za-z]/*)
      ;;
    *)
      report_error "Resolved Windows workspace must live under /mnt/<drive>/: $normalized_wsl"
      return 1
      ;;
  esac

  WSCOV_WORKROOT_WSL="$normalized_wsl"
  WSCOV_WORKROOT_WIN="$normalized_win"
}

copy_tree_contents() {
  local source_dir="$1"
  local dest_dir="$2"

  mkdir -p "$dest_dir" || fail "Failed to create directory: $dest_dir"
  cp -a "$source_dir/." "$dest_dir/" || fail "Failed to copy $source_dir to $dest_dir"
}

build_fallback_workspace_wsl() {
  local preferred_root_wsl="$1"
  local preferred_parent_wsl=""
  local runs_parent_wsl=""
  local timestamp=""
  local candidate=""
  local suffix=0

  preferred_parent_wsl="$(dirname "$preferred_root_wsl")"
  runs_parent_wsl="$preferred_parent_wsl/runs"
  timestamp="$(date +%Y%m%d-%H%M%S)"
  candidate="$runs_parent_wsl/wscov-$timestamp-$$"

  while [ -e "$candidate" ]; do
    suffix=$((suffix + 1))
    candidate="$runs_parent_wsl/wscov-$timestamp-$$-$suffix"
  done

  printf '%s' "$candidate"
}

report_workspace_prep_detail() {
  local detail="$1"
  local line=""

  [ -n "$detail" ] || return 0

  while IFS= read -r line; do
    [ -n "$line" ] || continue
    warn "$line"
  done <<EOF
$detail
EOF
}

try_prepare_workspace_root() {
  local root_wsl="$1"
  local tools_dir="$root_wsl/tools"
  local run_dir="$root_wsl/run"
  local sut_dir="$run_dir/sut"
  local tests_dir="$run_dir/tests"
  local out_dir="$run_dir/out"
  local output=""

  WORKSPACE_PREP_ERROR=""

  output="$(rm -rf "$tools_dir" "$run_dir" 2>&1)" || {
    WORKSPACE_PREP_ERROR="$output"
    return 1
  }

  output="$(mkdir -p "$tools_dir" "$sut_dir" "$tests_dir" "$out_dir" 2>&1)" || {
    WORKSPACE_PREP_ERROR="$output"
    return 1
  }

  return 0
}

select_managed_workspace() {
  local preferred_root_wsl="$1"
  local preferred_root_win="$2"
  local fallback_root_wsl=""
  local fallback_root_win=""

  WORKSPACE_PREP_ERROR=""
  PREFERRED_WORKROOT_WSL="$preferred_root_wsl"
  PREFERRED_WORKROOT_WIN="$preferred_root_win"

  log "Preferred Windows workspace: $PREFERRED_WORKROOT_WIN"
  if try_prepare_workspace_root "$PREFERRED_WORKROOT_WSL"; then
    ACTUAL_WORKROOT_WSL="$PREFERRED_WORKROOT_WSL"
    ACTUAL_WORKROOT_WIN="$PREFERRED_WORKROOT_WIN"
    log "Using Windows workspace: $ACTUAL_WORKROOT_WIN"
    return 0
  fi

  warn "Preferred Windows workspace could not be reset. Falling back to a fresh workspace."
  report_workspace_prep_detail "$WORKSPACE_PREP_ERROR"

  fallback_root_wsl="$(build_fallback_workspace_wsl "$PREFERRED_WORKROOT_WSL")"
  fallback_root_win="$(to_windows_path "$fallback_root_wsl")" || exit 2

  if ! try_prepare_workspace_root "$fallback_root_wsl"; then
    report_workspace_prep_detail "$WORKSPACE_PREP_ERROR"
    fail "Failed to prepare fallback workspace under $fallback_root_wsl"
  fi

  ACTUAL_WORKROOT_WSL="$fallback_root_wsl"
  ACTUAL_WORKROOT_WIN="$fallback_root_win"
  log "Using Windows workspace: $ACTUAL_WORKROOT_WIN"
}

if [ "$#" -ne 3 ] && [ "$#" -ne 4 ]; then
  usage >&2
  exit 2
fi

require_command realpath
require_command wslpath
require_command cmd.exe

INPUT_WSC="$(to_abs_existing_file "$1")" || exit 2
TESTS_DIR="$(to_abs_existing_dir "$2")" || exit 2
OUT_DIR="$(prepare_out_dir "$3")" || exit 2
COMPONENT_ID="${4-}"

resolve_windows_workroot "${WSCOV_WINDOWS_WORKROOT-}" || exit 2

select_managed_workspace "$WSCOV_WORKROOT_WSL" "$WSCOV_WORKROOT_WIN"

MANAGED_TOOLS_DIR="$ACTUAL_WORKROOT_WSL/tools"
MANAGED_RUN_DIR="$ACTUAL_WORKROOT_WSL/run"
MANAGED_SUT_DIR="$MANAGED_RUN_DIR/sut"
MANAGED_TESTS_DIR="$MANAGED_RUN_DIR/tests"
MANAGED_OUT_DIR="$MANAGED_RUN_DIR/out"

copy_tree_contents "$ROOT_DIR/tools" "$MANAGED_TOOLS_DIR"

STAGED_INPUT_WSC="$MANAGED_SUT_DIR/$(basename "$INPUT_WSC")"
cp -a "$INPUT_WSC" "$STAGED_INPUT_WSC" || fail "Failed to copy input WSC into managed workspace"
copy_tree_contents "$TESTS_DIR" "$MANAGED_TESTS_DIR"

APPLY_TARGET_WIN="$(to_windows_path "$MANAGED_TOOLS_DIR/wscov_apply_target.cmd")" || exit 2
STAGED_INPUT_WIN="$(to_windows_path "$STAGED_INPUT_WSC")" || exit 2
STAGED_TESTS_WIN="$(to_windows_path "$MANAGED_TESTS_DIR")" || exit 2
STAGED_OUT_WIN="$(to_windows_path "$MANAGED_OUT_DIR")" || exit 2

APPLY_TARGET_CMD_ESCAPED="$(escape_for_bash_double_quotes "$APPLY_TARGET_WIN")"
STAGED_INPUT_ESCAPED="$(escape_for_bash_double_quotes "$STAGED_INPUT_WIN")"
STAGED_TESTS_ESCAPED="$(escape_for_bash_double_quotes "$STAGED_TESTS_WIN")"
STAGED_OUT_ESCAPED="$(escape_for_bash_double_quotes "$STAGED_OUT_WIN")"

CMD_LINE="\"\"$APPLY_TARGET_CMD_ESCAPED\" \"$STAGED_INPUT_ESCAPED\" \"$STAGED_TESTS_ESCAPED\" \"$STAGED_OUT_ESCAPED\""
if [ -n "$COMPONENT_ID" ]; then
  COMPONENT_ID_ESCAPED="$(escape_for_bash_double_quotes "$COMPONENT_ID")"
  CMD_LINE="$CMD_LINE \"$COMPONENT_ID_ESCAPED\""
fi
CMD_LINE="$CMD_LINE\""

log "Running Windows test flow via tools\\wscov_apply_target.cmd"
(
  cd "$ACTUAL_WORKROOT_WSL" || fail "Failed to enter workspace: $ACTUAL_WORKROOT_WSL"
  bash -lc "cmd.exe /d /c $CMD_LINE"
)
WINDOWS_RC=$?

log "Copying outputs back to: $OUT_DIR"
copy_tree_contents "$MANAGED_OUT_DIR" "$OUT_DIR"

if [ "$WINDOWS_RC" -eq 0 ]; then
  log "Completed successfully."
else
  log "Windows test flow exited with code $WINDOWS_RC."
fi

exit "$WINDOWS_RC"
