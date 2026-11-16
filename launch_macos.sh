#!/bin/zsh

APP_DIR="${0:A:h}"
APP_SCRIPT="$APP_DIR/condensate_analysis_app.py"
APP_LOG="$APP_DIR/fiji_pcc_manders_app.log"
DEPENDENCY_LOG="$APP_DIR/dependency_setup.log"
GUI_MODE=0
[[ "${1:-}" == "--gui" ]] && GUI_MODE=1

show_error() {
  echo "$1"
  if (( GUI_MODE )); then
    /usr/bin/osascript -e "display alert \"Condensate Analysis App\" message \"$1\" as critical" >/dev/null 2>&1
  fi
}

python_is_supported() {
  "$1" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)' >/dev/null 2>&1
}

find_python() {
  local candidate
  if [ -x "$APP_DIR/.venv_macos/bin/python3" ]; then
    echo "$APP_DIR/.venv_macos/bin/python3"
    return 0
  fi
  for candidate in \
    "${commands[python3]:-}" \
    /opt/homebrew/bin/python3 \
    /usr/local/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/Current/bin/python3 \
    /usr/bin/python3; do
    if [ -n "$candidate" ] && [ -x "$candidate" ] && python_is_supported "$candidate"; then
      echo "$candidate"
      return 0
    fi
  done
  return 1
}

cd "$APP_DIR" || exit 1
if [ ! -f "$APP_SCRIPT" ]; then
  show_error "The main app file is missing. Copy the complete app folder again."
  exit 1
fi

PYTHON_BIN="$(find_python)" || {
  show_error "Python 3.11 or newer was not found. Install Python, then run Setup macOS.command."
  exit 1
}

# Read the build identifier from the app being launched. This prevents a
# repaired app from being rejected by an older hard-coded launcher and routed
# back to a stale local server.
EXPECTED_BUILD=$("$PYTHON_BIN" -c '
import ast, pathlib, sys
tree = ast.parse(pathlib.Path(sys.argv[1]).read_text(encoding="utf-8"))
for node in tree.body:
    if isinstance(node, (ast.Assign, ast.AnnAssign)):
        targets = node.targets if isinstance(node, ast.Assign) else [node.target]
        if any(isinstance(target, ast.Name) and target.id == "APP_BUILD_ID" for target in targets):
            print(ast.literal_eval(node.value))
            break
else:
    raise SystemExit("APP_BUILD_ID is missing")
' "$APP_SCRIPT") || {
  show_error "Could not identify the installed app build. Copy the complete app folder again."
  exit 1
}

if ! "$PYTHON_BIN" -c 'import numpy, pandas, tifffile, imagecodecs, PIL, scipy, skimage, openpyxl' >>"$DEPENDENCY_LOG" 2>&1; then
  show_error "Required Python packages are missing. Double-click Setup macOS.command, then reopen the app."
  exit 1
fi

open_app_url() {
  "$PYTHON_BIN" -c '
import importlib.util, sys
spec = importlib.util.spec_from_file_location("condensate_app", sys.argv[1])
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
module.open_local_browser(sys.argv[2])
' "$APP_SCRIPT" "$1" >>"$APP_LOG" 2>&1
}

cleanup_duplicate_servers() {
  local keep_pid="$1"
  local pid cwd
  for pid in $(/usr/sbin/lsof -tiTCP -sTCP:LISTEN 2>/dev/null); do
    [ -n "$pid" ] || continue
    [ "$pid" = "$keep_pid" ] && continue
    cwd=$(/usr/sbin/lsof -a -p "$pid" -d cwd -Fn 2>/dev/null | /usr/bin/sed -n 's/^n//p')
    if [ "$cwd" = "$APP_DIR" ]; then
      /bin/kill -TERM "$pid" 2>/dev/null
    fi
  done
}

health_matches_current_app() {
  local url="$1"
  local payload
  payload=$(/usr/bin/curl -fsS --max-time 2 "$url/health" 2>/dev/null) || return 1
  echo "$payload" | /usr/bin/grep -Fq "\"state_file\": \"$APP_DIR/review_state.json\"" || return 1
  echo "$payload" | /usr/bin/grep -Fq "\"build\": \"$EXPECTED_BUILD\""
}

if [ -f "$APP_DIR/.server_port" ]; then
  RUNNING_PORT=$(/bin/cat "$APP_DIR/.server_port" 2>/dev/null)
  case "$RUNNING_PORT" in
    ''|*[!0-9]*) ;;
    *)
      RUNNING_URL="http://127.0.0.1:$RUNNING_PORT"
      if health_matches_current_app "$RUNNING_URL"; then
        KEEP_PID=$(/usr/sbin/lsof -tiTCP:"$RUNNING_PORT" -sTCP:LISTEN 2>/dev/null | /usr/bin/head -n 1)
        cleanup_duplicate_servers "$KEEP_PID"
        open_app_url "$RUNNING_URL"
        exit 0
      fi
      ;;
  esac
fi

URL="http://127.0.0.1:18818"
if health_matches_current_app "$URL"; then
  KEEP_PID=$(/usr/sbin/lsof -tiTCP:18818 -sTCP:LISTEN 2>/dev/null | /usr/bin/head -n 1)
  cleanup_duplicate_servers "$KEEP_PID"
  open_app_url "$URL"
  exit 0
fi

cleanup_duplicate_servers ""

echo "[$(/bin/date '+%Y-%m-%d %H:%M:%S')] Starting with $PYTHON_BIN" >>"$APP_LOG"
export CONDENSATE_OPEN_BROWSER="${CONDENSATE_OPEN_BROWSER:-1}"

if (( GUI_MODE )); then
  exec "$PYTHON_BIN" "$APP_SCRIPT" 18818 </dev/null >>"$APP_LOG" 2>&1
fi

echo "Starting Condensate Analysis App. Your browser will open automatically."
echo "Keep this Terminal window open while using the app."
exec "$PYTHON_BIN" "$APP_SCRIPT" 18818
