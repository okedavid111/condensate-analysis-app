#!/bin/zsh

APP_DIR="${0:A:h}"
APP_SCRIPT="$APP_DIR/condensate_analysis_app.py"
RECOVERY_PORT=64721

find_python() {
  local candidate
  for candidate in \
    "$APP_DIR/.venv_macos/bin/python3" \
    /usr/local/bin/python3 \
    /opt/homebrew/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/Current/bin/python3 \
    /usr/bin/python3; do
    if [ -x "$candidate" ] && "$candidate" -c 'import numpy, pandas, tifffile, PIL, scipy, skimage, openpyxl' >/dev/null 2>&1; then
      echo "$candidate"
      return 0
    fi
  done
  return 1
}

echo "Preserving the open ROI drawing page on port $RECOVERY_PORT..."
for pid in $(/usr/sbin/lsof -tiTCP:$RECOVERY_PORT -sTCP:LISTEN 2>/dev/null); do
  case "$pid" in
    ''|*[!0-9]*) ;;
    *) /bin/kill -TERM "$pid" 2>/dev/null ;;
  esac
done

for attempt in {1..20}; do
  /usr/sbin/lsof -tiTCP:$RECOVERY_PORT -sTCP:LISTEN >/dev/null 2>&1 || break
  /bin/sleep 0.1
done

if /usr/sbin/lsof -tiTCP:$RECOVERY_PORT -sTCP:LISTEN >/dev/null 2>&1; then
  echo "Could not stop the restricted server. Close this window and contact Codex."
  read -r "?Press Return to close."
  exit 1
fi

PYTHON_BIN="$(find_python)" || {
  echo "A compatible Python installation was not found."
  read -r "?Press Return to close."
  exit 1
}

cd "$APP_DIR" || exit 1
export CONDENSATE_OPEN_BROWSER=0
echo "Recovery server is ready when you see the local address below."
echo "Leave this Terminal window open, return to the existing browser tab, and click Save ROI + condensates + backgrounds."
exec "$PYTHON_BIN" "$APP_SCRIPT" "$RECOVERY_PORT"
