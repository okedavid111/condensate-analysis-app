#!/bin/zsh

APP_DIR="${0:A:h}"
cd "$APP_DIR" || exit 1

find_system_python() {
  local candidate
  for candidate in \
    "${commands[python3]:-}" \
    /opt/homebrew/bin/python3 \
    /usr/local/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/Current/bin/python3 \
    /usr/bin/python3; do
    if [ -n "$candidate" ] && [ -x "$candidate" ] && \
      "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)' >/dev/null 2>&1; then
      echo "$candidate"
      return 0
    fi
  done
  return 1
}

PYTHON_BIN="$(find_system_python)" || {
  echo "Python 3.11 or newer is required."
  echo "Install it from https://www.python.org/downloads/ and run this setup again."
  read -r "?Press Return to close."
  exit 1
}

echo "Using $PYTHON_BIN"
echo "Creating a private macOS Python environment inside the app folder..."
"$PYTHON_BIN" -m venv "$APP_DIR/.venv_macos" || exit 1
cd /private/tmp || exit 1
"$APP_DIR/.venv_macos/bin/python3" -m pip install --upgrade pip || {
  echo "Python package setup failed. Check the internet connection and macOS privacy permissions, then try again."
  read -r "?Press Return to close."
  exit 1
}
"$APP_DIR/.venv_macos/bin/python3" -m pip install -r "$APP_DIR/requirements.txt" || {
  echo "Package installation failed. Check the internet connection and macOS privacy permissions, then try again."
  read -r "?Press Return to close."
  exit 1
}
"$APP_DIR/.venv_macos/bin/python3" -c 'import numpy, pandas, tifffile, imagecodecs, PIL, scipy, skimage, openpyxl' || exit 1

echo
echo "Setup completed successfully."
echo "You can now double-click OPEN CONDENSATE ANALYSIS APP.command"
read -r "?Press Return to close."
