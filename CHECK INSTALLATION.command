#!/bin/zsh

APP_DIR="${0:A:h}"
cd "$APP_DIR" || exit 1
PYTHON_BIN="$APP_DIR/.venv_macos/bin/python3"

echo "Condensate Analysis installation check"
echo "Folder: $APP_DIR"
echo

if [ ! -x "$PYTHON_BIN" ]; then
  echo "FAIL: The private Python environment is missing."
  echo "Run Setup macOS.command first."
  read -r "?Press Return to close."
  exit 1
fi

if ! "$PYTHON_BIN" -c 'import numpy, pandas, tifffile, imagecodecs, PIL, scipy, skimage, openpyxl'; then
  echo "FAIL: One or more Python packages are missing."
  echo "Run Setup macOS.command again."
  read -r "?Press Return to close."
  exit 1
fi
echo "PASS: Python and required packages"

FIJI_PATH=$("$PYTHON_BIN" -c '
import importlib.util
spec = importlib.util.spec_from_file_location("condensate_app", "condensate_analysis_app.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
print(module.FIJI_EXECUTABLE)
') || exit 1

if [ ! -x "$FIJI_PATH" ]; then
  echo "FAIL: Fiji executable was not found at:"
  echo "$FIJI_PATH"
  echo "Keep FIJI FOR MAC/Fiji.app beside this app folder or copy it to /Applications."
  read -r "?Press Return to close."
  exit 1
fi
echo "PASS: Fiji executable"
echo "      $FIJI_PATH"

FIJI_ROOT="${FIJI_PATH:h:h:h}"
JACOP_PATH="$FIJI_ROOT/plugins/JACoP_.class"
if [ ! -f "$JACOP_PATH" ]; then
  echo "FAIL: JACoP plugin was not found at:"
  echo "$JACOP_PATH"
  echo "Restore the included Fiji bundle, then run this check again."
  read -r "?Press Return to close."
  exit 1
fi
echo "PASS: JACoP plugin"
echo "      $JACOP_PATH"
echo
echo "ALL CHECKS PASSED. The app is ready to open."
read -r "?Press Return to close."
