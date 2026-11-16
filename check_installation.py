#!/usr/bin/env python3
"""Verify the private Python environment, Fiji, and JACoP installation."""

import importlib.util
import os
import sys


def fail(message):
    print(f"FAIL: {message}")
    raise SystemExit(1)


try:
    import imagecodecs  # noqa: F401
    import numpy  # noqa: F401
    import openpyxl  # noqa: F401
    import pandas  # noqa: F401
    import PIL  # noqa: F401
    import scipy  # noqa: F401
    import skimage  # noqa: F401
    import tifffile  # noqa: F401
except Exception as exc:
    fail(f"a required Python package could not be imported: {exc}")

print(f"PASS: Python {sys.version.split()[0]} and required packages")

app_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "condensate_analysis_app.py")
spec = importlib.util.spec_from_file_location("condensate_app", app_file)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
fiji_path = module.FIJI_EXECUTABLE
if not os.path.isfile(fiji_path):
    fail(f"Fiji executable was not found: {fiji_path}")

print(f"PASS: Fiji executable: {fiji_path}")
jacop_path = os.path.join(module.fiji_installation_root(), "plugins", "JACoP_.class")
if not os.path.isfile(jacop_path):
    fail(f"JACoP plugin was not found: {jacop_path}")

print(f"PASS: JACoP plugin: {jacop_path}")
print("ALL CHECKS PASSED. The app is ready to open.")
