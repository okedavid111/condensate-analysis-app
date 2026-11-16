#!/usr/bin/env python3
"""
Condensate Analysis App

This separate app copy does not modify the original analysis-code folders or the original app.py.
It reviews cropped ROI TIFFs beside the matching raw merged image whose filename contains _M_,
lets you reject unwanted cells/ROIs, then analyzes only the kept crops.
"""

import base64
import copy
import hashlib
import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import webbrowser
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from uuid import uuid4
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, unquote, urlparse

import numpy as np
import pandas as pd
import tifffile as tiff
from scipy import ndimage as ndi
from scipy import stats
from skimage.feature import peak_local_max
from skimage.filters import threshold_otsu, threshold_triangle
from skimage.measure import label, regionprops
from skimage.morphology import closing, dilation, disk, remove_small_objects, white_tophat
from skimage.segmentation import watershed
from openpyxl import load_workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception as exc:  # pragma: no cover
    raise SystemExit("Pillow is required for preview display. Install pillow or use your Python that has PIL.") from exc

APP_ROOT = os.path.dirname(os.path.abspath(__file__))
APP_BUILD_ID = "v1.0.0"
STATE_FILE = os.path.join(APP_ROOT, "review_state.json")
JOB_STATE_FILENAME = "CONDENSATE_REVIEW_STATE_FIXED_CELL_MS2.json"
COUNT_SIZE_THRESHOLDS_FILENAME = "COUNT_SIZE_THRESHOLDS.json"
SERVER_PORT_FILE = os.path.join(APP_ROOT, ".server_port")
STATE_SAVE_LOCK = threading.Lock()
JOB_STATE_LOCK = threading.RLock()
DEFAULT_CROP_ROOT = ""
DEFAULT_OUTPUT_ROOT = ""
DEFAULT_PREPROCESSOR_FOLDER_NAME = "ENHANCED_JACOP_PREPROCESSOR_RESULTS"
DEFAULT_CROP_FOLDER_NAME = "03_cropped_rois_for_final_choice"
DEFAULT_FINAL_OUTPUT_FOLDER_NAME = "REAL FIJI PCC MANDERS RESULTS"
DEFAULT_ANALYSIS_CODE_ROOT = os.path.join(APP_ROOT, "CODE FOR ANALYSIS")
ANALYSIS_CODE_ROOT = os.environ.get("FIJI_PCC_MANDERS_CODE_ROOT", DEFAULT_ANALYSIS_CODE_ROOT)
STAGE_SCRIPT_REL_PATHS = {
    "preprocess": os.path.join("1st code", "enhanced_jacop_preprocessor.py"),
    "crop": os.path.join("2nd code", "crop_raw_and_enhanced_rois_for_jacop.py"),
    "stage3_dapi_red_green": os.path.join("3rd code", "dapi_per_cell_pcc_manders TWO COLOR RED GREEN.py"),
    "stage3_no_dapi_three_color": os.path.join("3rd code", "no_dapi_per_cell_pcc_manders THREE COLOR.py"),
}
RUNNER_DIR = os.path.join(APP_ROOT, "generated_runners")
UPLOAD_ROOT = os.path.join(APP_ROOT, "uploaded_raw_image_folders")
MAX_BROWSER_UPLOAD_BYTES = 256 * 1024 * 1024
def default_fiji_executable():
    configured = os.environ.get("FIJI_EXECUTABLE", "").strip()
    if configured:
        configured = os.path.abspath(os.path.expanduser(configured))
        if os.name == "nt" and os.path.isfile(configured) and os.path.basename(configured).lower().startswith("fiji-windows"):
            imagej_console_launcher = os.path.join(os.path.dirname(configured), "ImageJ-win64.exe")
            if os.path.isfile(imagej_console_launcher):
                return imagej_console_launcher
        if os.path.isdir(configured):
            macos_dir = os.path.join(configured, "Contents", "MacOS")
            machine = os.uname().machine.lower() if sys.platform == "darwin" else ""
            names = (("fiji-macos", "fiji-macos-arm64", "fiji-macos-x64", "ImageJ-macosx")
                     if machine in ("arm64", "aarch64") else
                     ("fiji-macos", "fiji-macos-x64", "ImageJ-macosx", "fiji-macos-arm64"))
            for name in names:
                candidate = os.path.join(macos_dir, name)
                if os.path.isfile(candidate):
                    return candidate
        return configured
    if os.name == "nt":
        program_files = os.environ.get("PROGRAMFILES", r"C:\\Program Files")
        local_roots = (APP_ROOT, os.path.dirname(APP_ROOT))
        bundled_windows_root = os.path.join(os.path.dirname(APP_ROOT), "FIJI FOR WINDOWS", "Fiji.app")
        candidates = [
            *[os.path.join(bundled_windows_root, name)
              for name in ("ImageJ-win64.exe", "fiji-windows-x64.exe", "fiji-windows.exe")],
            *[os.path.join(root, "Fiji.app", name) for root in local_roots
              for name in ("ImageJ-win64.exe", "fiji-windows-x64.exe", "fiji-windows.exe")],
            os.path.join(program_files, "Fiji.app", "ImageJ-win64.exe"),
            os.path.join(program_files, "Fiji.app", "fiji-windows-x64.exe"),
            os.path.join(program_files, "Fiji", "ImageJ-win64.exe"),
        ]
        for candidate in candidates:
            if os.path.isfile(candidate):
                return candidate
        return shutil.which("ImageJ-win64.exe") or shutil.which("fiji-windows-x64.exe") or "ImageJ-win64.exe"
    if sys.platform == "darwin":
        machine = os.uname().machine.lower()
        if machine in ("arm64", "aarch64"):
            launcher_names = ("fiji-macos", "fiji-macos-arm64", "fiji-macos-x64", "ImageJ-macosx")
        else:
            launcher_names = ("fiji-macos", "fiji-macos-x64", "ImageJ-macosx", "fiji-macos-arm64")
        macos_dirs = [
            os.path.join(os.path.dirname(APP_ROOT), "FIJI FOR MAC", "Fiji.app", "Contents", "MacOS"),
            os.path.join(APP_ROOT, "Fiji.app", "Contents", "MacOS"),
            os.path.join(os.path.dirname(APP_ROOT), "Fiji.app", "Contents", "MacOS"),
            "/Applications/Fiji.app/Contents/MacOS",
            "/Applications/Fiji/Fiji.app/Contents/MacOS",
            "/Applications/Fiji/Contents/MacOS",
            os.path.expanduser("~/Applications/Fiji.app/Contents/MacOS"),
            os.path.expanduser("~/Applications/Fiji/Fiji.app/Contents/MacOS"),
        ]
        candidates = [
            os.path.join(folder, name)
            for folder in macos_dirs
            for name in launcher_names
        ]
        for candidate in candidates:
            if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
                return candidate
        return "/Applications/Fiji.app/Contents/MacOS/fiji-macos"
    for name in ("fiji", "ImageJ-linux64", "fiji-linux-x64"):
        candidate = shutil.which(name)
        if candidate:
            return candidate
    return "fiji"


FIJI_EXECUTABLE = default_fiji_executable()


def fiji_installation_root(executable=None):
    """Return the Fiji.app/Fiji folder on both macOS and Windows."""
    executable = os.path.abspath(executable or FIJI_EXECUTABLE)
    executable_folder = os.path.dirname(executable)
    if os.path.basename(executable_folder).lower() == "macos":
        contents_folder = os.path.dirname(executable_folder)
        if os.path.basename(contents_folder).lower() == "contents":
            return os.path.dirname(contents_folder)
    return executable_folder


def windows_direct_fiji_java_command(input_path, result_path, pair_name, channel_a, channel_b):
    root = fiji_installation_root()
    java_root = os.path.join(root, "java", "win64")
    java_candidates = []
    if os.path.isdir(java_root):
        for folder in sorted(os.listdir(java_root), reverse=True):
            candidate = os.path.join(java_root, folder, "bin", "java.exe")
            if os.path.isfile(candidate):
                java_candidates.append(candidate)
    jars_root = os.path.join(root, "jars")
    ij_jars = []
    if os.path.isdir(jars_root):
        ij_jars = sorted(
            [os.path.join(jars_root, name) for name in os.listdir(jars_root) if re.fullmatch(r"ij-[0-9].*\.jar", name)],
            reverse=True,
        )
    source_path = os.path.join(root, "FijiPccMandersDirect.java")
    if not java_candidates or not ij_jars or not os.path.isfile(source_path):
        raise RuntimeError(f"Bundled Windows Fiji Java runner is incomplete in: {root}")
    return [
        java_candidates[0], "-Djava.awt.headless=true", "-cp", ij_jars[0], source_path,
        input_path, result_path, pair_name, channel_a, channel_b,
    ]


def open_local_browser(url):
    """Open the local UI even when the OS has no working default URL handler."""
    if sys.platform == "darwin":
        candidates = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Firefox.app/Contents/MacOS/firefox",
            "/Applications/Safari.app/Contents/MacOS/Safari",
        ]
        for executable in candidates:
            if os.path.isfile(executable):
                try:
                    subprocess.Popen(
                        [executable, url],
                        stdin=subprocess.DEVNULL,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        start_new_session=True,
                    )
                    return
                except OSError:
                    continue
    try:
        webbrowser.open(url, new=1)
    except Exception:
        pass

CHANNELS = {"red": 0, "green": 1, "blue": 2}
BACKGROUND_VALUE = 0
CROP_INDEX_CONTEXT_CACHE = {}
RAW_M_IMAGE_INDEX_CACHE = {}
LEGACY_CROP_MATCH_CACHE = {}
AUXILIARY_PREVIEW_LIMIT = 60
CHANNEL_CLEANING_PREVIEW_WORKERS = 4
if os.name == "nt":
    _windows_local_app_data = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
    THUMBNAIL_CACHE_ROOT = os.path.join(_windows_local_app_data, "CondensateAnalysis", "thumbnail_cache", "app_v1")
else:
    THUMBNAIL_CACHE_ROOT = os.path.join(APP_ROOT, "thumbnail_cache")


class MissingAnalysisScriptError(FileNotFoundError):
    pass


def analysis_code_roots():
    roots = [
        ANALYSIS_CODE_ROOT,
        os.path.join(APP_ROOT, "CODE FOR ANALYSIS"),
    ]
    unique_roots = []
    seen = set()
    for root in roots:
        root = os.path.abspath(os.path.expanduser(str(root or "")))
        if root and root not in seen:
            seen.add(root)
            unique_roots.append(root)
    return unique_roots


def resolve_analysis_script(script_key):
    rel_path = STAGE_SCRIPT_REL_PATHS[script_key]
    tried = []
    for root in analysis_code_roots():
        candidate = os.path.join(root, rel_path)
        tried.append(candidate)
        if os.path.exists(candidate):
            return candidate
    details = "\n".join(f"  - {path}" for path in tried)
    raise MissingAnalysisScriptError(
        "The original analysis script is missing, so this pipeline stage cannot run yet.\n"
        f"Missing script: {rel_path}\n"
        "Put the original 'CODE FOR ANALYSIS' folder next to this app, or set "
        "FIJI_PCC_MANDERS_CODE_ROOT to its location before launching the app.\n"
        f"Paths checked:\n{details}"
    )


def normalize_analysis_input_mode(value):
    value = str(value or "").strip().lower()
    if value in ("auto_cleaned", "channel_cleaned", "cleaned"):
        return "auto_cleaned"
    return "raw"


def auto_clean_reload_scan_signature(state, cleaned_rgb_root):
    """Identify one stable crop/output combination for startup reload scanning."""
    item_rows = []
    for rel, item in sorted(state.get("items", {}).items()):
        if not isinstance(item, dict):
            continue
        item_rows.append(
            "|".join(
                (
                    rel,
                    str(item.get("status", "keep")),
                    str(item.get("crop_choice", "raw")),
                    os.path.abspath(str(item.get("path", "") or "")),
                )
            )
        )
    try:
        cleaned_names = sorted(
            entry.name
            for entry in os.scandir(cleaned_rgb_root)
            if entry.is_file()
            and not entry.name.startswith("._")
            and entry.name.lower().endswith((".tif", ".tiff"))
            and is_real_tiff_file(entry.path)
        )
        directory_mtime_ns = os.stat(cleaned_rgb_root).st_mtime_ns
    except OSError:
        cleaned_names = []
        directory_mtime_ns = 0
    payload = {
        "cleaned_rgb_root": os.path.abspath(cleaned_rgb_root),
        "directory_mtime_ns": directory_mtime_ns,
        "cleaned_names": cleaned_names,
        "items": item_rows,
    }
    return hashlib.sha256(
        json.dumps(payload, sort_keys=True).encode("utf-8")
    ).hexdigest()


def safe_name(name):
    base = os.path.splitext(os.path.basename(str(name)))[0]
    return "".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in base)


def bounded_output_stem(name, max_bytes=64):
    """Return a filesystem-safe, unique stem with room for output suffixes."""
    stem = safe_name(name)
    if len(stem.encode("utf-8")) <= max_bytes:
        return stem
    digest = hashlib.sha256(str(name).encode("utf-8")).hexdigest()[:16]
    suffix = "__" + digest
    prefix_bytes = stem.encode("utf-8")[: max_bytes - len(suffix)]
    while prefix_bytes:
        try:
            prefix = prefix_bytes.decode("utf-8")
            break
        except UnicodeDecodeError:
            prefix_bytes = prefix_bytes[:-1]
    else:
        prefix = "crop"
    prefix = prefix.rstrip(" ._-") or "crop"
    return prefix + suffix


def file_starts_with(path, signatures):
    if not path or not os.path.isfile(path) or os.path.basename(path).startswith("._"):
        return False
    try:
        with open(path, "rb") as f:
            header = f.read(max(len(signature) for signature in signatures))
    except OSError:
        return False
    return any(header.startswith(signature) for signature in signatures)


def is_real_tiff_file(path):
    return file_starts_with(path, (b"II*\x00", b"MM\x00*", b"II+\x00", b"MM\x00+"))


def is_real_png_file(path):
    return file_starts_with(path, (b"\x89PNG\r\n\x1a\n",))


def copy2_unless_same(src, dst):
    if os.path.abspath(src) == os.path.abspath(dst):
        return dst
    try:
        if os.path.exists(dst) and os.path.samefile(src, dst):
            return dst
    except OSError:
        pass
    return shutil.copy2(src, dst)


def load_state():
    defaults = {
        "crop_root": DEFAULT_CROP_ROOT,
        "output_root": DEFAULT_OUTPUT_ROOT,
        "analysis_mode": "red_green",
        "raw_image_folder": "",
        "preprocessor_folder_name": "ENHANCED_JACOP_PREPROCESSOR_RESULTS",
        "crop_folder_name": "03_cropped_rois_for_final_choice",
        "final_output_folder_name": "REAL FIJI PCC MANDERS RESULTS",
        "final_input_choice": "raw_masked",
        "stage3_mode": "dapi_red_green",
        "items": {},
        "last_scan": "",
        "pixel_um": 0.1290271,
        "pcc_red_label": "",
        "pcc_green_label": "",
        "pcc_blue_label": "",
        "analysis_input_mode": "raw",
        "pcc_analysis_input_mode": "raw",
        "count_size_analysis_input_mode": "raw",
        "count_size_particle_size": "10-Infinity",
        "fixed_cell_n1icd_channel": "green",
        "fixed_cell_condensate_label": "N1ICD",
        "fixed_cell_annotation_mode": "manual_required",
        "fixed_cell_ms2_channel": "red",
        "fixed_cell_ms2_label": "MCP/MS2",
        "fixed_cell_dapi_channel": "blue",
        "fixed_cell_color_mode": "three_color_dapi",
        "fixed_cell_modifier_channel": "blue",
        "fixed_cell_modifier_label": "HP1beta",
        "fixed_cell_focus_rule": "brightest",
        "fixed_cell_ms2_min_area_px2": 3,
        "fixed_cell_background_inner_px": 3,
        "fixed_cell_background_outer_px": 12,
        "fixed_cell_replicate_id": "",
        "fixed_cell_condition": "",
        "fixed_cell_intensity_source_mode": "reviewed_raw",
        "fixed_cell_condensate_source": "count_size",
        "partition_target_channel": "red",
        "partition_target_label": "UB-HA",
        "partition_condensate_channel": "red",
        "partition_red_label": "",
        "partition_green_label": "",
        "partition_blue_label": "",
        "partition_condensate_label": "NICD",
        "partition_intensity_source_mode": "untouched_acquisition",
        "partition_boundary_exclusion_px": 1,
        "partition_replicate_id": "",
        "partition_condition": "",
        "red_channel_role": "diffuse_marker",
        "green_channel_role": "condensate",
        "blue_channel_role": "nucleus",
        "pcc_pair_mode": "red_green",
        "cleaning_radius_px": 14,
        "cleaning_sensitivity": 50,
        "cleaning_min_area_px": 3,
        "cleaning_background_removal": 0,
        "cleaning_amount": 100,
        "red_cleaning_amount": 100,
        "green_cleaning_amount": 100,
        "blue_cleaning_amount": 100,
        "per_crop_cleaning_params": {},
    }
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        if isinstance(loaded, dict):
            merged = {**defaults, **loaded}
            legacy_mode = normalize_analysis_input_mode(merged.get("analysis_input_mode", "raw"))
            for key in ("pcc_analysis_input_mode", "count_size_analysis_input_mode"):
                merged[key] = normalize_analysis_input_mode(merged.get(key, legacy_mode))
            merged["analysis_input_mode"] = legacy_mode
            merged["fixed_cell_intensity_source_mode"] = normalize_fixed_cell_intensity_source_mode(
                merged.get("fixed_cell_intensity_source_mode")
            )
            merged["fixed_cell_condensate_source"] = (
                str(merged.get("fixed_cell_condensate_source", "count_size")).strip().lower()
                if str(merged.get("fixed_cell_condensate_source", "count_size")).strip().lower() in ("count_size", "manual")
                else "count_size"
            )
            loaded_items = merged.get("items", {})
            has_decision_evidence = bool(
                isinstance(loaded_items, dict)
                and any(
                    isinstance(item, dict)
                    and (
                        item.get("status", "keep") == "reject"
                        or "explicit_keep" in item
                    )
                    for item in loaded_items.values()
                )
            )
            output_root = str(merged.get("output_root", "") or "")
            review_workbook = find_latest_result_excel(
                os.path.join(output_root, "real_fiji_pcc_manders_analysis"),
                "pcc_manders",
            )
            recovery_signature = ""
            if review_workbook and os.path.isfile(review_workbook):
                recovery_signature = hashlib.sha256(
                    "|".join(
                        (
                            os.path.abspath(review_workbook),
                            str(os.path.getmtime(review_workbook)),
                            str(len(loaded_items) if isinstance(loaded_items, dict) else 0),
                        )
                    ).encode("utf-8")
                ).hexdigest()
            if (
                isinstance(loaded_items, dict)
                and loaded_items
                and not has_decision_evidence
                and recovery_signature
                and merged.get("last_default_decision_recovery_signature") != recovery_signature
            ):
                # Self-heal portable states that were saved immediately after
                # a scan initialized every crop to the default "keep" value.
                # The reviewed workbook is authoritative for decisions and
                # also carries per-crop phenotype/AutoClean recovery metadata.
                merged["_suspend_portable_job_state_save"] = True
                try:
                    reload_existing_results_for_current_folder(merged)
                    merged["last_default_decision_recovery_signature"] = recovery_signature
                finally:
                    merged.pop("_suspend_portable_job_state_save", None)
                save_state(merged)
            # Crop regeneration can archive manual recrops in
            # _previous_crop_runs. Restore those paths before rendering any
            # page, so cached autocleaned images remain linked to the same
            # manually adjusted source crop after an app restart.
            try:
                restored = restore_archived_manual_recrops(merged)
                cleaned_rgb = os.path.join(
                    merged.get("output_root", ""),
                    "channel_cleaned_analysis_inputs",
                    "cleaned_rgb_crops",
                )
                missing_cached = any(
                    not item.get("auto_cleaned_crop_path")
                    or not is_real_tiff_file(item.get("auto_cleaned_crop_path", ""))
                    or os.path.abspath(item.get("auto_cleaned_source_path", "") or "")
                    != os.path.abspath(item.get("path", "") or "")
                    for item in merged.get("items", {}).values()
                    if item.get("status", "keep") == "keep"
                )
                scan_signature = auto_clean_reload_scan_signature(
                    merged,
                    cleaned_rgb,
                )
                previous_signature = merged.get(
                    "last_auto_cleaned_reload_scan_signature",
                    "",
                )
                should_scan = restored or (
                    missing_cached
                    and os.path.isdir(cleaned_rgb)
                    and scan_signature != previous_signature
                )
                reloaded = (
                    reload_auto_cleaned_inputs_for_loaded(merged)
                    if should_scan
                    else 0
                )
                if should_scan:
                    # Persist even when only some crops have auto-cleaned files.
                    # Without this guard, every thumbnail request repeats the
                    # same scan and creates another full state backup.
                    merged["last_auto_cleaned_reload_scan_signature"] = (
                        auto_clean_reload_scan_signature(merged, cleaned_rgb)
                    )
                if restored or should_scan:
                    save_state(merged)
            except Exception:
                pass
            return merged
    return defaults


def save_state(state):
    with STATE_SAVE_LOCK:
        if os.path.exists(STATE_FILE):
            backup_dir = os.path.join(APP_ROOT, "state_backups")
            os.makedirs(backup_dir, exist_ok=True)
            backup_name = "review_state_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".json"
            shutil.copy2(STATE_FILE, os.path.join(backup_dir, backup_name))
        tmp_file = os.path.join(APP_ROOT, f"review_state.{uuid4().hex}.tmp")
        persisted_state = dict(state)
        persisted_state.pop("_suspend_portable_job_state_save", None)
        try:
            with open(tmp_file, "w", encoding="utf-8") as f:
                json.dump(persisted_state, f, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_file, STATE_FILE)
        finally:
            if os.path.exists(tmp_file):
                os.unlink(tmp_file)
        if not state.get("_suspend_portable_job_state_save"):
            try:
                save_portable_job_state(state)
            except Exception:
                # The app-local state must remain usable when an external drive is
                # read-only, disconnected, or temporarily unavailable.
                pass


def portable_job_state_path(state):
    raw_root = os.path.abspath(str(state.get("raw_image_folder", "") or "").strip())
    if not raw_root or not os.path.isdir(raw_root):
        return ""
    output_root = str(state.get("output_root", "") or "").strip()
    if not output_root:
        output_name = str(state.get("final_output_folder_name", DEFAULT_FINAL_OUTPUT_FOLDER_NAME) or DEFAULT_FINAL_OUTPUT_FOLDER_NAME).strip()
        output_root = os.path.join(raw_root, output_name, "raw_masked_analysis")
    return os.path.join(os.path.abspath(output_root), JOB_STATE_FILENAME)


PORTABLE_JOB_GLOBAL_KEYS = (
    "analysis_mode", "stage3_mode", "pixel_um",
    "pcc_red_label", "pcc_green_label", "pcc_blue_label",
    "analysis_input_mode", "pcc_analysis_input_mode", "count_size_analysis_input_mode",
    "count_size_particle_size", "count_size_red", "count_size_green", "count_size_blue",
    "count_size_thresholds", "per_crop_cleaning_params",
    "red_channel_role", "green_channel_role", "blue_channel_role", "pcc_pair_mode",
    "cleaning_radius_px", "cleaning_sensitivity", "cleaning_min_area_px",
    "cleaning_background_removal", "cleaning_amount",
    "red_cleaning_amount", "green_cleaning_amount", "blue_cleaning_amount",
    "fixed_cell_n1icd_channel", "fixed_cell_condensate_label",
    "fixed_cell_annotation_mode", "fixed_cell_ms2_channel", "fixed_cell_ms2_label",
    "fixed_cell_dapi_channel", "fixed_cell_color_mode",
    "fixed_cell_modifier_channel", "fixed_cell_modifier_label",
    "fixed_cell_focus_rule", "fixed_cell_ms2_min_area_px2",
    "fixed_cell_background_inner_px", "fixed_cell_background_outer_px",
    "fixed_cell_replicate_id", "fixed_cell_condition",
    "fixed_cell_intensity_source_mode", "fixed_cell_condensate_source",
    "partition_target_channel", "partition_target_label", "partition_condensate_channel",
    "partition_red_label", "partition_green_label", "partition_blue_label",
    "partition_condensate_label", "partition_intensity_source_mode",
    "partition_boundary_exclusion_px",
    "partition_replicate_id", "partition_condition",
)

PORTABLE_JOB_OUTPUT_KEYS = (
    "last_stage_output", "last_stage_returncode",
    "last_real_fiji_pcc_manders_excel", "last_pcc_manders_excel",
    "last_pcc_manders_preview_root", "last_pcc_manders_approved_root",
    "last_count_size_excel", "last_count_size_output_root", "last_count_size_approved_root",
    "last_fixed_cell_ms2_excel", "last_fixed_cell_ms2_output_root",
    "last_fixed_cell_ms2_preview_root", "last_partition_excel",
    "last_partition_output_root", "last_partition_preview_root",
    "final_preview_root", "last_final_excel",
)

PORTABLE_JOB_ITEM_KEYS = (
    "status", "explicit_keep", "crop_choice", "note", "channel_roles",
    "manual_crop_path", "manual_original_path",
    "auto_cleaned_crop_path", "auto_cleaned_preview_path",
    "auto_cleaned_params_path", "auto_cleaned_source_path",
    "auto_cleaned_source_mtime", "auto_cleaned_generated_at",
    "auto_cleaned_input_mode", "auto_cleaned_channels",
    "auto_cleaned_settings_signature",
    "fixed_cell_manual_annotation", "fixed_cell_manual_annotation_preview_path",
    "partition_manual_annotation", "partition_annotation_preview_path",
)


def rebase_portable_job_path(path, snapshot, state):
    """Rebase a saved absolute job path when the complete experiment was moved."""
    value = str(path or "").strip()
    if not value:
        return ""
    saved_abs = os.path.abspath(value)
    for root_key in ("raw_image_folder", "crop_root", "output_root"):
        old_root = str(snapshot.get(root_key, "") or "").strip()
        new_root = str(state.get(root_key, "") or "").strip()
        if not old_root or not new_root:
            continue
        old_abs = os.path.abspath(old_root)
        try:
            common = os.path.commonpath([saved_abs, old_abs])
        except ValueError:
            continue
        if common == old_abs:
            return os.path.abspath(os.path.join(new_root, os.path.relpath(saved_abs, old_abs)))
    return saved_abs


def rebase_portable_annotation(annotation, snapshot, state):
    if not isinstance(annotation, dict):
        return annotation
    restored = dict(annotation)
    if restored.get("source_path"):
        restored["source_path"] = rebase_portable_job_path(restored["source_path"], snapshot, state)
    return restored


def restore_portable_job_globals(state):
    """Restore job-wide settings/output links from this experiment's own file."""
    path = portable_job_state_path(state)
    if not path or not os.path.isfile(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8") as handle:
            snapshot = json.load(handle)
    except Exception:
        return ""
    if not isinstance(snapshot, dict):
        return ""
    for key in PORTABLE_JOB_GLOBAL_KEYS:
        if key in snapshot:
            state[key] = copy.deepcopy(snapshot[key])
    for key in PORTABLE_JOB_OUTPUT_KEYS:
        if key in snapshot:
            state[key] = rebase_portable_job_path(snapshot[key], snapshot, state)
    return path


def save_portable_job_state(state):
    """Persist the complete workflow with the experiment from its first stage."""
    items = state.get("items", {})
    # Never replace an experiment's recoverable decision snapshot with an
    # empty transient state while a folder is being selected or rescanned.
    if not isinstance(items, dict) or not items:
        return ""
    path = portable_job_state_path(state)
    if not path:
        return ""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    snapshot = dict(state)
    snapshot["portable_job_state_version"] = 2
    snapshot["portable_job_state_saved_at"] = datetime.now().isoformat(timespec="seconds")
    tmp_path = path + "." + uuid4().hex + ".tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(snapshot, f, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    return path


def named_pipeline_paths(state):
    raw = state.get("raw_image_folder", "").strip()
    preprocessor_name = state.get("preprocessor_folder_name", DEFAULT_PREPROCESSOR_FOLDER_NAME).strip() or DEFAULT_PREPROCESSOR_FOLDER_NAME
    crop_name = state.get("crop_folder_name", DEFAULT_CROP_FOLDER_NAME).strip() or DEFAULT_CROP_FOLDER_NAME
    output_name = state.get("final_output_folder_name", DEFAULT_FINAL_OUTPUT_FOLDER_NAME).strip() or DEFAULT_FINAL_OUTPUT_FOLDER_NAME
    preprocessor_root = os.path.join(raw, preprocessor_name)
    state["final_input_choice"] = "raw_masked"
    if crop_name == "01_raw_masked_for_jacop":
        preprocessor_crop_root = os.path.join(preprocessor_root, crop_name)
        nested_crop_root = os.path.join(preprocessor_crop_root, "01_raw_masked_roi_crops_for_jacop")
        crop_root = nested_crop_root if os.path.isdir(nested_crop_root) else os.path.join(preprocessor_root, DEFAULT_CROP_FOLDER_NAME, "01_raw_masked_roi_crops_for_jacop")
    else:
        crop_base = os.path.join(preprocessor_root, crop_name)
        crop_root = os.path.join(crop_base, "01_raw_masked_roi_crops_for_jacop")
    output_root = os.path.join(raw, output_name, "raw_masked_analysis")
    return raw, preprocessor_name, crop_name, output_name, preprocessor_root, crop_root, output_root


def valid_analysis_mode(value, fallback="red_green"):
    return value if value in ("red_green", "three_color") else fallback


CHANNEL_ROLE_OPTIONS = ("condensate", "diffuse_marker", "nucleus", "ignore")


def valid_channel_role(value, fallback="ignore"):
    return value if value in CHANNEL_ROLE_OPTIONS else fallback


def channel_roles_from_state(state):
    return {
        channel_name: valid_channel_role(
            state.get(f"{channel_name}_channel_role"),
            {
                "red": "diffuse_marker",
                "green": "condensate",
                "blue": "nucleus",
            }.get(channel_name, "ignore"),
        )
        for channel_name in CHANNELS
    }


def channel_roles_for_item(state, item):
    roles = channel_roles_from_state(state)
    overrides = item.get("channel_roles", {}) if isinstance(item, dict) else {}
    if isinstance(overrides, dict):
        for channel_name in CHANNELS:
            roles[channel_name] = valid_channel_role(overrides.get(channel_name), roles[channel_name])
    return roles


def set_item_channel_roles(item, data):
    roles = {}
    for channel_name in CHANNELS:
        roles[channel_name] = valid_channel_role(data.get(f"{channel_name}_role"), "ignore")
    item["channel_roles"] = roles
    return roles


def role_label(role):
    return {
        "condensate": "Condensate / puncta",
        "diffuse_marker": "Marker / protein signal",
        "nucleus": "Nucleus / reference",
        "ignore": "Ignore",
    }.get(role, str(role or "ignore"))


def role_select_html(channel_name, selected_role):
    title = channel_name.title()
    options = []
    for value, label in (
        ("condensate", "Condensate / puncta"),
        ("diffuse_marker", "Marker / protein signal"),
        ("nucleus", "Nucleus / reference"),
        ("ignore", "Ignore"),
    ):
        selected = "selected" if selected_role == value else ""
        options.append(f'<option value="{value}" {selected}>{label}</option>')
    return f"""
    <label>{title} role<br>
      <select name="{channel_name}_role">{"".join(options)}</select>
    </label>
    """


def crop_condensate_channels(state, item):
    return [channel_name for channel_name, role in channel_roles_for_item(state, item).items() if role == "condensate"]


def pcc_pair_name(channel_a, channel_b):
    return f"{channel_a.title()}_vs_{channel_b.title()}"


def pcc_pair_specs_for_item(state, item):
    roles = channel_roles_for_item(state, item)
    condensate_channels = [channel for channel, role in roles.items() if role == "condensate"]
    marker_channels = [channel for channel, role in roles.items() if role == "diffuse_marker"]
    pair_mode = valid_pcc_pair_mode(state.get("pcc_pair_mode"), "red_green")
    selected_pair_sets = {
        "red_green": {frozenset(("red", "green"))},
        "red_blue": {frozenset(("red", "blue"))},
        "green_blue": {frozenset(("green", "blue"))},
        "all_three": {
            frozenset(("red", "green")),
            frozenset(("red", "blue")),
            frozenset(("green", "blue")),
        },
    }[pair_mode]
    pairs = []
    for condensate_channel in condensate_channels:
        for marker_channel in marker_channels:
            if condensate_channel == marker_channel:
                continue
            if frozenset((condensate_channel, marker_channel)) not in selected_pair_sets:
                continue
            pairs.append((pcc_pair_name(condensate_channel, marker_channel), condensate_channel, marker_channel, "condensate_marker"))
    for index, channel_a in enumerate(condensate_channels):
        for channel_b in condensate_channels[index + 1:]:
            if frozenset((channel_a, channel_b)) not in selected_pair_sets:
                continue
            pairs.append((pcc_pair_name(channel_a, channel_b), channel_a, channel_b, "condensate_condensate"))
    for index, channel_a in enumerate(marker_channels):
        for channel_b in marker_channels[index + 1:]:
            if frozenset((channel_a, channel_b)) not in selected_pair_sets:
                continue
            pairs.append((pcc_pair_name(channel_a, channel_b), channel_a, channel_b, "marker_marker"))
    return pairs


def condensate_channels_from_state(state):
    roles = channel_roles_from_state(state)
    return {channel_name for channel_name, role in roles.items() if role == "condensate"}


def valid_pcc_pair_mode(value, fallback="red_green"):
    return value if value in ("red_green", "red_blue", "green_blue", "all_three") else fallback


def pcc_pair_specs_from_state(state):
    pair_mode = valid_pcc_pair_mode(state.get("pcc_pair_mode"), "red_green")
    if pair_mode == "all_three":
        return [
            ("Red_vs_Green", "red", "green"),
            ("Red_vs_Blue", "red", "blue"),
            ("Blue_vs_Green", "blue", "green"),
        ]
    if pair_mode == "red_blue":
        return [("Red_vs_Blue", "red", "blue")]
    if pair_mode == "green_blue":
        return [("Blue_vs_Green", "blue", "green")]
    return [("Red_vs_Green", "red", "green")]


def pcc_selected_channels_from_state(state):
    selected = set()
    for _pair_name, channel_a, channel_b in pcc_pair_specs_from_state(state):
        selected.add(channel_a)
        selected.add(channel_b)
    return selected or set(CHANNELS)


def pcc_image_yxc_for_selected_pair(state, image_yxc):
    pair_mode = valid_pcc_pair_mode(state.get("pcc_pair_mode"), "red_green")
    if pair_mode == "all_three":
        return image_yxc
    selected_channels = pcc_selected_channels_from_state(state)
    filtered = np.array(image_yxc, copy=True)
    for channel_name, index in CHANNELS.items():
        if channel_name not in selected_channels:
            filtered[..., index] = 0
    return filtered


def refresh_derived_paths(state):
    state["final_input_choice"] = "raw_masked"
    _, _, _, _, _, crop_root, output_root = named_pipeline_paths(state)
    state["crop_root"] = crop_root
    state["output_root"] = output_root
    state["analysis_mode"] = valid_analysis_mode(state.get("analysis_mode"), "red_green")
    state["stage3_mode"] = "dapi_red_green" if state["analysis_mode"] == "red_green" else "no_dapi_three_color"


def ensure_sibling_output_layout(state):
    raw = state.get("raw_image_folder", "").strip()
    if not raw:
        return state.get("output_root", "")
    canonical_preprocessor = os.path.join(raw, DEFAULT_PREPROCESSOR_FOLDER_NAME)
    canonical_output = os.path.join(raw, DEFAULT_FINAL_OUTPUT_FOLDER_NAME)
    uses_canonical_layout = os.path.isdir(canonical_preprocessor) or os.path.isdir(canonical_output)
    if uses_canonical_layout or not state.get("preprocessor_folder_name", "").strip():
        state["preprocessor_folder_name"] = DEFAULT_PREPROCESSOR_FOLDER_NAME
    if uses_canonical_layout and state.get("crop_folder_name") == "01_raw_masked_for_jacop":
        state["crop_folder_name"] = DEFAULT_CROP_FOLDER_NAME
    if uses_canonical_layout or not state.get("final_output_folder_name", "").strip():
        state["final_output_folder_name"] = DEFAULT_FINAL_OUTPUT_FOLDER_NAME
    _, _, _, _, _, _, output_root = named_pipeline_paths(state)
    state["output_root"] = output_root
    return output_root


def redirect(location):
    data = b""
    return data


def safe_upload_relpath(filename):
    filename = filename.replace("\\", "/").lstrip("/")
    parts = []
    for part in filename.split("/"):
        if not part or part in (".", ".."):
            continue
        # Preserve file extensions. The older sanitizer used safe_name() for
        # single files, which stripped .xlsx and made valid uploads look invalid.
        parts.append("".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in part))
    return os.path.join(*parts) if parts else "uploaded_file"


def multipart_parts(body, content_type):
    match = re.search(r'boundary=(?:"([^"]+)"|([^;]+))', content_type or "", flags=re.IGNORECASE)
    if not match:
        raise ValueError("Upload did not include a multipart boundary.")
    boundary = (match.group(1) or match.group(2) or "").strip()
    delimiter = ("--" + boundary).encode("utf-8")
    for raw_part in body.split(delimiter):
        raw_part = raw_part.strip(b"\r\n")
        if not raw_part or raw_part == b"--":
            continue
        if raw_part.endswith(b"--"):
            raw_part = raw_part[:-2]
        header_blob, sep, content = raw_part.partition(b"\r\n\r\n")
        if not sep:
            continue
        headers = header_blob.decode("utf-8", errors="replace").split("\r\n")
        header_map = {}
        for line in headers:
            if ":" in line:
                key, value = line.split(":", 1)
                header_map[key.lower().strip()] = value.strip()
        yield header_map, content.rstrip(b"\r\n")


def parse_content_disposition(value):
    result = {}
    for piece in value.split(";"):
        piece = piece.strip()
        if "=" in piece:
            key, val = piece.split("=", 1)
            key = key.strip().lower()
            val = val.strip().strip('"')
            if key == "filename*":
                # Browser may send RFC 5987 filenames such as UTF-8''file.xlsx.
                if "''" in val:
                    val = val.split("''", 1)[1]
                val = unquote(val)
                result["filename"] = val
            else:
                result[key] = val
        else:
            result[piece.lower()] = True
    return result


def save_uploaded_raw_folder(headers, body):
    upload_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid4().hex[:8]
    upload_dir = os.path.join(UPLOAD_ROOT, upload_id)
    os.makedirs(upload_dir, exist_ok=True)
    file_count = 0
    first_parts = []

    for part_headers, content in multipart_parts(body, headers.get("Content-Type", "")):
        disposition = parse_content_disposition(part_headers.get("content-disposition", ""))
        filename = disposition.get("filename")
        if not filename:
            continue
        rel = safe_upload_relpath(filename)
        if os.path.basename(rel).startswith("._") or not rel.lower().endswith((".tif", ".tiff", ".png", ".jpg", ".jpeg", ".xlsx", ".csv", ".json")):
            # Keep this broad but avoid writing browser metadata files.
            continue
        destination = os.path.join(upload_dir, rel)
        os.makedirs(os.path.dirname(destination), exist_ok=True)
        with open(destination, "wb") as f:
            f.write(content)
        file_count += 1
        first_parts.append(rel.split(os.sep)[0])

    if file_count == 0:
        raise ValueError("No files were uploaded. Choose the raw image folder, not an empty folder.")

    return upload_dir, file_count








def choose_uploaded_raw_image_root(upload_dir):
    entries = [entry for entry in os.listdir(upload_dir) if not entry.startswith(".")]
    files = [entry for entry in entries if os.path.isfile(os.path.join(upload_dir, entry))]
    dirs = [entry for entry in entries if os.path.isdir(os.path.join(upload_dir, entry))]

    # Folder uploads often arrive as upload_id/RAW IMAGES/*.TIF. In that case,
    # use the inner folder as the experiment root so all derived output folders
    # are created beside the uploaded TIFFs.
    if len(dirs) == 1 and not files:
        return os.path.join(upload_dir, dirs[0])

    return upload_dir


def count_tiff_files(root):
    count = 0
    for dirpath, _, filenames in os.walk(root):
        for filename in filenames:
            path = os.path.join(dirpath, filename)
            if not filename.startswith("._") and filename.lower().endswith((".tif", ".tiff")) and is_real_tiff_file(path):
                count += 1
    return count


def normalize_selected_folder(folder):
    folder = str(folder or "").strip()
    if folder.endswith("/") and folder != "/":
        folder = folder[:-1]
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"Selected folder does not exist: {folder}")
    return folder


def choose_folder_with_osascript(prompt):
    script = (
        f'set chosenFolder to choose folder with prompt "{prompt}"\n'
        'POSIX path of chosenFolder'
    )
    result = subprocess.run(
        ["/usr/bin/osascript", "-e", script],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if result.returncode != 0:
        message = (result.stderr or result.stdout or "Folder selection cancelled.").strip()
        raise RuntimeError(message)
    return normalize_selected_folder(result.stdout)


def choose_folder_with_tk_subprocess(prompt):
    script = r'''
import os
import sys
import tkinter as tk
from tkinter import filedialog

prompt = sys.argv[1] if len(sys.argv) > 1 else "Choose folder"
root = tk.Tk()
root.withdraw()
root.update()
try:
    root.attributes("-topmost", True)
except Exception:
    pass
folder = filedialog.askdirectory(parent=root, title=prompt, mustexist=True)
root.destroy()
if not folder:
    sys.exit(2)
print(os.path.abspath(folder))
'''
    result = subprocess.run(
        [sys.executable, "-c", script, prompt],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=300,
    )
    if result.returncode != 0:
        message = (result.stderr or result.stdout or "Folder selection cancelled.").strip()
        raise RuntimeError(message)
    return normalize_selected_folder(result.stdout)


def choose_folder_with_appkit_subprocess(prompt):
    script = r'''
import os
import sys

from AppKit import NSApplication, NSOpenPanel
from AppKit import NSApplicationActivationPolicyRegular

prompt = sys.argv[1] if len(sys.argv) > 1 else "Choose folder"
app = NSApplication.sharedApplication()
app.setActivationPolicy_(NSApplicationActivationPolicyRegular)
app.activateIgnoringOtherApps_(True)

panel = NSOpenPanel.openPanel()
panel.setTitle_(prompt)
panel.setMessage_(prompt)
panel.setCanChooseDirectories_(True)
panel.setCanChooseFiles_(False)
panel.setAllowsMultipleSelection_(False)
panel.setCanCreateDirectories_(False)
panel.setResolvesAliases_(True)

response = panel.runModal()
if int(response) != 1:
    sys.exit(2)
urls = panel.URLs()
if not urls:
    sys.exit(2)
path = urls[0].path()
if not path:
    sys.exit(2)
print(os.path.abspath(str(path)))
'''
    result = subprocess.run(
        [sys.executable, "-c", script, prompt],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=300,
    )
    if result.returncode != 0:
        message = (result.stderr or result.stdout or "Folder selection cancelled.").strip()
        raise RuntimeError(message)
    return normalize_selected_folder(result.stdout)


def choose_local_folder_with_dialog(prompt):
    errors = []
    try:
        return choose_folder_with_osascript(prompt)
    except Exception as error:
        errors.append(f"AppleScript picker failed: {error}")
    try:
        return choose_folder_with_appkit_subprocess(prompt)
    except Exception as error:
        errors.append(f"AppKit picker failed: {error}")
    try:
        return choose_folder_with_tk_subprocess(prompt)
    except Exception as error:
        errors.append(f"Tk picker failed: {error}")
    raise RuntimeError(" | ".join(errors))


def choose_local_raw_folder_with_dialog():
    return choose_local_folder_with_dialog("Choose the raw image folder for this experiment")


def use_local_raw_image_folder(state, raw_image_root):
    raw_image_root = os.path.abspath(os.path.expanduser(str(raw_image_root or "").strip()))
    if not raw_image_root:
        raise ValueError("Paste the full path to the raw image folder.")
    if not os.path.isdir(raw_image_root):
        raise FileNotFoundError(f"Raw image folder does not exist: {raw_image_root}")
    tiff_count = count_tiff_files(raw_image_root)
    if tiff_count <= 0:
        raise ValueError(f"No TIFF images were found in this raw image folder: {raw_image_root}")
    # Never write a freshly cleared/default state into the selected
    # experiment before its portable state has been read.  Doing so destroys
    # the exact kept/rejected decisions and per-crop AutoClean/phenotype
    # settings that this reload operation is supposed to restore.
    state["_suspend_portable_job_state_save"] = True
    try:
        state["raw_image_folder"] = raw_image_root
        refresh_derived_paths(state)
        clear_cross_job_output_state(state)
        state["last_scan"] = ""
        state["last_stage_output"] = ""
        state["last_stage_returncode"] = None
        folder_state_source = restore_portable_job_globals(state)
        if os.path.isdir(state.get("crop_root", "")):
            crop_count = scan_crops(state)
            summary = reload_existing_results_for_current_folder(state)
            summary["folder_state_source"] = folder_state_source
            if summary.get("status") == "recovered":
                msg = f"Selected real raw image folder and found {tiff_count} TIFF images. Loaded {crop_count} existing crops and recovered previous work: kept {summary.get('kept')}, rejected {summary.get('rejected')}, manual recrops {summary.get('manual_recrops')}."
            elif crop_count > 0:
                msg = f"Selected real raw image folder and found {tiff_count} TIFF images. Loaded {crop_count} existing crops, but no previous decisions were found."
            else:
                msg = f"Selected real raw image folder and found {tiff_count} TIFF images. The expected crop folder exists but contains no TIFF crops; run preprocessing and crop selected ROIs first."
        else:
            msg = f"Selected real raw image folder and found {tiff_count} TIFF images. Results will be written beside this folder: {raw_image_root}"
    finally:
        state.pop("_suspend_portable_job_state_save", None)
    # Only the fully restored state is allowed to replace the portable file.
    save_state(state)
    return msg


def choose_local_previous_job_folder_with_dialog():
    return choose_local_folder_with_dialog("Choose any previous job folder: raw folder, results folder, raw_masked_analysis, crop folder, or parent folder")









def find_tiffs(root):
    files = []
    for dirpath, _, filenames in os.walk(root):
        if "rejected_crops" in dirpath or "approved_crops_for_both_analyses" in dirpath or "manual_corrected_crops" in dirpath:
            continue
        for filename in filenames:
            path = os.path.join(dirpath, filename)
            if not filename.startswith("._") and filename.lower().endswith((".tif", ".tiff")) and is_real_tiff_file(path):
                files.append(path)
    return sorted(files)


def resolve_appledouble_image_path(path):
    """Return the real data file when a saved path points at macOS `._` metadata."""
    path = str(path or "")
    basename = os.path.basename(path)
    if basename.startswith("._"):
        real_path = os.path.join(os.path.dirname(path), basename[2:])
        if os.path.isfile(real_path):
            return real_path
    return path


def scan_crops(state):
    crop_root = state["crop_root"]
    if not os.path.isdir(crop_root):
        raise FileNotFoundError(f"Crop folder does not exist: {crop_root}")

    existing = {
        rel: item
        for rel, item in state.get("items", {}).items()
        if not any(part.startswith("._") for part in rel.split(os.sep))
    }
    new_items = {}
    for idx, path in enumerate(find_tiffs(crop_root), start=1):
        rel = os.path.relpath(path, crop_root)
        item = existing.get(rel, {})
        for saved_path_key in ("path", "manual_crop_path", "manual_original_path", "auto_cleaned_source_path"):
            if saved_path_key in item:
                item[saved_path_key] = resolve_appledouble_image_path(item[saved_path_key])
        item.update({
            "rel_path": rel,
            "path": path,
            "crop_number": idx,
            "status": item.get("status", "keep"),
            "crop_choice": item.get("crop_choice", "raw"),
            "note": item.get("note", ""),
        })
        new_items[rel] = item
    state["items"] = new_items
    state["last_scan"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Scanning discovers files before recovery runs. Save the app-local state,
    # but do not overwrite the experiment's portable snapshot with default
    # all-KEEP decisions before those saved decisions can be reloaded.
    was_suspended = bool(state.get("_suspend_portable_job_state_save"))
    state["_suspend_portable_job_state_save"] = True
    try:
        save_state(state)
    finally:
        if not was_suspended:
            state.pop("_suspend_portable_job_state_save", None)
    return len(new_items)


def parse_crop_number_selection(selection, max_number):
    selected = set()
    tokens = re.split(r"[\s,]+", str(selection or "").strip())
    for token in tokens:
        if not token:
            continue
        if "-" in token:
            start_text, end_text = token.split("-", 1)
            if not start_text.strip().isdigit() or not end_text.strip().isdigit():
                raise ValueError(f"Invalid crop number range: {token}")
            start = int(start_text)
            end = int(end_text)
            if start > end:
                start, end = end, start
            selected.update(range(start, end + 1))
        else:
            if not token.isdigit():
                raise ValueError(f"Invalid crop number: {token}")
            selected.add(int(token))
    return {number for number in selected if 1 <= number <= max_number}


def ensure_crop_numbers(state):
    items = state.get("items", {})
    changed = False
    for idx, rel in enumerate(sorted(items), start=1):
        item = items[rel]
        if int(item.get("crop_number", 0) or 0) <= 0:
            item["crop_number"] = idx
            changed = True
    if changed:
        save_state(state)
    return sorted(items.values(), key=lambda item: (int(item.get("crop_number", 0) or 0), item.get("rel_path", "")))


def apply_crop_number_selection(state, selection):
    items = ensure_crop_numbers(state)
    max_number = max((int(item.get("crop_number", 0) or 0) for item in items), default=0)
    selected_numbers = parse_crop_number_selection(selection, max_number)
    if not selected_numbers:
        raise ValueError("Enter at least one valid crop number, such as 1-20 or 1,3,5.")
    for item in items:
        number = int(item.get("crop_number", 0) or 0)
        if number not in selected_numbers:
            item["status"] = "reject"
        else:
            item["status"] = item.get("status", "keep")
    save_state(state)
    kept_now = sum(1 for item in items if item.get("status", "keep") == "keep")
    rejected_now = sum(1 for item in items if item.get("status", "keep") == "reject")
    return kept_now, rejected_now, sorted(selected_numbers)


def recovery_crop_token(value):
    stem = os.path.splitext(os.path.basename(str(value or "")))[0]
    for suffix in (
        "_raw_masked_crop_for_JACoP",
        "_enhanced_crop_optional_for_JACoP",
        "_rgb_crop_sent_to_fiji",
        "_sent_to_fiji_preview",
    ):
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)]
    return stem


def recovery_status_from_text(value):
    text = str(value or "").strip().lower()
    if text in ("keep", "kept", "included", "approved"):
        return "keep"
    if text in ("reject", "rejected", "excluded"):
        return "reject"
    return ""


def recovery_int_value(value, fallback, minimum=None, maximum=None):
    try:
        if pd.isna(value):
            raise ValueError
        result = int(round(float(value)))
    except (TypeError, ValueError, OverflowError):
        result = int(fallback)
    if minimum is not None:
        result = max(int(minimum), result)
    if maximum is not None:
        result = min(int(maximum), result)
    return result


def recovery_find_tiffs(root):
    files = []
    if not os.path.isdir(root):
        return files
    for dirpath, _, filenames in os.walk(root):
        for filename in filenames:
            path = os.path.join(dirpath, filename)
            if not filename.startswith("._") and filename.lower().endswith((".tif", ".tiff")) and is_real_tiff_file(path):
                files.append(path)
    return sorted(files)


def recovery_find_result_files(root):
    files = []
    if not os.path.isdir(root):
        return files
    allowed_ext = (".tif", ".tiff", ".png", ".csv", ".txt", ".xlsx", ".xls")
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".") and d not in {"__pycache__", "state_backups"}]
        for filename in filenames:
            if not filename.startswith("._") and filename.lower().endswith(allowed_ext):
                files.append(os.path.join(dirpath, filename))
    return sorted(files)


def recovery_is_keep_evidence_file(path):
    lower_path = path.lower()
    lower_name = os.path.basename(path).lower()
    evidence_markers = (
        "approved_crops_for_both_analyses",
        "approved_crops_for_real_fiji_pcc_manders",
        "approved_crops_for_condensate_count_size",
        "rgb_crops_sent_to_fiji",
        "channel_inputs_sent_to_fiji",
        "preview_images_sent_to_fiji",
        "real_fiji_binary_watershed_masks",
        "real_fiji_tables",
        "fiji_jacop_logs",
    )
    if any(marker in lower_path for marker in evidence_markers):
        return True
    return any(
        marker in lower_name
        for marker in (
            "_rgb_crop_sent_to_fiji",
            "_sent_to_fiji_preview",
            "_fiji_log",
            "_mask",
            "particles.csv",
            "summary.csv",
        )
    )


def recovery_is_approved_crop_file(path):
    lower_path = path.lower()
    return lower_path.endswith((".tif", ".tiff")) and any(
        marker in lower_path
        for marker in (
            "approved_crops_for_both_analyses",
            "approved_crops_for_real_fiji_pcc_manders",
            "approved_crops_for_condensate_count_size",
        )
    )


def is_manual_recrop_file(path):
    name = os.path.basename(str(path or "")).lower()
    return name.endswith((".tif", ".tiff")) and (
        "manual_recrop" in name or name.startswith("manual_recrop__")
    )


def find_latest_matching_file(root, predicates):
    matches = []
    if not root or not os.path.isdir(root):
        return ""
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".") and d not in {"__pycache__", "state_backups"}]
        for filename in filenames:
            if filename.startswith("._"):
                continue
            lower = filename.lower()
            if any(predicate(filename, lower, dirpath) for predicate in predicates):
                path = os.path.join(dirpath, filename)
                if os.path.isfile(path):
                    matches.append(path)
    return max(matches, key=lambda path: os.path.getmtime(path)) if matches else ""


def find_latest_result_excel(root, result_kind):
    if result_kind == "pcc_manders":
        return find_latest_matching_file(root, [
            lambda name, lower, _dir: lower.endswith((".xlsx", ".xls")) and "pcc" in lower and "manders" in lower and "reviewed_rois" in lower,
            lambda name, lower, _dir: lower.endswith((".xlsx", ".xls")) and lower.startswith("real_fiji_pcc_manders_reviewed_rois"),
        ])
    if result_kind == "count_size":
        return find_latest_matching_file(root, [
            lambda name, lower, _dir: lower.endswith((".xlsx", ".xls")) and "count" in lower and "size" in lower and ("reviewed_rois" in lower or "condensate" in lower),
            lambda name, lower, _dir: lower.endswith((".xlsx", ".xls")) and lower.startswith("real_fiji_condensate_count_size_reviewed_rois"),
        ])
    return ""


def set_existing_path_or_blank(state, key, path):
    state[key] = path if path and os.path.exists(path) else ""


def clear_cross_job_output_state(state):
    for key in (
        "last_real_fiji_pcc_manders_excel",
        "last_pcc_manders_excel",
        "last_pcc_manders_preview_root",
        "last_pcc_manders_approved_root",
        "last_count_size_excel",
        "last_count_size_output_root",
        "last_count_size_approved_root",
        "last_fixed_cell_ms2_excel",
        "last_fixed_cell_ms2_output_root",
        "last_fixed_cell_ms2_preview_root",
        "last_partition_excel",
        "last_partition_output_root",
        "last_partition_preview_root",
        "final_preview_root",
        "last_final_excel",
    ):
        state[key] = ""
    # These values are keyed by crop-relative paths and must never leak from
    # the previously opened experiment into the next one.  The selected job's
    # portable state/backups will restore only its own entries below.
    state["items"] = {}
    state["per_crop_cleaning_params"] = {}
    state["count_size_thresholds"] = {}
    state.pop("last_auto_cleaned_reload_scan_signature", None)


def find_named_dir(root, target_names, max_depth=6):
    if not root or not os.path.isdir(root):
        return ""
    targets = {str(name).lower() for name in target_names if str(name or "").strip()}
    root = os.path.abspath(root)
    for dirpath, dirnames, _ in os.walk(root):
        rel = os.path.relpath(dirpath, root)
        depth = 0 if rel == "." else rel.count(os.sep) + 1
        if depth > max_depth:
            dirnames[:] = []
            continue
        dirnames[:] = [d for d in dirnames if not d.startswith(".") and d not in {"__pycache__", "state_backups"}]
        if os.path.basename(dirpath).lower() in targets:
            return dirpath
    return ""


def infer_previous_job_paths(folder, state):
    folder = os.path.abspath(str(folder or "").strip())
    if not folder or not os.path.isdir(folder):
        raise FileNotFoundError(f"Previous job/results folder does not exist: {folder}")

    default_preprocessor_name = "ENHANCED_JACOP_PREPROCESSOR_RESULTS"
    default_output_name = "REAL FIJI PCC MANDERS RESULTS"
    default_crop_name = "03_cropped_rois_for_final_choice"
    preprocessor_name = state.get("preprocessor_folder_name", default_preprocessor_name).strip() or default_preprocessor_name
    crop_name = state.get("crop_folder_name", default_crop_name).strip() or default_crop_name
    output_name = state.get("final_output_folder_name", default_output_name).strip() or default_output_name
    preprocessor_names = [name for name in (preprocessor_name, default_preprocessor_name) if name]
    if crop_name == "01_raw_masked_for_jacop":
        crop_names = [default_crop_name, crop_name]
    else:
        crop_names = [crop_name, default_crop_name]
    output_names = [name for name in (output_name, default_output_name) if name]
    crop_leaf = "01_raw_masked_roi_crops_for_jacop"
    preprocessor_crop_leaf = "01_raw_masked_for_jacop"

    def nested_preprocessor_crop_root(path):
        nested = os.path.join(path, crop_leaf)
        return nested if os.path.isdir(nested) else ""

    def first_existing_preprocessor_root(raw_path):
        if os.path.isdir(os.path.join(raw_path, default_output_name)):
            return os.path.join(raw_path, default_preprocessor_name), default_preprocessor_name
        for name in preprocessor_names:
            path = os.path.join(raw_path, name)
            if os.path.isdir(path):
                return path, name
        return os.path.join(raw_path, preprocessor_name), preprocessor_name

    def first_existing_output_root(raw_path):
        if os.path.isdir(os.path.join(raw_path, default_preprocessor_name)):
            return os.path.join(raw_path, default_output_name, "raw_masked_analysis"), default_output_name
        for name in output_names:
            path = os.path.join(raw_path, name, "raw_masked_analysis")
            if os.path.isdir(path):
                return path, name
        return os.path.join(raw_path, output_name, "raw_masked_analysis"), output_name

    def crop_root_from_preprocessor(preprocessor_root):
        candidate_crop = ""
        for name in crop_names:
            if name == preprocessor_crop_leaf:
                continue
            candidate_crop = os.path.join(preprocessor_root, name, crop_leaf)
            if os.path.isdir(candidate_crop):
                return candidate_crop
        preprocessor_crop = os.path.join(preprocessor_root, preprocessor_crop_leaf)
        if os.path.isdir(preprocessor_crop):
            nested_crop = nested_preprocessor_crop_root(preprocessor_crop)
            if nested_crop:
                return nested_crop
        return candidate_crop

    raw_root = ""
    crop_root = ""
    output_root = ""
    base = os.path.basename(folder)

    if base == crop_leaf:
        crop_root = folder
        crop_base = os.path.dirname(folder)
        preprocessor_root = os.path.dirname(crop_base)
        raw_root = os.path.dirname(preprocessor_root)
    elif base == preprocessor_crop_leaf:
        preprocessor_root = os.path.dirname(folder)
        raw_root = os.path.dirname(preprocessor_root)
        crop_root = nested_preprocessor_crop_root(folder) or os.path.join(preprocessor_root, default_crop_name, crop_leaf)
    elif base == crop_name:
        crop_root = os.path.join(folder, crop_leaf)
        preprocessor_root = os.path.dirname(folder)
        raw_root = os.path.dirname(preprocessor_root)
    elif base in preprocessor_names:
        raw_root = os.path.dirname(folder)
        crop_root = crop_root_from_preprocessor(folder)
    elif base in ("raw_masked_analysis", "enhanced_masked_analysis"):
        output_root = folder
        raw_root = os.path.dirname(os.path.dirname(folder))
    elif base in (
        "real_fiji_pcc_manders_analysis",
        "real_fiji_condensate_count_size_analysis",
        "approved_crops_for_real_fiji_pcc_manders",
        "approved_crops_for_condensate_count_size",
        "channel_cleaned_analysis_inputs",
    ):
        parent = os.path.dirname(folder)
        if os.path.basename(parent) in ("raw_masked_analysis", "enhanced_masked_analysis"):
            output_root = parent
            raw_root = os.path.dirname(os.path.dirname(parent))
    elif base in output_names:
        raw_root = os.path.dirname(folder)
        output_root = os.path.join(folder, "raw_masked_analysis")
    else:
        preprocessor_root, _ = first_existing_preprocessor_root(folder)
        candidate_crop = crop_root_from_preprocessor(preprocessor_root)
        candidate_output, _ = first_existing_output_root(folder)
        if os.path.isdir(candidate_crop) or os.path.isdir(candidate_output):
            raw_root = folder
            crop_root = candidate_crop
            output_root = candidate_output
        else:
            found_crop = find_named_dir(folder, [crop_leaf, preprocessor_crop_leaf], max_depth=7)
            found_output = find_named_dir(folder, ["raw_masked_analysis"], max_depth=7)
            if found_crop:
                crop_root = found_crop
                if os.path.basename(found_crop) == preprocessor_crop_leaf:
                    preprocessor_root = os.path.dirname(found_crop)
                    crop_root = nested_preprocessor_crop_root(found_crop) or os.path.join(preprocessor_root, default_crop_name, crop_leaf)
                else:
                    crop_base = os.path.dirname(found_crop)
                    preprocessor_root = os.path.dirname(crop_base)
                raw_root = os.path.dirname(preprocessor_root)
            if found_output:
                output_root = found_output
                possible_raw = os.path.dirname(os.path.dirname(found_output))
                if os.path.isdir(possible_raw):
                    raw_root = possible_raw

    if raw_root and not crop_root:
        preprocessor_root, _ = first_existing_preprocessor_root(raw_root)
        crop_root = crop_root_from_preprocessor(preprocessor_root)
    if raw_root and not output_root:
        output_root, _ = first_existing_output_root(raw_root)
    preprocessor_ready = bool(raw_root and os.path.isdir(os.path.join(raw_root, default_preprocessor_name, preprocessor_crop_leaf)))
    if (not crop_root or not os.path.isdir(crop_root)) and (not output_root or not os.path.isdir(output_root)) and not preprocessor_ready:
        raise FileNotFoundError(
            "Could not find previous crops or Fiji result outputs for this previous job. Expected a folder like "
            f"{preprocessor_name}/{crop_name}/{crop_leaf} or {output_name}/raw_masked_analysis under: {folder}"
        )
    if not output_root:
        output_root = os.path.join(raw_root or os.path.dirname(folder), output_name, "raw_masked_analysis")
    return raw_root, crop_root, output_root


def previous_job_stage_summary(state):
    raw_root = state.get("raw_image_folder", "")
    crop_root = state.get("crop_root", "")
    output_root = state.get("output_root", "")
    pcc_root = os.path.join(output_root, "real_fiji_pcc_manders_analysis")
    count_size_root = os.path.join(output_root, "real_fiji_condensate_count_size_analysis")
    fixed_cell_root = os.path.join(output_root, "fixed_cell_ms2_condensate_analysis")
    partition_root = os.path.join(output_root, "manual_nuclear_partition_coefficient_analysis")
    pcc_excel = find_latest_result_excel(pcc_root if os.path.isdir(pcc_root) else output_root, "pcc_manders")
    count_size_excel = find_latest_result_excel(count_size_root if os.path.isdir(count_size_root) else output_root, "count_size")
    summary = {
        "raw_folder": raw_root,
        "raw_tiff_count": count_tiff_files(raw_root) if os.path.isdir(raw_root) else 0,
        "preprocessor_root": os.path.join(raw_root, state.get("preprocessor_folder_name", "ENHANCED_JACOP_PREPROCESSOR_RESULTS")),
        "preprocessor_exists": os.path.isdir(os.path.join(raw_root, state.get("preprocessor_folder_name", "ENHANCED_JACOP_PREPROCESSOR_RESULTS"))),
        "crop_root": crop_root,
        "crop_count": len(recovery_find_tiffs(crop_root)),
        "output_root": output_root,
        "pcc_manders_excel": pcc_excel or os.path.join(pcc_root, "Real_Fiji_PCC_Manders_Reviewed_ROIs.xlsx"),
        "pcc_manders_preview_root": os.path.join(pcc_root, "preview_images_sent_to_fiji"),
        "count_size_excel": count_size_excel or os.path.join(count_size_root, "Real_Fiji_Condensate_Count_Size_Reviewed_ROIs.xlsx"),
        "count_size_output_root": count_size_root,
        "fixed_cell_ms2_excel": os.path.join(fixed_cell_root, "Fixed_Cell_MS2_Condensate_Per_Cell_Measurements.xlsx"),
        "fixed_cell_ms2_output_root": fixed_cell_root,
        "partition_excel": os.path.join(partition_root, "Manual_Nuclear_Partition_Coefficient_Measurements.xlsx"),
        "partition_output_root": partition_root,
        "saved_count_size_thresholds": len(count_size_thresholds_from_state(state)),
        "per_crop_cleaning_settings": len(state.get("per_crop_cleaning_params", {}) if isinstance(state.get("per_crop_cleaning_params", {}), dict) else {}),
        "result_evidence_files": len([path for path in recovery_find_result_files(output_root) if recovery_is_keep_evidence_file(path)]),
        "approved_crop_files": len([
            path for path in recovery_find_result_files(output_root)
            if recovery_is_approved_crop_file(path)
        ]),
    }
    for key in (
        "pcc_manders_excel", "pcc_manders_preview_root",
        "count_size_excel", "count_size_output_root",
        "fixed_cell_ms2_excel", "fixed_cell_ms2_output_root",
        "partition_excel", "partition_output_root",
    ):
        path = summary[key]
        summary[key + "_exists"] = os.path.exists(path)
    return summary


def _load_previous_job_folder_unlocked(state, folder):
    raw_root, crop_root, output_root = infer_previous_job_paths(folder, state)
    state["_suspend_portable_job_state_save"] = True
    try:
        clear_cross_job_output_state(state)
        if crop_root and os.path.basename(crop_root) == "01_raw_masked_roi_crops_for_jacop":
            crop_base = os.path.dirname(crop_root)
            preprocessor_root = os.path.dirname(crop_base)
            if os.path.basename(crop_base):
                state["crop_folder_name"] = os.path.basename(crop_base)
            if os.path.basename(preprocessor_root):
                state["preprocessor_folder_name"] = os.path.basename(preprocessor_root)
        if output_root and os.path.basename(output_root) in ("raw_masked_analysis", "enhanced_masked_analysis"):
            output_parent = os.path.dirname(output_root)
            if os.path.basename(output_parent):
                state["final_output_folder_name"] = os.path.basename(output_parent)
        state["raw_image_folder"] = raw_root
        state["crop_root"] = crop_root
        state["output_root"] = output_root
        state["final_input_choice"] = "raw_masked"
        folder_state_source = restore_portable_job_globals(state)
        if crop_root and os.path.isdir(crop_root) and os.path.basename(crop_root) == "01_raw_masked_roi_crops_for_jacop":
            count = scan_crops(state)
        else:
            state["items"] = {}
            state["last_scan"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            count = 0
            save_state(state)
        summary = reload_existing_results_for_current_folder(state)
        preprocessor_ready_root = os.path.join(raw_root, state.get("preprocessor_folder_name", DEFAULT_PREPROCESSOR_FOLDER_NAME), "01_raw_masked_for_jacop")
        if count == 0 and os.path.isdir(preprocessor_ready_root) and (not crop_root or not os.path.isdir(crop_root)):
            summary["status"] = "preprocessor_ready"
            summary["error"] = ""
            summary["preprocessor_ready_root"] = preprocessor_ready_root
            summary["preprocessor_ready_tiffs"] = len(recovery_find_tiffs(preprocessor_ready_root))
        summary["loaded_crops"] = count
        summary["raw_image_folder"] = raw_root
        summary["crop_root"] = crop_root
        summary["output_root"] = output_root
        summary["workflow_steps_run"] = []
        summary["waiting_for_next_button"] = True
        summary["folder_state_source"] = folder_state_source
        summary["stage_summary"] = previous_job_stage_summary(state)
        state["last_workflow_action"] = "reload_detect_only"
        state["last_workflow_note"] = "Reload detected existing files only. No crop, cleaning, Fiji, PCC/Manders, or count/size step was run."
        state["last_previous_job_folder"] = os.path.abspath(folder)
        state["last_previous_job_load_summary"] = summary
    finally:
        state.pop("_suspend_portable_job_state_save", None)
    save_state(state)
    return summary


def load_previous_job_folder(state, folder):
    # ThreadingHTTPServer may receive duplicate form submissions at the same
    # time.  Serialize the complete job switch/recovery transaction so one
    # request cannot save the other request's half-scanned default state.
    with JOB_STATE_LOCK:
        return _load_previous_job_folder_unlocked(state, folder)


def recovery_output_candidates(state):
    output_root = state.get("output_root", "")
    raw_root = state.get("raw_image_folder", "")
    paths = []
    seen = set()

    def add_path(path):
        key = os.path.abspath(path)
        if key in seen:
            return
        seen.add(key)
        paths.append(path)

    def scan_result_root(root):
        if not root or not os.path.isdir(root):
            return
        for path in recovery_find_result_files(root):
            if recovery_is_keep_evidence_file(path):
                add_path(path)
        for folder_name in ("approved_crops_for_real_fiji_pcc_manders", "approved_crops_for_condensate_count_size"):
            folder = os.path.join(root, folder_name)
            if os.path.isdir(folder):
                for path in recovery_find_tiffs(folder):
                    add_path(path)
        for excel_path in (
            os.path.join(root, "real_fiji_pcc_manders_analysis", "Real_Fiji_PCC_Manders_Reviewed_ROIs.xlsx"),
            os.path.join(root, "real_fiji_condensate_count_size_analysis", "Real_Fiji_Condensate_Count_Size_Reviewed_ROIs.xlsx"),
        ):
            if os.path.isfile(excel_path):
                add_path(excel_path)
        for result_kind, subfolder in (
            ("pcc_manders", "real_fiji_pcc_manders_analysis"),
            ("count_size", "real_fiji_condensate_count_size_analysis"),
        ):
            excel_path = find_latest_result_excel(os.path.join(root, subfolder), result_kind)
            if excel_path:
                add_path(excel_path)

    scan_result_root(output_root)

    if raw_root and os.path.isdir(raw_root):
        raw_abs = os.path.abspath(raw_root)
        skipped = {
            "ENHANCED_JACOP_PREPROCESSOR_RESULTS",
            "03_cropped_rois_for_final_choice",
            "01_raw_masked_roi_crops_for_jacop",
            "channel_cleaned_analysis_inputs",
            "__pycache__",
        }
        for dirpath, dirnames, filenames in os.walk(raw_root):
            dirnames[:] = [d for d in dirnames if d not in skipped and not d.startswith(".")]
            dir_abs = os.path.abspath(dirpath)
            if not dir_abs.startswith(raw_abs):
                continue
            base = os.path.basename(dirpath)
            if base in ("raw_masked_analysis", "enhanced_masked_analysis"):
                scan_result_root(dirpath)
            if base in ("approved_crops_for_real_fiji_pcc_manders", "approved_crops_for_condensate_count_size"):
                for path in recovery_find_tiffs(dirpath):
                    add_path(path)
            for filename in filenames:
                path = os.path.join(dirpath, filename)
                if recovery_is_keep_evidence_file(path):
                    add_path(path)
                    continue
                lower = filename.lower()
                if (
                    filename in ("Real_Fiji_PCC_Manders_Reviewed_ROIs.xlsx", "Real_Fiji_Condensate_Count_Size_Reviewed_ROIs.xlsx")
                    or (lower.endswith((".xlsx", ".xls")) and "reviewed_rois" in lower and ("pcc" in lower or "manders" in lower or ("count" in lower and "size" in lower)))
                ):
                    add_path(path)

    return paths


def recovery_manual_corrected_matches(state, items):
    crop_root = state.get("crop_root", "")
    matches = {}
    if not os.path.isdir(crop_root):
        return matches
    # A later crop run archives the previous manual-corrected directory under
    # _previous_crop_runs/<timestamp>. Search those archives as well as the
    # active directory so reload does not silently fall back to the raw ROI.
    manual_paths = []
    search_roots = [crop_root]
    crop_parent = os.path.dirname(crop_root)
    archive_root = os.path.join(crop_parent, "_previous_crop_runs")
    if os.path.isdir(archive_root):
        search_roots.append(archive_root)
    for search_root in search_roots:
        for path in recovery_find_tiffs(search_root):
            lower = path.lower()
            if "manual_corrected_crops" in lower or "manual_recrop" in os.path.basename(path).lower():
                manual_paths.append(path)
    for rel in items:
        token = recovery_crop_token(rel)
        candidates = [path for path in manual_paths if token and token in os.path.basename(path)]
        if candidates:
            matches[rel] = max(candidates, key=lambda path: os.path.getmtime(path))
    return matches


def restore_archived_manual_recrops(state):
    """Restore manual-recrops moved into a previous crop-run archive."""
    items = state.get("items", {})
    if not isinstance(items, dict) or not items:
        return 0
    matches = recovery_manual_corrected_matches(state, items)
    restored = 0
    crop_root = state.get("crop_root", "")
    for rel, manual_path in matches.items():
        item = items.get(rel)
        if not isinstance(item, dict) or not os.path.isfile(manual_path):
            continue
        current_manual = item.get("manual_crop_path", "")
        if current_manual and os.path.isfile(current_manual):
            continue
        item["path"] = manual_path
        item["manual_crop_path"] = manual_path
        item["manual_original_path"] = os.path.join(crop_root, rel)
        item["crop_choice"] = "manual_recrop"
        item.setdefault("status", "keep")
        item["note"] = item.get("note") or "Restored from a previous crop-run manual recrop archive."
        restored += 1
    return restored


def recovery_state_backup_paths():
    backup_dirs = [os.path.join(APP_ROOT, "state_backups")]
    paths = []
    for backup_dir in backup_dirs:
        if not os.path.isdir(backup_dir):
            continue
        paths.extend(
            os.path.join(backup_dir, filename)
            for filename in os.listdir(backup_dir)
            if filename.lower().endswith(".json")
        )
    return sorted(paths, key=lambda path: os.path.getmtime(path), reverse=True)


def recover_from_state_backups(state, items):
    current_crop_root = os.path.abspath(state.get("crop_root", ""))
    current_raw_root = os.path.abspath(state.get("raw_image_folder", ""))
    candidates = []
    for path in recovery_state_backup_paths():
        try:
            with open(path, "r", encoding="utf-8") as f:
                backup = json.load(f)
        except Exception:
            continue
        backup_crop_root = os.path.abspath(backup.get("crop_root", ""))
        backup_raw_root = os.path.abspath(backup.get("raw_image_folder", ""))
        if current_crop_root and backup_crop_root and backup_crop_root != current_crop_root:
            continue
        if current_raw_root and backup_raw_root and backup_raw_root != current_raw_root:
            continue
        backup_items = backup.get("items", {})
        if not isinstance(backup_items, dict):
            continue
        recovered_status = {}
        approved_matches = {}
        recovered_item_settings = {}
        matched_this_file = False
        for rel, item in items.items():
            backup_item = backup_items.get(rel)
            if not isinstance(backup_item, dict):
                continue
            status = recovery_status_from_text(backup_item.get("status", ""))
            if status:
                recovered_status[rel] = status
                matched_this_file = True
            manual_path = str(backup_item.get("manual_crop_path", "") or "").strip()
            selected_path = str(backup_item.get("path", "") or "").strip()
            for candidate in (manual_path, selected_path):
                if candidate and os.path.isfile(candidate) and is_manual_recrop_file(candidate):
                    approved_matches[rel] = candidate
                    matched_this_file = True
                    break
            item_settings = {}
            backup_roles = backup_item.get("channel_roles")
            if isinstance(backup_roles, dict):
                roles = {}
                for channel_name in CHANNELS:
                    roles[channel_name] = valid_channel_role(
                        backup_roles.get(channel_name),
                        channel_roles_from_state(backup).get(channel_name, "ignore"),
                    )
                item_settings["channel_roles"] = roles
                matched_this_file = True
            for key in ("explicit_keep",):
                if key in backup_item:
                    item_settings[key] = bool(backup_item[key])
            if item_settings:
                recovered_item_settings[rel] = item_settings
        if matched_this_file:
            rejected = sum(1 for status in recovered_status.values() if status == "reject")
            kept = sum(1 for status in recovered_status.values() if status == "keep")
            manual = len(approved_matches)
            matched = len(recovered_status)
            explicit = sum(
                1 for rel in items
                if isinstance(backup_items.get(rel), dict) and "explicit_keep" in backup_items.get(rel, {})
            )
            # A fresh crop scan assigns "keep" to every crop before the user
            # reviews anything.  Such a backup is not decision evidence and
            # must not override the Review Decisions/Approved Crops outputs.
            # This was the reason a completed 40-kept/120-rejected job could
            # reopen as 160 kept.
            if rejected == 0 and explicit == 0:
                continue
            # Prefer backups that contain real user review actions. Later preview
            # or scan-only states can accidentally look valid but have no explicit
            # keep/reject markers, so they should not replace reviewed decisions.
            score = (1 if explicit else 0, explicit, 1 if rejected else 0, os.path.getmtime(path), manual, matched)
            candidates.append((score, recovered_status, approved_matches, recovered_item_settings, path, kept, rejected, manual, backup))
    if not candidates:
        return {}, {}, {}, []
    candidates.sort(key=lambda row: row[0], reverse=True)
    _, recovered_status, approved_matches, recovered_item_settings, path, _, _, _, backup = candidates[0]
    for key in (
        "count_size_thresholds",
        "per_crop_cleaning_params",
        "pcc_red_label",
        "pcc_green_label",
        "pcc_blue_label",
        "analysis_input_mode",
        "pcc_analysis_input_mode",
        "count_size_analysis_input_mode",
        "red_channel_role",
        "green_channel_role",
        "blue_channel_role",
        "pcc_pair_mode",
        "count_size_red",
        "count_size_green",
        "count_size_blue",
        "pixel_um",
    ):
        if key in backup:
            state[key] = backup[key]
    return recovered_status, approved_matches, recovered_item_settings, [path]


def recover_from_portable_job_state(state, items):
    """Recover the complete workflow state from the selected experiment folder."""
    candidates = []
    direct = portable_job_state_path(state)
    if direct and os.path.isfile(direct):
        candidates.append(direct)
    raw_root = state.get("raw_image_folder", "")
    if raw_root and os.path.isdir(raw_root):
        for dirpath, dirnames, filenames in os.walk(raw_root):
            rel = os.path.relpath(dirpath, raw_root)
            depth = 0 if rel == "." else rel.count(os.sep) + 1
            if depth > 6:
                dirnames[:] = []
                continue
            dirnames[:] = [name for name in dirnames if not name.startswith(".") and name not in {"__pycache__", "state_backups"}]
            if JOB_STATE_FILENAME in filenames:
                candidate = os.path.join(dirpath, JOB_STATE_FILENAME)
                if candidate not in candidates:
                    candidates.append(candidate)
    # The state at the inferred output location is authoritative. Recursive
    # candidates are only a compatibility fallback for older folder layouts.
    candidates.sort(key=lambda candidate: (candidate != direct, -os.path.getmtime(candidate)))
    for path in candidates:
        try:
            with open(path, "r", encoding="utf-8") as f:
                snapshot = json.load(f)
        except Exception:
            continue
        # The direct file is accepted even after the whole experiment folder
        # was moved. A recursively discovered compatibility candidate still
        # has to identify the current job.
        mismatched_job = False
        if path != direct:
            for key in ("raw_image_folder", "crop_root", "output_root"):
                current_path = str(state.get(key, "") or "").strip()
                saved_path = str(snapshot.get(key, "") or "").strip()
                if current_path and saved_path and os.path.abspath(current_path) != os.path.abspath(saved_path):
                    mismatched_job = True
                    break
        if mismatched_job:
            continue
        saved_items = snapshot.get("items", {})
        if not isinstance(saved_items, dict):
            continue
        statuses = {}
        item_settings = {}
        for rel in items:
            saved = saved_items.get(rel)
            if not isinstance(saved, dict):
                continue
            status = recovery_status_from_text(saved.get("status", ""))
            if status:
                statuses[rel] = status
            settings = {
                key: copy.deepcopy(saved[key])
                for key in PORTABLE_JOB_ITEM_KEYS
                if key in saved
            }
            for key in (
                "manual_crop_path", "manual_original_path",
                "auto_cleaned_crop_path", "auto_cleaned_source_path",
                "fixed_cell_manual_annotation_preview_path",
                "partition_annotation_preview_path",
            ):
                if settings.get(key):
                    settings[key] = rebase_portable_job_path(settings[key], snapshot, state)
            for key in ("fixed_cell_manual_annotation", "partition_manual_annotation"):
                if key in settings:
                    settings[key] = rebase_portable_annotation(settings[key], snapshot, state)
            if settings:
                item_settings[rel] = settings
        rejected_count = sum(1 for status in statuses.values() if status == "reject")
        explicit_decision_count = sum(
            1
            for rel in items
            if isinstance(saved_items.get(rel), dict)
            and "explicit_keep" in saved_items.get(rel, {})
        )
        has_decision_evidence = rejected_count > 0 or explicit_decision_count > 0
        # Restore only this job's settings. Per-crop mappings are filtered to
        # the crops that were freshly scanned in the selected experiment.
        for key in PORTABLE_JOB_GLOBAL_KEYS:
            if key == "per_crop_cleaning_params":
                saved_map = snapshot.get(key, {})
                state[key] = {
                    rel: saved_map[rel]
                    for rel in items
                    if isinstance(saved_map, dict) and rel in saved_map
                }
            elif key == "count_size_thresholds":
                saved_map = snapshot.get(key, {})
                state[key] = {
                    threshold_key: row
                    for threshold_key, row in (saved_map.items() if isinstance(saved_map, dict) else [])
                    if (
                        str(row.get("rel", "") if isinstance(row, dict) else "") in items
                        or str(threshold_key).rsplit("|||", 1)[0] in items
                    )
                }
            elif key in snapshot:
                state[key] = copy.deepcopy(snapshot[key])
        for key in PORTABLE_JOB_OUTPUT_KEYS:
            if key in snapshot:
                state[key] = rebase_portable_job_path(snapshot[key], snapshot, state)
        if not has_decision_evidence:
            # A scan initializes every crop as "keep" before the user reviews
            # anything.  That default-only snapshot is useful for metadata,
            # but it is not decision evidence.  Let the reviewed workbook or
            # approved-crop outputs restore keep/reject instead of reopening
            # every crop as kept.
            return {}, item_settings, [path]
        if not statuses:
            continue
        return statuses, item_settings, [path]
    return {}, {}, []


def recover_review_state_from_existing_outputs(state):
    refresh_derived_paths(state)
    count = scan_crops(state)
    items = state.get("items", {})
    if not items:
        raise ValueError("No crop images are loaded. Check the raw folder and crop folder names first.")

    recovery_paths = recovery_output_candidates(state)
    approved_matches = {}
    recovered_status, backup_item_settings, portable_sources = recover_from_portable_job_state(state, items)
    backup_sources = list(portable_sources)
    # A reviewed workbook is authoritative over loose approved-crop files.
    # Output directories can retain stale TIFFs from an older analysis run.
    if not recovered_status:
        for path in recovery_paths:
            if not path.lower().endswith((".xlsx", ".xls")):
                continue
            try:
                xls = pd.ExcelFile(path)
            except Exception:
                continue
            for sheet_name in ("Review Decisions", "Approved Crops"):
                if sheet_name not in xls.sheet_names:
                    continue
                try:
                    df = pd.read_excel(path, sheet_name=sheet_name)
                except Exception:
                    continue
                if not portable_sources and sheet_name == "Review Decisions":
                    # Older jobs may predate the portable JSON.  The modal
                    # per-crop phenotype is the safest available fallback for
                    # reconstructing each global channel role; explicit
                    # per-crop roles are restored below for every matched row.
                    for channel_name in CHANNELS:
                        column = f"{channel_name.title()}_Channel_Role_For_Crop"
                        if column not in df.columns:
                            continue
                        valid_roles = [
                            str(value).strip().lower()
                            for value in df[column].dropna().tolist()
                            if str(value).strip().lower() in CHANNEL_ROLE_OPTIONS
                        ]
                        if valid_roles:
                            state[f"{channel_name}_channel_role"] = max(
                                set(valid_roles), key=valid_roles.count
                            )
                if not portable_sources and sheet_name == "Approved Crops" and not df.empty:
                    global_cleaning_columns = {
                        "Cleaning_Radius_px": ("cleaning_radius_px", 14, 4, 40),
                        "Cleaning_Sensitivity": ("cleaning_sensitivity", 50, 1, 100),
                        "Cleaning_Min_Area_px": ("cleaning_min_area_px", 3, 1, 100),
                        "Cleaning_Background_Removal_Percent": ("cleaning_background_removal", 0, 0, 100),
                        "Cleaning_Amount_Percent": ("cleaning_amount", 100, 0, 100),
                    }
                    for column, (state_key, fallback, minimum, maximum) in global_cleaning_columns.items():
                        if column in df.columns and not df[column].dropna().empty:
                            state[state_key] = recovery_int_value(
                                df[column].dropna().mode().iloc[0],
                                state.get(state_key, fallback), minimum, maximum,
                            )
                for _, row in df.iterrows():
                    status = recovery_status_from_text(row.get("Status", ""))
                    text = " ".join(
                        str(row.get(column, "") or "")
                        for column in (
                            "Selected_Crop_File",
                            "Original_Review_Item",
                            "Original_Crop_File",
                            "Approved_Crop_File",
                            "Review_Relative_Path",
                        )
                    )
                    for rel in items:
                        token = recovery_crop_token(rel)
                        if not token or token not in text:
                            continue
                        if status:
                            recovered_status[rel] = status
                        elif sheet_name == "Approved Crops":
                            recovered_status[rel] = "keep"
                        recovered_settings = backup_item_settings.setdefault(rel, {})
                        recovered_roles = {}
                        for channel_name in CHANNELS:
                            role_value = str(
                                row.get(f"{channel_name.title()}_Channel_Role_For_Crop", "") or ""
                            ).strip().lower()
                            if role_value in CHANNEL_ROLE_OPTIONS:
                                recovered_roles[channel_name] = role_value
                        if recovered_roles:
                            recovered_settings["channel_roles"] = recovered_roles
                        if sheet_name == "Approved Crops":
                            cleaning_source = str(row.get("Cleaning_Settings_Source", "") or "").strip().lower()
                            if "per-crop" in cleaning_source or "per crop" in cleaning_source:
                                cleaned_channels_text = str(row.get("Cleaned_Channels", "") or "").lower()
                                recovered_settings["recovered_per_crop_cleaning_params"] = {
                                    "radius_px": recovery_int_value(
                                        row.get("Cleaning_Radius_px"),
                                        state.get("cleaning_radius_px", 14), 4, 40,
                                    ),
                                    "sensitivity": recovery_int_value(
                                        row.get("Cleaning_Sensitivity"),
                                        state.get("cleaning_sensitivity", 50), 1, 100,
                                    ),
                                    "min_area_px": recovery_int_value(
                                        row.get("Cleaning_Min_Area_px"),
                                        state.get("cleaning_min_area_px", 3), 1, 100,
                                    ),
                                    "background_removal": recovery_int_value(
                                        row.get("Cleaning_Background_Removal_Percent"),
                                        state.get("cleaning_background_removal", 0), 0, 100,
                                    ),
                                    "amount": recovery_int_value(
                                        row.get("Cleaning_Amount_Percent"),
                                        state.get("cleaning_amount", 100), 0, 100,
                                    ),
                                    "channels": [
                                        channel_name
                                        for channel_name in CHANNELS
                                        if channel_name in cleaned_channels_text
                                    ],
                                }
                            analysis_input_path = str(row.get("Analysis_Input_File", "") or "").strip()
                            if analysis_input_path and os.path.isfile(analysis_input_path):
                                recovered_settings["auto_cleaned_crop_path"] = analysis_input_path
                            preview_path = str(row.get("Channel_Cleaning_Preview_PNG", "") or "").strip()
                            if preview_path and os.path.isfile(preview_path):
                                recovered_settings["auto_cleaned_preview_path"] = preview_path
                            input_mode = str(row.get("Analysis_Input_Mode", "") or "").strip()
                            if input_mode:
                                recovered_settings["auto_cleaned_input_mode"] = input_mode
                                if normalize_analysis_input_mode(input_mode) == "auto_cleaned":
                                    state["analysis_input_mode"] = "auto_cleaned"
                                    state["pcc_analysis_input_mode"] = "auto_cleaned"
                                    state["count_size_analysis_input_mode"] = "auto_cleaned"
                        for column in ("Approved_Crop_File", "Selected_Crop_File", "Analysis_Input_File"):
                            approved_path = str(row.get(column, "") or "").strip()
                            if approved_path and os.path.isfile(approved_path):
                                approved_matches[rel] = approved_path
                                break

    # Use loose approved TIFFs only to resolve the selected image for decisions
    # already established by a portable review state or workbook. If neither
    # exists, the directory itself is a last-resort keep list.
    loose_approved = [
        path for path in recovery_paths
        if path.lower().endswith((".tif", ".tiff")) and recovery_is_approved_crop_file(path)
    ]
    if recovered_status:
        for path in loose_approved:
            name = os.path.basename(path)
            for rel, status in recovered_status.items():
                token = recovery_crop_token(rel)
                if status == "keep" and token and token in name:
                    approved_matches.setdefault(rel, path)
    else:
        for path in loose_approved:
            name = os.path.basename(path)
            for rel in items:
                token = recovery_crop_token(rel)
                if token and token in name:
                    approved_matches[rel] = path
                    recovered_status[rel] = "keep"

    if not recovered_status:
        raise ValueError("No matching previous crop decisions could be recovered from the selected experiment folder.")

    manual_matches = recovery_manual_corrected_matches(state, items)
    kept = rejected = manual = 0
    for rel, item in items.items():
        item_settings = backup_item_settings.get(rel, {}) if isinstance(backup_item_settings, dict) else {}
        if isinstance(item_settings.get("channel_roles"), dict):
            item["channel_roles"] = copy.deepcopy(item_settings["channel_roles"])
        if "explicit_keep" in item_settings:
            item["explicit_keep"] = bool(item_settings["explicit_keep"])
        recovered_cleaning = item_settings.get("recovered_per_crop_cleaning_params")
        if isinstance(recovered_cleaning, dict):
            per_crop_cleaning = state.setdefault("per_crop_cleaning_params", {})
            if not isinstance(per_crop_cleaning, dict):
                per_crop_cleaning = {}
                state["per_crop_cleaning_params"] = per_crop_cleaning
            per_crop_cleaning[rel] = copy.deepcopy(recovered_cleaning)
        status = recovered_status.get(rel, "reject")
        item["status"] = status
        approved_path = approved_matches.get(rel, "")
        portable_manual_path = str(item_settings.get("manual_crop_path", "") or "")
        manual_path = portable_manual_path if os.path.isfile(portable_manual_path) else manual_matches.get(rel, "")
        recovered_manual_path = ""
        if status == "keep" and (manual_path or (approved_path and is_manual_recrop_file(approved_path))):
            recovered_manual_path = manual_path or approved_path
            item["path"] = recovered_manual_path
            item["manual_crop_path"] = recovered_manual_path
            item["manual_original_path"] = os.path.join(state.get("crop_root", ""), rel)
            item["crop_choice"] = "manual_recrop"
            item["note"] = (
                "Recovered from previous manual-corrected crop."
                if manual_path
                else "Recovered from previous approved manual-recrops output."
            )
            manual += 1
        if status == "keep":
            kept += 1
            if not recovered_manual_path:
                item.pop("manual_crop_path", None)
                item.pop("manual_original_path", None)
                item["path"] = os.path.join(state.get("crop_root", ""), rel)
                item["crop_choice"] = "raw"
                item["note"] = item.get("note", "")
        else:
            rejected += 1
            if recovered_manual_path:
                item["note"] = item.get("note", "Recovered previous manual-corrected crop; status is rejected.")
            else:
                item.pop("manual_crop_path", None)
                item.pop("manual_original_path", None)
                item["path"] = os.path.join(state.get("crop_root", ""), rel)
                item["crop_choice"] = "raw"
                item["note"] = "Recovered as rejected because it was not found in previous approved outputs."
        for key in PORTABLE_JOB_ITEM_KEYS:
            if key in ("status", "explicit_keep", "crop_choice", "note", "channel_roles", "manual_crop_path", "manual_original_path"):
                continue
            if key not in item_settings:
                continue
            value = copy.deepcopy(item_settings[key])
            if key.endswith("_path") and value and not os.path.exists(str(value)):
                continue
            item[key] = value

    state["last_scan"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state["last_recovery_summary"] = {
        "loaded_crops": count,
        "kept": kept,
        "rejected": rejected,
        "manual_recrops": manual,
        "sources": backup_sources + recovery_paths,
    }
    save_state(state)
    return count, kept, rejected, manual, backup_sources + recovery_paths


def reload_existing_results_for_current_folder(state):
    """Best-effort reload of previous work for the selected raw folder."""
    refresh_derived_paths(state)
    output_root = state.get("output_root", "")
    pcc_root = os.path.join(output_root, "real_fiji_pcc_manders_analysis")
    count_size_root = os.path.join(output_root, "real_fiji_condensate_count_size_analysis")
    fixed_cell_root = os.path.join(output_root, "fixed_cell_ms2_condensate_analysis")
    partition_root = os.path.join(output_root, "manual_nuclear_partition_coefficient_analysis")
    existing_outputs = {
        "last_real_fiji_pcc_manders_excel": find_latest_result_excel(pcc_root, "pcc_manders"),
        "last_pcc_manders_preview_root": os.path.join(pcc_root, "preview_images_sent_to_fiji"),
        "last_pcc_manders_approved_root": os.path.join(output_root, "approved_crops_for_real_fiji_pcc_manders"),
        "last_count_size_excel": find_latest_result_excel(count_size_root, "count_size"),
        "last_count_size_output_root": count_size_root,
        "last_count_size_approved_root": os.path.join(output_root, "approved_crops_for_condensate_count_size"),
        "last_fixed_cell_ms2_excel": os.path.join(fixed_cell_root, "Fixed_Cell_MS2_Condensate_Per_Cell_Measurements.xlsx"),
        "last_fixed_cell_ms2_output_root": fixed_cell_root,
        "last_fixed_cell_ms2_preview_root": os.path.join(fixed_cell_root, "previews"),
        "last_partition_excel": os.path.join(partition_root, "Manual_Nuclear_Partition_Coefficient_Measurements.xlsx"),
        "last_partition_output_root": partition_root,
        "last_partition_preview_root": os.path.join(partition_root, "previews"),
    }
    for key, path in existing_outputs.items():
        set_existing_path_or_blank(state, key, path)
    if os.path.isdir(existing_outputs["last_pcc_manders_preview_root"]):
        state["final_preview_root"] = existing_outputs["last_pcc_manders_preview_root"]

    reloaded_autocleaned = 0
    try:
        count, kept, rejected, manual, sources = recover_review_state_from_existing_outputs(state)
        reloaded_autocleaned = reload_auto_cleaned_inputs_for_loaded(state)
        state["last_auto_reload_summary"] = {
            "status": "recovered",
            "loaded_crops": count,
            "kept": kept,
            "rejected": rejected,
            "manual_recrops": manual,
            "auto_cleaned_reloaded": reloaded_autocleaned,
            "sources": sources,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        save_state(state)
        return state["last_auto_reload_summary"]
    except Exception as error:
        # A job may have saved channel-cleaned crops without producing a
        # PCC/count-size workbook yet.  Do not discard those cached crops
        # merely because decision recovery has no workbook to inspect.
        try:
            reloaded_autocleaned = reload_auto_cleaned_inputs_for_loaded(state)
        except Exception:
            reloaded_autocleaned = 0
        state["last_auto_reload_summary"] = {
            "status": "loaded_without_previous_decisions",
            "error": str(error),
            "auto_cleaned_reloaded": reloaded_autocleaned,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        save_state(state)
        return state["last_auto_reload_summary"]


def raw_and_enhanced_crop_roots(state):
    raw_folder, preprocessor_name, crop_name, output_name, preprocessor_root, _, _ = named_pipeline_paths(state)
    crop_base = os.path.join(preprocessor_root, crop_name)
    raw_crop_root = os.path.join(crop_base, "01_raw_masked_roi_crops_for_jacop")
    enhanced_crop_root = os.path.join(crop_base, "02_enhanced_roi_crops_for_optional_analysis")
    return raw_crop_root, enhanced_crop_root


def crop_preview_items(root):
    if not os.path.isdir(root):
        return []
    return [
        {"rel_path": os.path.relpath(path, root), "path": path}
        for path in find_tiffs(root)
    ]


def crop_index_path(state):
    raw_folder, preprocessor_name, crop_name, output_name, preprocessor_root, _, _ = named_pipeline_paths(state)
    crop_base = os.path.join(preprocessor_root, crop_name)
    raw_only_path = os.path.join(crop_base, "Raw_Cropped_ROI_Index.xlsx")
    old_path = os.path.join(crop_base, "Raw_vs_Enhanced_Cropped_ROI_Index.xlsx")
    return raw_only_path if os.path.isfile(raw_only_path) else old_path


def load_crop_index_context(state):
    index_path = crop_index_path(state)
    if not os.path.isfile(index_path):
        return {}
    try:
        index_stat = os.stat(index_path)
        cache_key = (index_path, index_stat.st_mtime_ns, index_stat.st_size)
    except OSError:
        cache_key = (index_path, 0, 0)
    if cache_key in CROP_INDEX_CONTEXT_CACHE:
        return CROP_INDEX_CONTEXT_CACHE[cache_key]
    try:
        df = pd.read_excel(index_path)
    except Exception:
        return {}

    raw_crop_root, enhanced_crop_root = raw_and_enhanced_crop_roots(state)
    preprocessor_root = os.path.dirname(os.path.dirname(raw_crop_root))

    def relocated_index_path(saved_path, current_root):
        """Resolve crop-index paths after an experiment folder has been moved."""
        saved_path = str(saved_path or "").strip()
        if not saved_path:
            return "", ""
        normalized = saved_path.replace("\\", "/")
        marker = "/" + os.path.basename(current_root).lower() + "/"
        marker_at = normalized.lower().find(marker)
        if marker_at >= 0:
            rel_key = normalized[marker_at + len(marker):]
        else:
            try:
                rel_key = os.path.relpath(saved_path, current_root)
            except Exception:
                rel_key = os.path.basename(saved_path)
        candidate = os.path.abspath(os.path.join(current_root, rel_key))
        return (candidate if os.path.isfile(candidate) else saved_path), rel_key

    def relocated_source_path(saved_path, folder_name):
        saved_path = str(saved_path or "").strip()
        if not saved_path:
            return ""
        if os.path.isfile(saved_path):
            return saved_path
        candidate = os.path.join(preprocessor_root, folder_name, os.path.basename(saved_path))
        return candidate if os.path.isfile(candidate) else saved_path

    context = {}
    for row in df.to_dict("records"):
        raw_crop = str(row.get("Raw_Masked_Crop_File", "") or "")
        enhanced_crop = str(row.get("Enhanced_Crop_File", "") or row.get("Optional_Enhanced_Analysis_File", "") or "")
        raw_source = str(row.get("Raw_Masked_Source_Image", "") or "")
        enhanced_source = str(row.get("Enhanced_Source_Image", "") or "")
        image_name = str(row.get("Image", "") or "")
        bbox = {
            "y0": int(row.get("Crop_Y0", 0) or 0),
            "y1": int(row.get("Crop_Y1", 0) or 0),
            "x0": int(row.get("Crop_X0", 0) or 0),
            "x1": int(row.get("Crop_X1", 0) or 0),
        }
        raw_crop, raw_rel_key = relocated_index_path(raw_crop, raw_crop_root)
        enhanced_crop, enhanced_rel_key = relocated_index_path(enhanced_crop, enhanced_crop_root)
        raw_source = relocated_source_path(raw_source, "01_raw_masked_for_jacop")
        enhanced_source = relocated_source_path(enhanced_source, "02_enhanced_for_visual_qc")
        payload = {
            "raw_crop": raw_crop,
            "enhanced_crop": enhanced_crop,
            "raw_source": raw_source,
            "enhanced_source": enhanced_source,
            "image_name": image_name,
            "bbox": bbox,
        }
        for crop_path, rel_key in ((raw_crop, raw_rel_key), (enhanced_crop, enhanced_rel_key)):
            if crop_path:
                context[rel_key] = payload
                context.setdefault(os.path.basename(crop_path), payload)
    CROP_INDEX_CONTEXT_CACHE.clear()
    CROP_INDEX_CONTEXT_CACHE[cache_key] = context
    return context


def crop_context_for_rel(state, rel_path):
    return load_crop_index_context(state).get(rel_path, {})


def selected_crop_path_for_item(state, rel_path, item):
    manual_crop_path = resolve_appledouble_image_path(item.get("manual_crop_path"))
    if manual_crop_path and os.path.isfile(manual_crop_path):
        return manual_crop_path, "manual_recrop"
    context = crop_context_for_rel(state, rel_path)
    choice = item.get("crop_choice", "raw")
    if choice == "enhanced" and context.get("enhanced_crop"):
        return context["enhanced_crop"], "enhanced"
    if context.get("raw_crop"):
        return context["raw_crop"], "raw"
    return item.get("path") or os.path.join(state.get("crop_root", ""), rel_path), choice


def thumbnail_cache_key(kind, path, max_size, extra=""):
    try:
        stat = os.stat(path)
        stamp = f"{stat.st_mtime_ns}:{stat.st_size}"
    except OSError:
        stamp = "missing"
    raw = f"{kind}|{os.path.abspath(str(path or ''))}|{max_size}|{extra}|{stamp}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def cached_png_bytes(kind, path, max_size, producer, extra=""):
    cache_path = ""
    try:
        os.makedirs(THUMBNAIL_CACHE_ROOT, exist_ok=True)
        cache_dir = os.path.join(THUMBNAIL_CACHE_ROOT, safe_name(kind))
        os.makedirs(cache_dir, exist_ok=True)
        cache_path = os.path.join(cache_dir, thumbnail_cache_key(kind, path, max_size, extra) + ".png")
        if os.path.isfile(cache_path):
            with open(cache_path, "rb") as f:
                return f.read()
    except OSError:
        cache_path = ""
    data = producer()
    if cache_path:
        try:
            tmp_path = cache_path + f".{uuid4().hex}.tmp"
            with open(tmp_path, "wb") as f:
                f.write(data)
            os.replace(tmp_path, cache_path)
        except OSError:
            try:
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except OSError:
                pass
    return data


def png_bytes_for_index_path(path, max_size=520):
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(path or "Missing image path")
    return cached_png_bytes("index_crop_thumb", path, max_size, lambda: png_bytes_for_tiff_uncached(path, max_size=max_size))


def png_bytes_for_context_path(path, bbox, max_size=520):
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(path or "Missing context image path")
    extra = json.dumps(bbox or {}, sort_keys=True)
    def build():
        image, _ = read_yxc(path)
        rgb = display_rgb(image)
        pil = Image.fromarray(rgb)
        draw = ImageDraw.Draw(pil)
        try:
            x0 = int(bbox.get("x0", 0)); x1 = int(bbox.get("x1", 0))
            y0 = int(bbox.get("y0", 0)); y1 = int(bbox.get("y1", 0))
        except Exception:
            x0 = x1 = y0 = y1 = 0
        if x1 > x0 and y1 > y0:
            for offset in range(5):
                draw.rectangle([x0-offset, y0-offset, x1+offset, y1+offset], outline=(0, 255, 255))
        pil.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        pil.save(buf, format="PNG")
        return buf.getvalue()
    return cached_png_bytes("context_thumb", path, max_size, build, extra=extra)


def crop_image_base_from_rel(rel_path):
    parent = os.path.dirname(rel_path)
    if parent:
        return os.path.basename(parent)
    name = os.path.splitext(os.path.basename(rel_path))[0]
    for token in ("_MARKER_COMPONENT_", "_ROI_", "_CELL_"):
        if token in name:
            return name.split(token, 1)[0]
    return name


def crop_source_key(rel_path):
    return crop_image_base_from_rel(rel_path)


def raw_m_base_candidates(image_base):
    candidates = []
    base = image_base
    if " Slide_M_" in base:
        candidates.append(base)
    for token in (" Slide_D_", " Slide_R_", " Slide_G_", " Slide_B_"):
        if token in base:
            candidates.append(base.replace(token, " Slide_M_", 1))
    # Some experiments have only one M field and use d0 for the merged image.
    for candidate in list(candidates):
        if "d" in candidate:
            prefix, dpart = candidate.rsplit("d", 1)
            if dpart.isdigit():
                candidates.append(prefix + "d0")
    unique = []
    for candidate in candidates:
        if candidate and candidate not in unique:
            unique.append(candidate)
    return unique


def raw_m_image_index(state):
    raw_root = state.get("raw_image_folder", "")
    if not raw_root or not os.path.isdir(raw_root):
        return {}
    try:
        root_stat = os.stat(raw_root)
        cache_key = (raw_root, root_stat.st_mtime_ns, root_stat.st_size)
    except OSError:
        cache_key = (raw_root, 0, 0)
    if cache_key in RAW_M_IMAGE_INDEX_CACHE:
        return RAW_M_IMAGE_INDEX_CACHE[cache_key]
    excluded_tokens = (
        "ENHANCED_JACOP_PREPROCESSOR_RESULTS",
        "REAL FIJI PCC MANDERS RESULTS",
        "pcc_manders_analysis",
    )
    index = {}
    for dirpath, dirnames, filenames in os.walk(raw_root):
        dirnames[:] = [dirname for dirname in dirnames if not any(token in dirname for token in excluded_tokens)]
        for filename in filenames:
            lower = filename.lower()
            if not lower.endswith((".tif", ".tiff")):
                continue
            stem = os.path.splitext(filename)[0]
            path = os.path.join(dirpath, filename)
            index[lower] = path
            index.setdefault(stem, path)
            index.setdefault(stem.lower(), path)
    RAW_M_IMAGE_INDEX_CACHE.clear()
    RAW_M_IMAGE_INDEX_CACHE[cache_key] = index
    return index


def find_raw_m_image(state, rel_path):
    candidates = raw_m_base_candidates(crop_image_base_from_rel(rel_path))
    if not candidates:
        return ""
    index = raw_m_image_index(state)
    for candidate in candidates:
        for ext in (".tif", ".tiff"):
            path = index.get((candidate + ext).lower())
            if path:
                return path
    for candidate in candidates:
        path = index.get(candidate) or index.get(candidate.lower())
        if path and "_Top Slide_M_" in os.path.splitext(os.path.basename(path))[0]:
            return path
    return ""


def png_bytes_for_raw_m(state, rel_path, max_size=520):
    path = find_raw_m_image(state, rel_path)
    if not path:
        raise FileNotFoundError(f"No matching raw M image found for {rel_path}")
    return png_bytes_for_tiff(path, max_size=max_size)


def choose_crop_set_for_review(state, choice):
    state["final_input_choice"] = "raw_masked"
    refresh_derived_paths(state)
    count = scan_crops(state)
    return count


def approved_crop_items(state):
    approved_root = os.path.join(state.get("output_root", ""), "approved_crops_for_real_fiji_pcc_manders")
    if not os.path.isdir(approved_root):
        return [], approved_root
    items = []
    for path in find_tiffs(approved_root):
        items.append({
            "rel_path": os.path.relpath(path, approved_root),
            "path": path,
        })
    return items, approved_root


def final_preview_png_items(state):
    preview_root = state.get("final_preview_root") or os.path.join(state.get("output_root", ""), "final_analyzed_roi_previews")
    if not os.path.isdir(preview_root):
        return [], preview_root
    items = []
    for dirpath, _, filenames in os.walk(preview_root):
        for filename in filenames:
            if not filename.startswith("._") and filename.lower().endswith(".png") and is_real_png_file(os.path.join(dirpath, filename)):
                path = os.path.join(dirpath, filename)
                items.append({"rel_path": os.path.relpath(path, preview_root), "path": path})
    return sorted(items, key=lambda item: item["rel_path"]), preview_root


def fixed_cell_preview_png_items(state):
    preview_root = state.get("last_fixed_cell_ms2_preview_root", "")
    if not preview_root or not os.path.isdir(preview_root):
        return [], preview_root
    items = []
    for dirpath, _, filenames in os.walk(preview_root):
        for filename in filenames:
            path = os.path.join(dirpath, filename)
            if not filename.startswith("._") and filename.lower().endswith(".png") and is_real_png_file(path):
                items.append({"rel_path": os.path.relpath(path, preview_root), "path": path})
    return sorted(items, key=lambda item: item["rel_path"]), preview_root


def channel_cleaning_preview_items(state):
    preview_root = os.path.join(
        state.get("output_root", ""),
        "channel_cleaned_analysis_inputs",
        "raw_vs_cleaned_channel_previews",
    )
    if not os.path.isdir(preview_root):
        return [], preview_root
    items = []
    for dirpath, _, filenames in os.walk(preview_root):
        for filename in filenames:
            if not filename.startswith("._") and filename.lower().endswith(".png") and is_real_png_file(os.path.join(dirpath, filename)):
                path = os.path.join(dirpath, filename)
                items.append({"rel_path": os.path.relpath(path, preview_root), "path": path})
    return sorted(items, key=lambda item: item["rel_path"]), preview_root



def read_yxc(path):
    image = np.asarray(tiff.imread(path))
    if image.ndim == 3 and image.shape[-1] >= 3:
        return np.array(image[..., :3], copy=True), "YXC"
    if image.ndim == 3 and image.shape[0] >= 3:
        return np.stack([image[0], image[1], image[2]], axis=-1), "CYX"
    if image.ndim == 4 and image.shape[1] >= 3:
        projected = np.stack([image[:, 0].max(axis=0), image[:, 1].max(axis=0), image[:, 2].max(axis=0)], axis=-1)
        return projected, "ZCYX_projected_to_YXC"
    if image.ndim == 4 and image.shape[0] >= 3:
        projected = np.stack([image[0].max(axis=0), image[1].max(axis=0), image[2].max(axis=0)], axis=-1)
        return projected, "CZYX_projected_to_YXC"
    raise ValueError(f"Could not read as 3-channel TIFF: {path}; shape={image.shape}")


def normalize_fixed_cell_intensity_source_mode(value):
    value = str(value or "reviewed_raw").strip().lower()
    return value if value in ("reviewed_raw", "untouched_acquisition") else "reviewed_raw"


def untouched_acquisition_path_for_item(state, rel_path, item):
    """Find one untouched acquisition matching a reviewed M-channel image."""
    raw_root = os.path.abspath(str(state.get("raw_image_folder", "") or ""))
    if not raw_root or not os.path.isdir(raw_root):
        raise ValueError("Select the folder containing the untouched acquisition TIFF images first.")
    rel_parent = os.path.basename(os.path.dirname(str(rel_path or ""))).strip()
    source_candidates = [
        rel_parent,
        os.path.splitext(os.path.basename(str(item.get("manual_original_path", "") or "")))[0],
        os.path.splitext(os.path.basename(str(item.get("path", "") or "")))[0],
    ]
    target_stems = []
    suffixes = (
        "_raw_masked_jacop_ready",
        "_enhanced_jacop_ready",
        "_raw_masked_crop_for_jacop",
    )
    for value in source_candidates:
        stem = str(value or "").strip()
        lower = stem.lower()
        for suffix in suffixes:
            if lower.endswith(suffix):
                stem = stem[: -len(suffix)]
                lower = stem.lower()
                break
        marker_at = lower.find("_marker_component_")
        if marker_at >= 0:
            stem = stem[:marker_at]
        if stem and stem.lower() not in {row.lower() for row in target_stems}:
            target_stems.append(stem)
    # Acquisition folders are commonly nested one directory per microscope
    # capture. Search the complete selected raw tree, but never treat a TIFF
    # generated by this app as an untouched acquisition.
    excluded_directory_names = {
        str(state.get("preprocessor_folder_name", DEFAULT_PREPROCESSOR_FOLDER_NAME) or "").strip().lower(),
        str(state.get("final_output_folder_name", DEFAULT_FINAL_OUTPUT_FOLDER_NAME) or "").strip().lower(),
        DEFAULT_PREPROCESSOR_FOLDER_NAME.lower(),
        DEFAULT_FINAL_OUTPUT_FOLDER_NAME.lower(),
        "manual_corrected_crops",
        "generated_runners",
        "channel_cleaned_analysis_inputs",
        "fixed_cell_ms2_condensate_analysis",
        "real_fiji_condensate_count_size_analysis",
        "real_fiji_pcc_manders_analysis",
        "approved_crops_for_real_fiji_pcc_manders",
        "approved_crops_for_real_fiji_count_size",
    }
    excluded_directory_names.discard("")
    available = []
    for dirpath, dirnames, filenames in os.walk(raw_root):
        dirnames[:] = [
            dirname
            for dirname in dirnames
            if not dirname.startswith(".")
            and dirname.lower() not in excluded_directory_names
            and dirname.lower() != "__macosx"
        ]
        for name in filenames:
            if name.startswith("._") or os.path.splitext(name)[1].lower() not in (".tif", ".tiff"):
                continue
            path = os.path.abspath(os.path.join(dirpath, name))
            available.append((os.path.splitext(name)[0].lower(), path))
    for target in target_stems:
        matches = [path for file_stem, path in available if file_stem == target.lower()]
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raise ValueError(
                "More than one untouched acquisition TIFF matched this reviewed image. "
                "Keep only one acquisition copy or select a narrower raw-image folder: "
                + "; ".join(matches)
            )
    expected = target_stems[0] if target_stems else "matching image name"
    raise ValueError(
        f"No untouched acquisition TIFF matching '{expected}' was found inside the selected raw-image folder or its acquisition subfolders: {raw_root}"
    )


def legacy_manual_crop_metadata_from_job_folder(state, rel_path, item):
    """Recover legacy manual-recrop coordinates from this experiment's saved state."""
    output_root = str(state.get("output_root", "") or "").strip()
    if not output_root:
        return {}
    legacy_state_path = os.path.join(output_root, "CONDENSATE_REVIEW_STATE.json")
    if not os.path.isfile(legacy_state_path):
        return {}
    try:
        with open(legacy_state_path, "r", encoding="utf-8") as handle:
            legacy_state = json.load(handle)
        legacy_item = legacy_state.get("items", {}).get(rel_path, {})
    except (OSError, ValueError, TypeError, AttributeError):
        return {}
    if not isinstance(legacy_item, dict):
        return {}
    current_manual = os.path.basename(str(item.get("manual_crop_path", "") or ""))
    legacy_manual = os.path.basename(str(legacy_item.get("manual_crop_path", "") or ""))
    if current_manual and legacy_manual and current_manual != legacy_manual:
        return {}
    return legacy_item


def locate_legacy_manual_crop_on_untouched(state, rel_path, item):
    """Recover a legacy manual recrop's full-image box by exact raw-pixel matching."""
    manual_path = resolve_appledouble_image_path(item.get("manual_crop_path"))
    if not manual_path or not os.path.isfile(manual_path):
        raise ValueError("The recovered manual crop file is missing.")
    untouched_path = untouched_acquisition_path_for_item(state, rel_path, item)
    try:
        manual_stat = os.stat(manual_path)
        untouched_stat = os.stat(untouched_path)
        cache_key = (
            os.path.abspath(manual_path), manual_stat.st_mtime_ns, manual_stat.st_size,
            os.path.abspath(untouched_path), untouched_stat.st_mtime_ns, untouched_stat.st_size,
        )
    except OSError:
        cache_key = (os.path.abspath(manual_path), os.path.abspath(untouched_path))
    cached = LEGACY_CROP_MATCH_CACHE.get(cache_key)
    if cached:
        return dict(cached)

    crop_yxc, _ = read_yxc(manual_path)
    full_yxc, _ = read_yxc(untouched_path)
    crop_height, crop_width = crop_yxc.shape[:2]
    full_height, full_width = full_yxc.shape[:2]
    if crop_height > full_height or crop_width > full_width:
        raise ValueError("The recovered manual crop is larger than its matching untouched acquisition.")

    nonzero_mask = np.any(crop_yxc != 0, axis=2)
    nonzero_y, nonzero_x = np.nonzero(nonzero_mask)
    if not nonzero_x.size:
        raise ValueError("The recovered manual crop contains no nonzero raw pixels that can be matched.")

    # Find a rare exact RGB pixel to generate only a few candidate offsets.
    sample_indices = np.linspace(0, nonzero_x.size - 1, min(512, nonzero_x.size), dtype=int)
    best_anchor = None
    for sample_index in sample_indices:
        anchor_y = int(nonzero_y[sample_index])
        anchor_x = int(nonzero_x[sample_index])
        anchor_value = crop_yxc[anchor_y, anchor_x]
        hits = np.argwhere(np.all(full_yxc == anchor_value, axis=2))
        if hits.size and (best_anchor is None or len(hits) < len(best_anchor[2])):
            best_anchor = (anchor_y, anchor_x, hits)
            if len(hits) == 1:
                break
    if best_anchor is None:
        raise ValueError("No exact raw-pixel anchor from the recovered crop was found in the untouched acquisition.")

    anchor_y, anchor_x, hits = best_anchor
    matches = []
    for hit_y, hit_x in hits:
        y0 = int(hit_y) - anchor_y
        x0 = int(hit_x) - anchor_x
        if x0 < 0 or y0 < 0 or x0 + crop_width > full_width or y0 + crop_height > full_height:
            continue
        candidate = full_yxc[y0:y0 + crop_height, x0:x0 + crop_width]
        exact_fraction = float(np.mean(np.all(candidate[nonzero_mask] == crop_yxc[nonzero_mask], axis=1)))
        if exact_fraction >= 0.999:
            matches.append((exact_fraction, x0, y0))
    if not matches:
        raise ValueError("The recovered manual crop could not be aligned exactly to the untouched acquisition.")
    matches.sort(reverse=True)
    best_fraction, x0, y0 = matches[0]
    tied = [row for row in matches if abs(row[0] - best_fraction) < 1e-12]
    if len(tied) > 1:
        raise ValueError("The recovered manual crop matched more than one location in the untouched acquisition.")

    result = {
        "x0": int(x0), "y0": int(y0),
        "x1": int(x0 + crop_width), "y1": int(y0 + crop_height),
        "width": int(crop_width), "height": int(crop_height),
        "mapping": f"Legacy manual crop matched to untouched raw pixels ({best_fraction:.3%} exact)",
    }
    LEGACY_CROP_MATCH_CACHE[cache_key] = dict(result)
    return result


def reviewed_crop_bounds_on_full_image(state, rel_path, item, full_shape=None):
    """Map the exact reviewed PCC/count-size crop back to full acquisition coordinates."""
    context = crop_context_for_rel(state, rel_path)
    bbox = context.get("bbox") if isinstance(context, dict) else None
    try:
        base_x0 = int(bbox.get("x0")); base_y0 = int(bbox.get("y0"))
        base_x1 = int(bbox.get("x1")); base_y1 = int(bbox.get("y1"))
    except (AttributeError, TypeError, ValueError):
        base_x0 = base_y0 = base_x1 = base_y1 = 0
    if base_x1 <= base_x0 or base_y1 <= base_y0:
        raise ValueError(
            "The reviewed crop has no valid full-image coordinates in the crop index. "
            "Rerun the crop-generation stage for this experiment before drawing on the untouched image."
        )

    mapping = "Crop-index coordinates used by PCC/Manders and Count/Size"
    if item.get("manual_crop_path") and item.get("crop_choice") == "manual_recrop":
        legacy_item = {}
        saved_values = (
            item.get("manual_crop_origin_x"),
            item.get("manual_crop_origin_y"),
            item.get("manual_crop_width"),
            item.get("manual_crop_height"),
        )
        if any(value is None for value in saved_values):
            legacy_item = legacy_manual_crop_metadata_from_job_folder(state, rel_path, item)
            saved_values = (
                legacy_item.get("manual_crop_origin_x"),
                legacy_item.get("manual_crop_origin_y"),
                legacy_item.get("manual_crop_width"),
                legacy_item.get("manual_crop_height"),
            )
        note = str(item.get("note", "") or "")
        legacy_note = str(legacy_item.get("note", "") or "")
        coordinate_note = legacy_note or note
        if any(value is None for value in saved_values):
            legacy_match = re.search(
                r"x=(\d+),\s*y=(\d+),\s*w=(\d+),\s*h=(\d+)",
                coordinate_note,
                flags=re.IGNORECASE,
            )
            if legacy_match:
                saved_values = tuple(int(value) for value in legacy_match.groups())
        try:
            local_x, local_y, width, height = [int(value) for value in saved_values]
        except (TypeError, ValueError):
            matched = locate_legacy_manual_crop_on_untouched(state, rel_path, item)
            if full_shape is not None:
                full_height, full_width = full_shape[:2]
                if matched["x1"] > full_width or matched["y1"] > full_height:
                    raise ValueError("The recovered manual crop match falls outside the untouched acquisition.")
            return matched
        manual_source_value = item.get("manual_crop_source_mode") or legacy_item.get("manual_crop_source_mode")
        if manual_source_value:
            manual_source = str(manual_source_value).lower()
        elif re.search(r"saved\s+from\s+original\b", coordinate_note, flags=re.IGNORECASE):
            manual_source = "original"
        elif re.search(r"saved\s+from\s+crop\b", coordinate_note, flags=re.IGNORECASE):
            manual_source = "crop"
        else:
            manual_source = "crop"
        if manual_source == "original":
            x0, y0 = local_x, local_y
            mapping = "Manual reviewed crop drawn on the full source image (legacy coordinates recovered when needed)"
        else:
            x0, y0 = base_x0 + local_x, base_y0 + local_y
            mapping = "Manual reviewed crop composed with its crop-index parent coordinates (legacy coordinates recovered when needed)"
        x1, y1 = x0 + width, y0 + height
    else:
        x0, y0, x1, y1 = base_x0, base_y0, base_x1, base_y1

    if full_shape is not None:
        full_height, full_width = full_shape[:2]
        if x0 < 0 or y0 < 0 or x1 > full_width or y1 > full_height:
            raise ValueError(
                f"The reviewed crop box x={x0}:{x1}, y={y0}:{y1} falls outside the full image "
                f"({full_width}x{full_height})."
            )
    return {
        "x0": int(x0), "y0": int(y0), "x1": int(x1), "y1": int(y1),
        "width": int(x1 - x0), "height": int(y1 - y0), "mapping": mapping,
    }


def require_polygons_inside_reviewed_crop(polygons, crop_box, label_text):
    """Prevent a full-image manual annotation from drifting into neighboring cells."""
    if not crop_box:
        return
    x0, y0 = float(crop_box["x0"]), float(crop_box["y0"])
    x1, y1 = float(crop_box["x1"]), float(crop_box["y1"])
    for polygon_index, points in enumerate(polygons, start=1):
        for x, y in points:
            if x < x0 or x >= x1 or y < y0 or y >= y1:
                raise ValueError(
                    f"{label_text} {polygon_index} extends outside the reviewed PCC/Count-Size cell box. "
                    "Keep every point inside the yellow crop boundary."
                )


def manual_crop_bounds_for_untouched(item, crop_shape, untouched_shape):
    """Return the full-image bounds for a manual crop, including legacy saved items."""
    values = (
        item.get("manual_crop_origin_x"),
        item.get("manual_crop_origin_y"),
        item.get("manual_crop_width"),
        item.get("manual_crop_height"),
    )
    if any(value is None for value in values):
        match = re.search(
            r"x=(\d+),\s*y=(\d+),\s*w=(\d+),\s*h=(\d+)",
            str(item.get("note", "") or ""),
            flags=re.IGNORECASE,
        )
        if match:
            values = tuple(int(value) for value in match.groups())
    try:
        x0, y0, width, height = [int(value) for value in values]
    except (TypeError, ValueError):
        raise ValueError(
            "This cropped ROI has no reliable full-image coordinates. Redraw it from the original image, "
            "draw the fixed-cell ROI on the original image, or use the reviewed raw intensity source."
        )
    crop_height, crop_width = crop_shape[:2]
    full_height, full_width = untouched_shape[:2]
    if width != crop_width or height != crop_height:
        raise ValueError(
            f"The saved crop bounds are {width}x{height}, but the reviewed crop is {crop_width}x{crop_height}."
        )
    if x0 < 0 or y0 < 0 or x0 + width > full_width or y0 + height > full_height:
        raise ValueError("The saved crop bounds fall outside the untouched acquisition image.")
    return x0, y0, width, height


def validated_fixed_cell_intensity_image(state, rel_path, item, drawing_source_path, drawing_yxc, source_mode):
    """Return fixed-cell intensity pixels without changing PCC/Manders or count/size inputs."""
    mode = normalize_fixed_cell_intensity_source_mode(state.get("fixed_cell_intensity_source_mode"))
    if mode == "reviewed_raw":
        return {
            "image": np.asarray(drawing_yxc),
            "path": os.path.abspath(drawing_source_path),
            "layout": "same_as_drawing_source",
            "mapping": "Direct reviewed raw source coordinates",
            "validation": "Not required; intensity and drawing source are identical",
        }
    untouched_path = untouched_acquisition_path_for_item(state, rel_path, item)
    untouched_yxc, untouched_layout = read_yxc(untouched_path)
    source_mode = str(source_mode or "crop").lower()
    if source_mode in ("original", "untouched"):
        if untouched_yxc.shape != drawing_yxc.shape:
            raise ValueError(
                f"Untouched acquisition shape {untouched_yxc.shape} does not match the original drawing source {drawing_yxc.shape}."
            )
        mapped = np.asarray(untouched_yxc)
        mapping = "Full-image coordinates" if source_mode == "original" else "Direct untouched-acquisition coordinates"
    else:
        crop_box = reviewed_crop_bounds_on_full_image(state, rel_path, item, untouched_yxc.shape)
        x0, y0, width, height = crop_box["x0"], crop_box["y0"], crop_box["width"], crop_box["height"]
        mapped = np.asarray(untouched_yxc[y0:y0 + height, x0:x0 + width, :])
        mapping = f"Reviewed PCC/Count-Size crop mapped to untouched image: x={x0}, y={y0}, w={width}, h={height}"
    if mapped.shape != drawing_yxc.shape:
        raise ValueError(
            f"Mapped untouched intensity image shape {mapped.shape} does not match the drawing source {drawing_yxc.shape}."
        )
    if source_mode == "untouched" and os.path.abspath(drawing_source_path) == os.path.abspath(untouched_path):
        exact_fraction = 1.0
        validation = "Drawing and intensity pixels are the same untouched acquisition TIFF"
        return {
            "image": mapped, "path": untouched_path, "layout": untouched_layout,
            "mapping": mapping, "validation": validation,
        }
    drawing_values = np.asarray(drawing_yxc)
    nonzero = np.any(drawing_values != 0, axis=2)
    if not np.any(nonzero):
        raise ValueError("The reviewed drawing source has no nonzero pixels for untouched-image alignment validation.")
    exact_fraction = float(np.mean(np.all(mapped[nonzero] == drawing_values[nonzero], axis=1)))
    if exact_fraction < 0.999:
        raise ValueError(
            "The untouched acquisition failed strict alignment/channel validation against the reviewed raw source "
            f"({exact_fraction:.3%} of nonzero RGB pixels matched; 99.9% required)."
        )
    return {
        "image": mapped,
        "path": untouched_path,
        "layout": untouched_layout,
        "mapping": mapping,
        "validation": f"Passed strict alignment/channel validation ({exact_fraction:.3%} exact nonzero RGB match)",
    }


def display_rgb(image_yxc):
    image = np.asarray(image_yxc)
    out = []
    for i in range(3):
        channel = image[..., i].astype(np.float64)
        nonzero = channel[channel > 0]
        if nonzero.size:
            low, high = np.percentile(nonzero, [1, 99.8])
        else:
            low, high = np.percentile(channel, [1, 99.8])
        if high <= low:
            out.append(np.zeros(channel.shape, dtype=np.uint8))
        else:
            scaled = np.clip((channel - low) / (high - low), 0, 1)
            out.append((scaled * 255).astype(np.uint8))
    return np.stack(out, axis=-1)


def direct_raw_preview_rgb(image_yxc):
    """Create a browser preview without per-channel contrast stretching.

    Untouched 8-bit acquisitions pass through exactly. Higher-bit images are
    linearly mapped by their storage range solely because the browser preview
    is an 8-bit PNG; the measurement array is never modified.
    """
    image = np.asarray(image_yxc)
    if image.dtype == np.uint8:
        return np.array(image, copy=True)
    if np.issubdtype(image.dtype, np.integer):
        info = np.iinfo(image.dtype)
        scale = 255.0 / max(1.0, float(info.max - info.min))
        return np.clip((image.astype(np.float64) - float(info.min)) * scale, 0, 255).astype(np.uint8)
    values = np.nan_to_num(image.astype(np.float64), nan=0.0, posinf=0.0, neginf=0.0)
    finite_min = float(np.min(values)) if values.size else 0.0
    finite_max = float(np.max(values)) if values.size else 1.0
    if finite_min >= 0.0 and finite_max <= 1.0:
        return np.clip(values * 255.0, 0, 255).astype(np.uint8)
    if finite_min >= 0.0 and finite_max <= 255.0:
        return np.clip(values, 0, 255).astype(np.uint8)
    scale = 255.0 / max(1e-12, finite_max - finite_min)
    return np.clip((values - finite_min) * scale, 0, 255).astype(np.uint8)


def png_bytes_for_tiff_uncached(path, max_size=420):
    image, _ = read_yxc(path)
    rgb = display_rgb(image)
    pil = Image.fromarray(rgb)
    pil.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    pil.save(buf, format="PNG")
    return buf.getvalue()


def png_bytes_for_tiff(path, max_size=420):
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(path or "Missing TIFF image path")
    return cached_png_bytes("tiff_thumb", path, max_size, lambda: png_bytes_for_tiff_uncached(path, max_size=max_size))


def manual_crop_image_path(state, rel_path, source_mode="crop"):
    item = state.get("items", {}).get(rel_path, {})
    if source_mode == "untouched":
        return untouched_acquisition_path_for_item(state, rel_path, item)
    if source_mode == "original":
        context = crop_context_for_rel(state, rel_path)
        path = context.get("raw_source") or context.get("enhanced_source") or find_raw_m_image(state, rel_path)
        if path and os.path.isfile(path):
            return path
        raise FileNotFoundError(f"No original source image found for manual recrop: {rel_path}")
    path, _ = selected_crop_path_for_item(state, rel_path, item)
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(f"No crop image found for manual recrop: {rel_path}")
    return path


def png_bytes_for_manual_crop_image(state, rel_path, source_mode="crop", channel_name="merged"):
    path = manual_crop_image_path(state, rel_path, source_mode=source_mode)
    image, _ = read_yxc(path)
    raw_preview = direct_raw_preview_rgb(image)
    channel_name = str(channel_name or "merged").strip().lower()
    if channel_name in CHANNELS:
        gray = np.asarray(raw_preview[..., CHANNELS[channel_name]], dtype=np.uint8)
        rgb = np.stack([gray, gray, gray], axis=-1)
    else:
        rgb = raw_preview
    buf = io.BytesIO()
    Image.fromarray(rgb).save(buf, format="PNG")
    return buf.getvalue()


def sanitize_polygon_points(points, image_width, image_height, label_text):
    if not isinstance(points, list) or len(points) < 3:
        raise ValueError(f"{label_text} needs at least three drawn points.")
    cleaned = []
    for point in points:
        if not isinstance(point, (list, tuple)) or len(point) < 2:
            continue
        try:
            x = float(point[0])
            y = float(point[1])
        except (TypeError, ValueError):
            continue
        if not np.isfinite(x) or not np.isfinite(y):
            continue
        cleaned.append([
            float(np.clip(x, 0, max(0, image_width - 1))),
            float(np.clip(y, 0, max(0, image_height - 1))),
        ])
    if len(cleaned) < 3:
        raise ValueError(f"{label_text} needs at least three valid drawn points.")
    return cleaned


def polygon_points_mask(image_shape, points):
    image_height, image_width = image_shape[:2]
    mask_image = Image.new("1", (image_width, image_height), 0)
    ImageDraw.Draw(mask_image).polygon([(float(x), float(y)) for x, y in points], fill=1)
    return np.asarray(mask_image, dtype=bool)


def fixed_cell_manual_annotation_preview_png(
    raw_yxc, roi_mask, focus_masks, output_path, background_mask=None, condensate_masks=None
):
    """Replace a stale analysis overlay with the newly saved manual drawing."""
    rgb = display_rgb(raw_yxc)
    overlay = np.array(rgb, copy=True)
    focus_union = (
        np.logical_or.reduce([np.asarray(mask, dtype=bool) for mask in focus_masks])
        if focus_masks
        else np.zeros(roi_mask.shape, dtype=bool)
    )
    for mask, color in (
        (roi_mask, np.array([0, 220, 255], dtype=np.uint8)),
        (
            np.logical_or.reduce(condensate_masks) if condensate_masks else None,
            np.array([0, 255, 80], dtype=np.uint8),
        ),
        (focus_union, np.array([255, 0, 255], dtype=np.uint8)),
        (background_mask, np.array([255, 140, 0], dtype=np.uint8)),
    ):
        if mask is None:
            continue
        mask = np.asarray(mask, dtype=bool)
        if np.any(mask):
            boundary = mask & ~ndi.binary_erosion(mask)
            overlay[boundary] = color

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    image = Image.fromarray(overlay, mode="RGB")
    banner_height = 58
    canvas = Image.new("RGB", (image.width, image.height + banner_height), (12, 12, 12))
    canvas.paste(image, (0, banner_height))
    draw = ImageDraw.Draw(canvas)
    draw.text((7, 5), "MANUAL ANNOTATION SAVED", fill=(255, 230, 0))
    draw.text((7, 22), "Run fixed-cell analysis for complete QC", fill=(245, 245, 245))
    draw.rectangle((7, 42, 19, 54), fill=(0, 220, 255))
    draw.text((23, 40), "ROI", fill=(245, 245, 245))
    draw.rectangle((62, 42, 74, 54), fill=(255, 0, 255))
    draw.text((78, 40), "MCP/MS2 focus", fill=(245, 245, 245))
    draw.rectangle((174, 42, 186, 54), fill=(0, 255, 80))
    draw.text((190, 40), "condensates", fill=(245, 245, 245))
    draw.rectangle((278, 42, 290, 54), fill=(255, 140, 0))
    draw.text((294, 40), "background", fill=(245, 245, 245))
    canvas.save(output_path, format="PNG")


def save_fixed_cell_manual_annotation(
    state,
    rel_path,
    roi_points,
    focus_polygons,
    source_mode="crop",
    background_polygon=None,
    condensate_polygons=None,
):
    item = state.get("items", {}).get(rel_path)
    if not isinstance(item, dict):
        raise ValueError(f"Crop is not loaded: {rel_path}")
    source_mode = str(source_mode or "crop").strip().lower()
    if source_mode not in ("crop", "original", "untouched"):
        raise ValueError("Choose the reviewed crop, original processed image, or matching untouched acquisition.")
    source_path = manual_crop_image_path(state, rel_path, source_mode=source_mode)
    if not source_path or not os.path.isfile(source_path):
        raise ValueError("The selected reviewed crop could not be found.")
    image, _ = read_yxc(source_path)
    image_height, image_width = image.shape[:2]
    roi_points = sanitize_polygon_points(roi_points, image_width, image_height, "Measurement ROI")
    if not isinstance(focus_polygons, list) or not focus_polygons:
        raise ValueError("Draw at least one MCP/MS2 focus before saving.")
    focus_polygons = [
        sanitize_polygon_points(points, image_width, image_height, f"MCP/MS2 focus {index}")
        for index, points in enumerate(focus_polygons, start=1)
    ]
    crop_box = None
    if source_mode in ("original", "untouched"):
        crop_box = reviewed_crop_bounds_on_full_image(state, rel_path, item, image.shape)
        require_polygons_inside_reviewed_crop([roi_points], crop_box, "Measurement ROI")
        require_polygons_inside_reviewed_crop(focus_polygons, crop_box, "MCP/MS2 focus")
    roi_mask = polygon_points_mask(image.shape, roi_points)
    if np.count_nonzero(roi_mask) < 20:
        raise ValueError("The measurement ROI is too small. Draw a larger ROI.")
    focus_masks = []
    for index, points in enumerate(focus_polygons, start=1):
        focus_mask = polygon_points_mask(image.shape, points)
        if not np.any(focus_mask & roi_mask):
            raise ValueError(f"MCP/MS2 focus {index} does not overlap the measurement ROI.")
        focus_masks.append(focus_mask & roi_mask)
    cleaned_condensate_polygons = []
    condensate_masks = []
    if isinstance(condensate_polygons, list):
        for index, points in enumerate(condensate_polygons, start=1):
            cleaned = sanitize_polygon_points(
                points, image_width, image_height, f"Condensate ROI {index}"
            )
            condensate_mask = polygon_points_mask(image.shape, cleaned)
            coords = np.argwhere(condensate_mask)
            if not coords.size:
                raise ValueError(f"Condensate ROI {index} is empty.")
            centroid_y, centroid_x = np.mean(coords, axis=0)
            centroid_row = int(np.clip(round(float(centroid_y)), 0, image_height - 1))
            centroid_column = int(np.clip(round(float(centroid_x)), 0, image_width - 1))
            if not roi_mask[centroid_row, centroid_column]:
                raise ValueError(
                    f"Condensate ROI {index} has its centroid outside the measurement ROI. "
                    "Redraw it inside the cyan ROI."
                )
            cleaned_condensate_polygons.append(cleaned)
            condensate_masks.append(condensate_mask)
    if crop_box:
        require_polygons_inside_reviewed_crop(cleaned_condensate_polygons, crop_box, "Condensate ROI")
    cleaned_background_polygon = []
    background_mask = None
    if isinstance(background_polygon, list) and background_polygon:
        cleaned_background_polygon = sanitize_polygon_points(
            background_polygon, image_width, image_height, "MS2 background ROI"
        )
        all_focus_mask = np.logical_or.reduce(focus_masks)
        background_mask = polygon_points_mask(image.shape, cleaned_background_polygon) & roi_mask & ~all_focus_mask
        if np.count_nonzero(background_mask) < 10:
            raise ValueError(
                "The MS2 background ROI needs at least 10 usable pixels inside the measurement ROI and outside all MS2 foci."
            )
        if crop_box:
            require_polygons_inside_reviewed_crop([cleaned_background_polygon], crop_box, "MS2 background ROI")
    annotation = {
        "version": 4,
        "source_mode": source_mode,
        "source_path": os.path.abspath(source_path),
        "image_width": int(image_width),
        "image_height": int(image_height),
        "roi_points": roi_points,
        "focus_polygons": focus_polygons,
        "condensate_polygons": cleaned_condensate_polygons,
        "background_polygon": cleaned_background_polygon,
        "reviewed_crop_box": crop_box or {},
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    item["fixed_cell_manual_annotation"] = annotation
    output_root = state.get("output_root", "")
    fixed_root = os.path.join(output_root, "fixed_cell_ms2_condensate_analysis")
    preview_root = os.path.join(fixed_root, "previews")
    safe_base = safe_name(str(item.get("crop_choice", "raw")) + "__" + rel_path.replace(os.sep, "__"))
    preview_path = os.path.join(preview_root, f"{safe_base}_fixed_cell_overlay.png")
    fixed_cell_manual_annotation_preview_png(
        image,
        roi_mask,
        focus_masks,
        preview_path,
        background_mask=background_mask,
        condensate_masks=condensate_masks,
    )
    item["fixed_cell_annotation_requires_analysis"] = True
    item["fixed_cell_manual_annotation_preview_path"] = preview_path
    item["status"] = "keep"
    item["explicit_keep"] = True
    state["last_fixed_cell_ms2_output_root"] = fixed_root
    state["last_fixed_cell_ms2_preview_root"] = preview_root
    state["last_workflow_action"] = "save_fixed_cell_manual_annotation"
    state["last_workflow_note"] = (
        f"Manual measurement ROI, MCP/MS2 focus, and optional background annotations were saved on the {source_mode} image source."
    )
    save_state(state)
    return annotation


def fixed_cell_manual_annotation_masks(item, source_path, image_shape):
    annotation = item.get("fixed_cell_manual_annotation") if isinstance(item, dict) else None
    if not isinstance(annotation, dict):
        return None
    image_height, image_width = image_shape[:2]
    if int(annotation.get("image_width", -1)) != int(image_width) or int(annotation.get("image_height", -1)) != int(image_height):
        raise ValueError("The saved manual ROI/focus annotation belongs to a different crop size. Redraw it for this crop.")
    saved_source = os.path.abspath(str(annotation.get("source_path", "") or ""))
    if saved_source and saved_source != os.path.abspath(source_path):
        raise ValueError("The selected crop changed after manual annotation. Redraw the ROI and MCP/MS2 focus for this crop.")
    roi_points = sanitize_polygon_points(annotation.get("roi_points"), image_width, image_height, "Saved measurement ROI")
    focus_rows = annotation.get("focus_polygons")
    if not isinstance(focus_rows, list) or not focus_rows:
        raise ValueError("The saved manual annotation has no MCP/MS2 focus. Redraw and save it.")
    roi_mask = polygon_points_mask(image_shape, roi_points)
    focus_masks = []
    for index, points in enumerate(focus_rows, start=1):
        cleaned = sanitize_polygon_points(points, image_width, image_height, f"Saved MCP/MS2 focus {index}")
        clipped = polygon_points_mask(image_shape, cleaned) & roi_mask
        if not np.any(clipped):
            raise ValueError(f"Saved MCP/MS2 focus {index} does not overlap the measurement ROI.")
        focus_masks.append(clipped)
    background_mask = None
    background_points = annotation.get("background_polygon")
    if isinstance(background_points, list) and background_points:
        cleaned_background = sanitize_polygon_points(
            background_points, image_width, image_height, "Saved MS2 background ROI"
        )
        all_focus_mask = np.logical_or.reduce(focus_masks)
        background_mask = polygon_points_mask(image_shape, cleaned_background) & roi_mask & ~all_focus_mask
        if np.count_nonzero(background_mask) < 10:
            raise ValueError("The saved MS2 background ROI has fewer than 10 usable pixels. Redraw it.")
    return roi_mask, focus_masks, background_mask, annotation


def fixed_cell_manual_condensate_particles(annotation, image_shape, measurement_roi_mask, pixel_um):
    """Convert user-traced condensate polygons into intact fixed-cell particle records."""
    polygons = annotation.get("condensate_polygons") if isinstance(annotation, dict) else None
    if not isinstance(polygons, list) or not polygons:
        raise ValueError(
            "Manual condensate mode is selected, but no condensate ROIs were saved. "
            "Open Draw analysis ROI + MCP focus, trace each condensate in green, and save again."
        )
    image_height, image_width = image_shape[:2]
    roi_mask = np.asarray(measurement_roi_mask, dtype=bool)
    particles = []
    for index, points in enumerate(polygons, start=1):
        cleaned = sanitize_polygon_points(
            points, image_width, image_height, f"Saved condensate ROI {index}"
        )
        particle_mask = polygon_points_mask(image_shape, cleaned)
        coords = np.argwhere(particle_mask)
        if not coords.size:
            raise ValueError(f"Saved condensate ROI {index} is empty. Redraw it.")
        centroid_y, centroid_x = [float(value) for value in np.mean(coords, axis=0)]
        centroid_row = int(np.clip(round(centroid_y), 0, image_height - 1))
        centroid_column = int(np.clip(round(centroid_x), 0, image_width - 1))
        if not roi_mask[centroid_row, centroid_column]:
            raise ValueError(
                f"Saved condensate ROI {index} has its centroid outside the measurement ROI. Redraw it."
            )
        area_px2 = float(np.count_nonzero(particle_mask))
        particles.append({
            "particle": index,
            "mask": particle_mask,
            "area_px2": area_px2,
            "area_um2": area_px2 * (pixel_um ** 2),
            "equivalent_diameter_um": 2.0 * np.sqrt(area_px2 / np.pi) * pixel_um,
            "fiji_centroid_x": np.nan,
            "fiji_centroid_y": np.nan,
            "mask_centroid_x": centroid_x,
            "mask_centroid_y": centroid_y,
            "centroid_inside_roi": True,
        })
    return particles


def clear_fixed_cell_manual_annotation(state, rel_path):
    item = state.get("items", {}).get(rel_path)
    if not isinstance(item, dict):
        raise ValueError(f"Crop is not loaded: {rel_path}")
    item.pop("fixed_cell_manual_annotation", None)
    save_state(state)


def partition_annotation_preview_png(
    raw_yxc, nucleus_mask, condensate_mask, output_path,
    background_mask=None, exclusion_mask=None, no_condensate=False,
):
    """Draw a compact audit overlay for the independent partition workflow."""
    overlay = np.array(display_rgb(raw_yxc), copy=True)
    for mask, color in (
        (nucleus_mask, np.array([0, 220, 255], dtype=np.uint8)),
        (condensate_mask, np.array([0, 255, 80], dtype=np.uint8)),
        (exclusion_mask, np.array([255, 225, 0], dtype=np.uint8)),
        (background_mask, np.array([255, 140, 0], dtype=np.uint8)),
    ):
        if mask is None:
            continue
        mask = np.asarray(mask, dtype=bool)
        if np.any(mask):
            overlay[mask & ~ndi.binary_erosion(mask)] = color
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    image = Image.fromarray(overlay, mode="RGB")
    banner_height = 58
    canvas = Image.new("RGB", (image.width, image.height + banner_height), (12, 12, 12))
    canvas.paste(image, (0, banner_height))
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((7, 10, 19, 22), fill=(0, 220, 255))
    draw.text((24, 8), "measurement/reference ROI", fill=(245, 245, 245))
    draw.rectangle((180, 10, 192, 22), fill=(0, 255, 80))
    draw.text((197, 8), "condensates", fill=(245, 245, 245))
    draw.rectangle((300, 10, 312, 22), fill=(255, 225, 0))
    draw.text((317, 8), "excluded edge", fill=(245, 245, 245))
    draw.rectangle((430, 10, 442, 22), fill=(255, 140, 0))
    draw.text((447, 8), "cell-free background", fill=(245, 245, 245))
    if no_condensate:
        draw.text((8, 32), "NO CONDENSATE explicitly recorded", fill=(255, 225, 0))
    canvas.save(output_path, format="PNG")


def save_partition_manual_annotation(
    state, rel_path, roi_points, condensate_polygons, source_mode="original",
    background_polygons=None, no_condensate=False,
):
    """Save a user-defined measurement/reference ROI and condensate polygons."""
    item = state.get("items", {}).get(rel_path)
    if not isinstance(item, dict):
        raise ValueError(f"Crop is not loaded: {rel_path}")
    source_mode = str(source_mode or "original").strip().lower()
    if source_mode not in ("crop", "original", "untouched"):
        raise ValueError("Choose the reviewed crop, original processed image, or matching untouched acquisition.")
    source_path = manual_crop_image_path(state, rel_path, source_mode=source_mode)
    image, _ = read_yxc(source_path)
    image_height, image_width = image.shape[:2]
    roi_points = sanitize_polygon_points(roi_points, image_width, image_height, "Measurement/reference ROI")
    crop_box = None
    if source_mode in ("original", "untouched"):
        crop_box = reviewed_crop_bounds_on_full_image(state, rel_path, item, image.shape)
        require_polygons_inside_reviewed_crop([roi_points], crop_box, "Measurement/reference ROI")
    nucleus_mask = polygon_points_mask(image.shape, roi_points)
    if np.count_nonzero(nucleus_mask) < 20:
        raise ValueError("The measurement/reference ROI is too small. Draw a larger ROI.")
    no_condensate = bool(no_condensate)
    if (not isinstance(condensate_polygons, list) or not condensate_polygons) and not no_condensate:
        raise ValueError("Draw at least one condensate ROI before saving.")
    cleaned_polygons = []
    condensate_masks = []
    for index, points in enumerate(condensate_polygons or [], start=1):
        cleaned = sanitize_polygon_points(points, image_width, image_height, f"Condensate ROI {index}")
        mask = polygon_points_mask(image.shape, cleaned) & nucleus_mask
        if np.count_nonzero(mask) < 1:
            raise ValueError(f"Condensate ROI {index} does not overlap the measurement/reference ROI.")
        cleaned_polygons.append(cleaned)
        condensate_masks.append(mask)
    if crop_box:
        require_polygons_inside_reviewed_crop(cleaned_polygons, crop_box, "Condensate ROI")
    union_mask = np.logical_or.reduce(condensate_masks) if condensate_masks else np.zeros(image.shape[:2], dtype=bool)
    outside_mask = nucleus_mask & ~union_mask
    if np.count_nonzero(outside_mask) < 10:
        raise ValueError("The measurement ROI minus condensates region needs at least 10 pixels.")
    cleaned_backgrounds = []
    background_masks = []
    for index, points in enumerate(background_polygons or [], start=1):
        cleaned = sanitize_polygon_points(points, image_width, image_height, f"Cell-free background ROI {index}")
        mask = polygon_points_mask(image.shape, cleaned)
        if np.count_nonzero(mask) < 5:
            raise ValueError(f"Cell-free background ROI {index} is too small.")
        if np.any(mask & nucleus_mask):
            raise ValueError(f"Cell-free background ROI {index} overlaps the measurement/reference ROI. Draw it in a cell-free area.")
        cleaned_backgrounds.append(cleaned)
        background_masks.append(mask)
    annotation = {
        "version": 2,
        "source_mode": source_mode,
        "source_path": os.path.abspath(source_path),
        "image_width": int(image_width),
        "image_height": int(image_height),
        "roi_points": roi_points,
        "condensate_polygons": cleaned_polygons,
        "background_polygons": cleaned_backgrounds,
        "no_condensate": no_condensate,
        "reviewed_crop_box": crop_box or {},
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    item["partition_manual_annotation"] = annotation
    output_root = state.get("output_root", "")
    partition_root = os.path.join(output_root, "manual_nuclear_partition_coefficient_analysis")
    preview_root = os.path.join(partition_root, "previews")
    safe_base = safe_name(str(item.get("crop_choice", "raw")) + "__" + rel_path.replace(os.sep, "__"))
    preview_path = os.path.join(preview_root, f"{safe_base}_partition_annotation.png")
    background_union = np.logical_or.reduce(background_masks) if background_masks else np.zeros(image.shape[:2], dtype=bool)
    partition_annotation_preview_png(
        image, nucleus_mask, union_mask, preview_path,
        background_mask=background_union, no_condensate=no_condensate,
    )
    item["partition_annotation_preview_path"] = preview_path
    item["status"] = "keep"
    item["explicit_keep"] = True
    state["last_partition_output_root"] = partition_root
    state["last_partition_preview_root"] = preview_root
    state["last_workflow_action"] = "save_partition_manual_annotation"
    state["last_workflow_note"] = "Manual measurement/reference, condensate, and cell-free background ROIs were saved for the independent partition-coefficient workflow."
    save_state(state)
    return annotation


def partition_annotation_masks(item, source_path, image_shape):
    annotation = item.get("partition_manual_annotation") if isinstance(item, dict) else None
    if not isinstance(annotation, dict):
        raise ValueError("No manual partition annotation is saved for this cell.")
    image_height, image_width = image_shape[:2]
    if int(annotation.get("image_width", -1)) != image_width or int(annotation.get("image_height", -1)) != image_height:
        raise ValueError("The saved partition annotation belongs to a different image size. Redraw it.")
    saved_source = os.path.abspath(str(annotation.get("source_path", "") or ""))
    if saved_source and saved_source != os.path.abspath(source_path):
        raise ValueError("The image source changed after partition annotation. Redraw it.")
    nucleus_points = sanitize_polygon_points(annotation.get("roi_points"), image_width, image_height, "Saved measurement/reference ROI")
    nucleus_mask = polygon_points_mask(image_shape, nucleus_points)
    polygons = annotation.get("condensate_polygons") or []
    no_condensate = bool(annotation.get("no_condensate", False))
    if (not isinstance(polygons, list) or not polygons) and not no_condensate:
        raise ValueError("The saved partition annotation has no condensate ROIs.")
    condensate_masks = []
    for index, points in enumerate(polygons, start=1):
        cleaned = sanitize_polygon_points(points, image_width, image_height, f"Saved condensate ROI {index}")
        clipped = polygon_points_mask(image_shape, cleaned) & nucleus_mask
        if not np.any(clipped):
            raise ValueError(f"Saved condensate ROI {index} does not overlap the measurement/reference ROI.")
        condensate_masks.append(clipped)
    background_masks = []
    for index, points in enumerate(annotation.get("background_polygons") or [], start=1):
        cleaned = sanitize_polygon_points(points, image_width, image_height, f"Saved cell-free background ROI {index}")
        mask = polygon_points_mask(image_shape, cleaned)
        if np.any(mask):
            background_masks.append(mask)
    return nucleus_mask, condensate_masks, background_masks, no_condensate, annotation


def manual_recrop_item(state, rel_path, x, y, width, height, source_mode="crop"):
    if rel_path not in state.get("items", {}):
        raise ValueError(f"Crop is not loaded: {rel_path}")

    source_path = manual_crop_image_path(state, rel_path, source_mode=source_mode)

    image, _ = read_yxc(source_path)
    image_h, image_w = image.shape[:2]
    x0 = max(0, min(int(round(float(x))), image_w - 1))
    y0 = max(0, min(int(round(float(y))), image_h - 1))
    x1 = max(x0 + 1, min(int(round(float(x) + float(width))), image_w))
    y1 = max(y0 + 1, min(int(round(float(y) + float(height))), image_h))
    if x1 - x0 < 5 or y1 - y0 < 5:
        raise ValueError("Manual crop is too small. Draw a larger rectangle around the cell.")

    crop = np.array(image[y0:y1, x0:x1, :], copy=True)
    manual_dir = os.path.join(state.get("crop_root", ""), "manual_corrected_crops")
    os.makedirs(manual_dir, exist_ok=True)
    base = safe_name(os.path.splitext(os.path.basename(rel_path))[0])
    manual_name = f"{base}_MANUAL_RECROP_{datetime.now().strftime('%H%M%S')}.tiff"
    manual_path = os.path.join(manual_dir, manual_name)
    tiff.imwrite(manual_path, crop, photometric="rgb")

    item = state["items"][rel_path]
    item.update({
        "rel_path": rel_path,
        "path": manual_path,
        "manual_crop_path": manual_path,
        "manual_original_path": source_path,
        "manual_crop_origin_x": x0,
        "manual_crop_origin_y": y0,
        "manual_crop_width": x1 - x0,
        "manual_crop_height": y1 - y0,
        "manual_crop_source_mode": source_mode,
        "status": "keep",
        "crop_choice": "manual_recrop",
        "note": f"Manual recrop saved from {source_mode} on this review item: x={x0}, y={y0}, w={x1 - x0}, h={y1 - y0}",
    })
    item.pop("fixed_cell_manual_annotation", None)
    state["last_scan"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    clear_auto_cleaned_item_fields(item)
    state["last_workflow_action"] = "manual_recrop_review"
    state["last_workflow_note"] = "Manual recrop was saved. Autoclean was not run; use Autoclean only this crop to update just this ROI."
    save_state(state)
    return rel_path, manual_path


def undo_manual_recrop(state, rel_path):
    if rel_path not in state.get("items", {}):
        raise ValueError(f"Crop is not loaded: {rel_path}")
    item = state["items"][rel_path]
    if not item.get("manual_crop_path"):
        raise ValueError("This crop does not have a manual recrop to undo.")
    original_path = item.get("manual_original_path") or os.path.join(state.get("crop_root", ""), rel_path)
    item["path"] = original_path
    item["status"] = "keep"
    item["crop_choice"] = "raw"
    item["note"] = "Manual recrop undone; restored original crop."
    item.pop("manual_crop_path", None)
    item.pop("manual_original_path", None)
    item.pop("fixed_cell_manual_annotation", None)
    clear_auto_cleaned_item_fields(item)
    state["last_scan"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state["last_workflow_action"] = "manual_recrop_undo_review"
    state["last_workflow_note"] = "Manual recrop was undone. Autoclean was not run; use Autoclean only this crop to update just this ROI."
    save_state(state)
    return rel_path, original_path


def save_display_png_for_tiff(tiff_path, png_path, max_size=900):
    image, _ = read_yxc(tiff_path)
    rgb = display_rgb(image)
    save_display_png_for_rgb(rgb, png_path, max_size=max_size)


def save_display_png_for_rgb(rgb, png_path, max_size=900):
    pil = Image.fromarray(rgb)
    pil.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
    os.makedirs(os.path.dirname(png_path), exist_ok=True)
    pil.save(png_path, format="PNG")


def valid_channel(value, fallback):
    value = str(value or "").strip().lower()
    return value if value in CHANNELS else fallback


def use_channel_cleaned_input(state):
    return normalize_analysis_input_mode(state.get("analysis_input_mode", "raw")) == "auto_cleaned"


def analysis_input_mode_for(state, analysis_key):
    key = f"{analysis_key}_analysis_input_mode"
    fallback = normalize_analysis_input_mode(state.get("analysis_input_mode", "raw"))
    return normalize_analysis_input_mode(state.get(key, fallback))


def use_auto_cleaned_input_for(state, analysis_key):
    return analysis_input_mode_for(state, analysis_key) == "auto_cleaned"


def analysis_input_label(mode):
    return "Auto-cleaned crop" if normalize_analysis_input_mode(mode) == "auto_cleaned" else "Original reviewed crop"


def clamp_int(value, fallback, minimum, maximum):
    try:
        number = int(float(value))
    except (TypeError, ValueError):
        number = fallback
    return max(minimum, min(maximum, number))


def cleaning_params_from_state(state):
    params = {
        "radius_px": clamp_int(state.get("cleaning_radius_px"), 14, 4, 40),
        "sensitivity": clamp_int(state.get("cleaning_sensitivity"), 50, 1, 100),
        "min_area_px": clamp_int(state.get("cleaning_min_area_px"), 3, 1, 100),
        "background_removal": clamp_int(state.get("cleaning_background_removal"), 0, 0, 100),
        "amount": clamp_int(state.get("cleaning_amount"), 100, 0, 100),
    }
    params["channel_params"] = {
        channel_name: {
            **params,
            "amount": clamp_int(state.get(f"{channel_name}_cleaning_amount"), params["amount"], 0, 100),
        }
        for channel_name in CHANNELS
    }
    return params


def cleaning_params_from_mapping(values, fallback):
    values = values or {}
    return {
        "radius_px": clamp_int(values.get("radius_px"), fallback["radius_px"], 4, 40),
        "sensitivity": clamp_int(values.get("sensitivity"), fallback["sensitivity"], 1, 100),
        "min_area_px": clamp_int(values.get("min_area_px"), fallback["min_area_px"], 1, 100),
        "background_removal": clamp_int(values.get("background_removal"), fallback.get("background_removal", 0), 0, 100),
        "amount": clamp_int(values.get("amount"), fallback["amount"], 0, 100),
    }


def channel_cleaning_params_from_mapping(values, fallback):
    params = cleaning_params_from_mapping(values, fallback)
    channel_params = {}
    stored = values.get("channel_params") if isinstance(values, dict) else {}
    if not isinstance(stored, dict):
        stored = {}
    fallback_stored = fallback.get("channel_params") if isinstance(fallback, dict) else {}
    if not isinstance(fallback_stored, dict):
        fallback_stored = {}
    for channel_name in CHANNELS:
        channel_fallback = cleaning_params_from_mapping(fallback_stored.get(channel_name, {}), params)
        channel_params[channel_name] = cleaning_params_from_mapping(stored.get(channel_name, values), channel_fallback)
    return channel_params


def cleaning_params_with_channels_from_mapping(values, fallback):
    params = cleaning_params_from_mapping(values, fallback)
    params["channel_params"] = channel_cleaning_params_from_mapping(values, params)
    return params


def posted_channel_cleaning_params(data, app_state):
    fallback = cleaning_params_from_state(app_state)
    channel_params = {}
    posted_roles = {}
    for channel_name in CHANNELS:
        channel_params[channel_name] = cleaning_params_from_mapping({
            "amount": data.get(f"{channel_name}_amount"),
            "radius_px": data.get(f"{channel_name}_radius_px"),
            "sensitivity": data.get(f"{channel_name}_sensitivity"),
            "min_area_px": data.get(f"{channel_name}_min_area_px"),
            "background_removal": data.get(f"{channel_name}_background_removal"),
        }, fallback)
        posted_roles[channel_name] = valid_channel_role(data.get(f"{channel_name}_role"), channel_roles_from_state(app_state).get(channel_name, "ignore"))
    posted = {
        **fallback,
        "channel_params": channel_params,
        "channels": [
            channel_name
            for channel_name in CHANNELS
            if posted_roles.get(channel_name) == "condensate"
        ],
    }
    for channel_name in CHANNELS:
        posted[f"{channel_name}_role"] = posted_roles[channel_name]
    return posted


def flatten_cleaning_params_for_ui(values, app_state):
    fallback = cleaning_params_from_state(app_state)
    params = cleaning_params_from_mapping(values, fallback)
    channel_params = channel_cleaning_params_from_mapping(values, params)
    flattened = dict(params)
    for channel_name, channel_values in channel_params.items():
        for key, value in channel_values.items():
            flattened[f"{channel_name}_{key}"] = value
    return flattened


def channel_set_from_mapping(values, fallback_channels):
    values = values or {}
    channels = values.get("channels")
    if isinstance(channels, str):
        channels = [part.strip() for part in channels.split(",")]
    if not isinstance(channels, (list, tuple, set)):
        return set(fallback_channels)
    cleaned = {valid_channel(channel_name, "") for channel_name in channels}
    cleaned = {channel_name for channel_name in cleaned if channel_name}
    # An explicit empty list is meaningful: no channel is assigned the
    # condensate role, so every channel must remain raw.  Do not replace it
    # with the previous/global condensate channels.
    return cleaned


def cleaning_params_for_rel(state, rel):
    global_params = cleaning_params_from_state(state)
    per_crop = state.get("per_crop_cleaning_params", {})
    if not isinstance(per_crop, dict):
        per_crop = {}
    return cleaning_params_from_mapping(per_crop.get(rel), global_params)


def channel_cleaning_params_for_rel(state, rel):
    global_params = cleaning_params_from_state(state)
    per_crop = state.get("per_crop_cleaning_params", {})
    if not isinstance(per_crop, dict):
        per_crop = {}
    return channel_cleaning_params_from_mapping(per_crop.get(rel) or {}, global_params)


def cleaning_channels_for_rel(state, rel):
    item = state.get("items", {}).get(rel, {}) if isinstance(state.get("items", {}), dict) else {}
    # Channel roles are the UI's source of truth.  A stale per-crop
    # "channels" list must never override a newly selected Marker/Nucleus/
    # Ignore role.
    return set(crop_condensate_channels(state, item)) if item else condensate_channels_from_state(state)


def has_per_crop_cleaning_params(state, rel):
    per_crop = state.get("per_crop_cleaning_params", {})
    return isinstance(per_crop, dict) and rel in per_crop


def cleaning_preview_base_name(rel, crop_choice):
    return bounded_output_stem("preview__" + crop_choice + "__" + rel.replace(os.sep, "__"))


def cleaning_preview_relpath(rel, crop_choice):
    return cleaning_preview_base_name(rel, crop_choice) + "_raw_vs_channel_cleaned.png"


def cleaning_preview_candidate_relpaths(rel, crop_choice):
    rel_token = rel.replace(os.sep, "__")
    choices = [crop_choice]
    if crop_choice != "manual_recrop":
        choices.append("manual_recrop")
    if crop_choice != "raw":
        choices.append("raw")
    candidates = [cleaning_preview_relpath(rel, crop_choice)]
    for choice in choices:
        candidates.append(auto_clean_base_name(rel, choice) + "_raw_vs_channel_cleaned.png")
        candidates.append(safe_name(choice + "__" + rel_token) + "_raw_vs_channel_cleaned.png")
    deduped = []
    seen = set()
    for candidate in candidates:
        if candidate in seen:
            continue
        seen.add(candidate)
        deduped.append(candidate)
    return deduped


def clean_single_channel(channel, radius_px=10):
    original = np.asarray(channel)
    arr = original.astype(np.float32)
    if arr.size == 0 or not np.any(arr > 0):
        return np.zeros(original.shape, dtype=original.dtype)
    radius_px = max(3, int(radius_px))
    try:
        corrected = white_tophat(arr, disk(radius_px))
    except Exception:
        background = ndi.gaussian_filter(arr, sigma=max(1.0, radius_px / 2.0))
        corrected = arr - background
    corrected = np.clip(corrected, 0, None)
    positive = corrected[corrected > 0]
    if positive.size:
        floor = float(np.percentile(positive, 35))
        corrected[corrected < floor] = 0
    max_value = np.iinfo(original.dtype).max if np.issubdtype(original.dtype, np.integer) else float(np.nanmax(original) or 1.0)
    return np.clip(corrected, 0, max_value).astype(original.dtype)


def clean_puncta_channel(channel, radius_px=14, min_area_px=3, sensitivity=50, mask_strength=100, background_removal=0):
    original = np.asarray(channel)
    arr = original.astype(np.float32)
    if arr.size == 0 or not np.any(arr > 0):
        return np.zeros(original.shape, dtype=original.dtype)
    sensitivity = clamp_int(sensitivity, 50, 1, 100)
    mask_strength = clamp_int(mask_strength, 100, 0, 100)
    background_removal = clamp_int(background_removal, 0, 0, 100)
    radius_px = max(3, int(radius_px))
    try:
        tiny_spots = white_tophat(arr, disk(2))
        small_spots = white_tophat(arr, disk(max(3, radius_px // 3)))
        medium_spots = white_tophat(arr, disk(max(4, radius_px // 2)))
    except Exception:
        tiny_spots = np.zeros_like(arr)
        small_spots = np.zeros_like(arr)
        medium_spots = np.zeros_like(arr)
    diffuse_background = ndi.gaussian_filter(arr, sigma=max(3.0, radius_px * 1.8))
    local_background = ndi.gaussian_filter(arr, sigma=max(1.0, radius_px / 3.0))
    spot_signal = np.maximum.reduce([
        tiny_spots,
        small_spots,
        medium_spots,
        arr - diffuse_background,
        arr - local_background,
    ])
    spot_signal = np.clip(spot_signal, 0, None)
    if np.any(spot_signal > 0):
        spot_signal = ndi.gaussian_filter(spot_signal, sigma=0.25)
    positive = spot_signal[spot_signal > 0]
    raw_positive = arr[arr > 0]
    if positive.size == 0 or raw_positive.size == 0:
        return np.zeros(original.shape, dtype=original.dtype)

    median = float(np.median(positive))
    mad = float(np.median(np.abs(positive - median)))
    robust_sigma = max(1.4826 * mad, 1e-6)
    sensitivity_offset = (50 - sensitivity) / 50.0
    strength_offset = (mask_strength - 50) / 50.0
    background_remove_strength = 100 - background_removal
    background_offset = (background_remove_strength - 50) / 50.0
    strictness = sensitivity_offset + (strength_offset * 0.45) + (background_offset * 0.9)
    seed_percentile = max(82, min(99.8, 96 + (strictness * 3)))
    body_percentile = max(68, min(98.5, 90 + (strictness * 5)))
    seed_threshold = max(
        float(np.percentile(positive, seed_percentile)),
        median + ((2.4 + strictness) * robust_sigma),
    )
    body_threshold = max(
        float(np.percentile(positive, body_percentile)),
        median + ((1.25 + (strictness * 0.65)) * robust_sigma),
    )
    raw_floor = float(np.percentile(raw_positive, max(42, min(96, 72 + (strictness * 14)))))
    peak_distance = max(1, int(round(radius_px / 4)))
    peak_coords = peak_local_max(
        spot_signal,
        min_distance=peak_distance,
        threshold_abs=seed_threshold,
        exclude_border=False,
    )
    if peak_coords.size == 0:
        fallback_threshold = max(
            float(np.percentile(positive, max(82, min(99, 93 + (sensitivity_offset * 4))))),
            median + ((1.8 + strictness) * robust_sigma),
        )
        fallback_mask = (spot_signal >= fallback_threshold) & (arr >= raw_floor)
        peak_coords = np.argwhere(fallback_mask)
    if peak_coords.size == 0:
        return np.zeros(original.shape, dtype=original.dtype)

    body_candidate = (spot_signal >= body_threshold) & (arr >= raw_floor)
    if not np.any(body_candidate):
        return np.zeros(original.shape, dtype=original.dtype)

    markers = np.zeros(original.shape, dtype=np.int32)
    for idx, (row, col) in enumerate(peak_coords, start=1):
        if body_candidate[row, col]:
            markers[row, col] = idx
    if not np.any(markers):
        return np.zeros(original.shape, dtype=original.dtype)

    # Use the background-corrected signal only to find reliable centers and
    # assign a non-overlapping watershed territory to each center.  Boundary
    # recovery below is deliberately performed against the original channel.
    territory_candidate = spot_signal > 0
    spot_labels = watershed(-spot_signal, markers, mask=territory_candidate)
    puncta_mask = np.zeros(original.shape, dtype=bool)
    max_spot_radius = max(3.0, radius_px * 0.9)
    minimum_area = max(1, int(min_area_px))
    row_grid, col_grid = np.ogrid[: original.shape[0], : original.shape[1]]
    annulus_width = max(3.0, radius_px * 0.20)
    for region in regionprops(spot_labels, intensity_image=spot_signal):
        territory = spot_labels == region.label
        seed_positions = np.argwhere(markers == region.label)
        if seed_positions.size == 0:
            continue
        peak_row, peak_col = seed_positions[0]
        peak_value = float(spot_signal[peak_row, peak_col])
        if peak_value < seed_threshold:
            continue

        distance_map = np.sqrt(
            (row_grid - peak_row) ** 2
            + (col_grid - peak_col) ** 2
        )
        radius_limit = distance_map <= max_spot_radius

        # Retain only the strict body connected to this peak as the
        # high-confidence hysteresis seed.  Disconnected bright/noisy pixels in
        # the same watershed territory must not become independent objects.
        core_candidate = territory & body_candidate & radius_limit
        local_threshold = max(body_threshold, peak_value * max(0.10, min(0.45, 0.16 + (background_offset * 0.12) + (strength_offset * 0.04))))
        core_candidate &= spot_signal >= local_threshold
        peak_seed = np.zeros(original.shape, dtype=bool)
        peak_seed[peak_row, peak_col] = True
        core_candidate |= peak_seed
        immediate_body = ndi.binary_propagation(
            peak_seed,
            mask=core_candidate,
        )

        # Estimate the original-channel background in an annulus surrounding
        # this punctum.  The threshold adapts to local haze/noise instead of
        # applying one global edge cutoff to the whole crop.
        annulus = (
            (distance_map > max_spot_radius)
            & (distance_map <= (max_spot_radius + annulus_width))
        )
        annulus_values = arr[annulus & np.isfinite(arr)]
        if annulus_values.size:
            local_background_median = float(np.median(annulus_values))
            local_background_mad = float(
                np.median(np.abs(annulus_values - local_background_median))
            )
        else:
            local_background_median = float(np.median(raw_positive))
            local_background_mad = float(
                np.median(np.abs(raw_positive - local_background_median))
            )
        local_background_sigma = max(1.4826 * local_background_mad, 1e-6)
        noise_multiplier = float(np.clip(1.25 + (strictness * 0.35), 0.50, 2.50))
        peak_raw_value = float(arr[peak_row, peak_col])
        edge_fraction = float(np.clip(0.12 + (strictness * 0.025), 0.06, 0.22))
        recovery_floor = max(
            local_background_median + (noise_multiplier * local_background_sigma),
            peak_raw_value * edge_fraction,
        )

        # Morphological reconstruction: grow only through connected original
        # pixels that are above the local background floor, stay inside this
        # center's watershed territory, and remain within the selected radius.
        recovery_mask = territory & radius_limit & (arr >= recovery_floor)
        recovery_mask |= immediate_body
        recovered_body = ndi.binary_propagation(
            immediate_body,
            mask=recovery_mask,
        )

        # Enforce the UI's requested minimum punctum area exactly.  The prior
        # implementation unintentionally capped this control at two pixels.
        if np.count_nonzero(recovered_body) < minimum_area:
            continue
        puncta_mask |= recovered_body

    if not np.any(puncta_mask):
        return np.zeros(original.shape, dtype=original.dtype)
    cleaned = np.zeros_like(original)
    cleaned[puncta_mask] = original[puncta_mask]
    return cleaned


def scale_single_channel(channel):
    """Scale one image channel to uint8 for raw-vs-cleaned previews."""
    arr = np.asarray(channel, dtype=np.float64)
    nonzero = arr[arr > 0]
    if nonzero.size:
        low, high = np.percentile(nonzero, [1, 99.8])
    else:
        low, high = np.percentile(arr, [1, 99.8])
    if high <= low:
        return np.zeros(arr.shape, dtype=np.uint8)
    return (np.clip((arr - low) / (high - low), 0, 1) * 255).astype(np.uint8)


def save_raw_vs_cleaned_preview(raw_yxc, cleaned_yxc, preview_path, preserve_raw_channels=None, puncta_cleaned_channels=None):
    os.makedirs(os.path.dirname(preview_path), exist_ok=True)
    preserve_raw_channels = set(preserve_raw_channels or [])
    puncta_cleaned_channels = set(puncta_cleaned_channels or [])
    panels = []
    for channel_name, color_index in (("red", 0), ("green", 1), ("blue", 2)):
        raw_scaled = scale_single_channel(raw_yxc[..., CHANNELS[channel_name]])
        cleaned_scaled = scale_single_channel(cleaned_yxc[..., CHANNELS[channel_name]])
        raw_panel = np.zeros((*raw_scaled.shape, 3), dtype=np.uint8)
        cleaned_panel = np.zeros((*cleaned_scaled.shape, 3), dtype=np.uint8)
        raw_panel[..., color_index] = raw_scaled
        cleaned_panel[..., color_index] = cleaned_scaled
        panels.append((f"Raw {channel_name}", raw_panel))
        if channel_name in preserve_raw_channels:
            cleaned_label = f"Raw-preserved {channel_name}"
        elif channel_name in puncta_cleaned_channels:
            cleaned_label = f"Puncta-cleaned {channel_name}"
        else:
            cleaned_label = f"Cleaned {channel_name}"
        panels.append((cleaned_label, cleaned_panel))

    pane_w = max(320, max(panel.shape[1] for _, panel in panels))
    pane_h = max(240, max(panel.shape[0] for _, panel in panels))
    label_h = 34
    cols = 2
    rows = 3
    canvas = Image.new("RGB", (pane_w * cols, (pane_h + label_h) * rows), (12, 14, 18))
    draw = ImageDraw.Draw(canvas)
    font = graph_font(17, bold=True)
    resample = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.BICUBIC)
    for idx, (label_text, panel) in enumerate(panels):
        row = idx // cols
        col = idx % cols
        x0 = col * pane_w
        y0 = row * (pane_h + label_h)
        draw.text((x0 + 10, y0 + 8), label_text, font=font, fill=(255, 255, 255))
        image = Image.fromarray(panel, mode="RGB")
        scale = min(pane_w / image.width, pane_h / image.height)
        image = image.resize((max(1, int(image.width * scale)), max(1, int(image.height * scale))), resample)
        canvas.paste(image, (x0 + (pane_w - image.width) // 2, y0 + label_h + (pane_h - image.height) // 2))
    canvas.save(preview_path)


def prepare_analysis_input_crop(src, output_root, base_name, use_cleaned, preserve_raw_channels=None, puncta_cleaned_channels=None, cleaning_params=None):
    if not use_cleaned:
        return src, "", "", "original_raw_crop"
    cleaning_params = cleaning_params or {}
    radius_px = clamp_int(cleaning_params.get("radius_px"), 14, 4, 40)
    sensitivity = clamp_int(cleaning_params.get("sensitivity"), 50, 1, 100)
    min_area_px = clamp_int(cleaning_params.get("min_area_px"), 3, 1, 100)
    background_removal = clamp_int(cleaning_params.get("background_removal"), 100, 0, 100)
    amount = clamp_int(cleaning_params.get("amount"), 100, 0, 100)
    per_channel_params = channel_cleaning_params_from_mapping(cleaning_params, cleaning_params_from_mapping(cleaning_params, cleaning_params_from_state({})))
    preserve_raw_channels = {
        valid_channel(channel_name, "")
        for channel_name in (preserve_raw_channels or [])
        if valid_channel(channel_name, "")
    }
    puncta_cleaned_channels = {
        valid_channel(channel_name, "")
        for channel_name in (puncta_cleaned_channels or [])
        if valid_channel(channel_name, "")
    } - preserve_raw_channels
    image_yxc, layout = read_yxc(src)
    cleaned = np.zeros_like(image_yxc)
    for channel_name, index in CHANNELS.items():
        if channel_name in preserve_raw_channels:
            cleaned[..., index] = image_yxc[..., index]
        elif channel_name in puncta_cleaned_channels:
            channel_params = per_channel_params.get(channel_name, {"radius_px": radius_px, "sensitivity": sensitivity, "min_area_px": min_area_px, "background_removal": background_removal, "amount": amount})
            channel_radius = channel_params["radius_px"]
            channel_sensitivity = channel_params["sensitivity"]
            channel_min_area = channel_params["min_area_px"]
            channel_background_removal = channel_params.get("background_removal", background_removal)
            channel_amount = channel_params["amount"]
            spot_masked = clean_puncta_channel(
                image_yxc[..., index],
                radius_px=channel_radius,
                min_area_px=channel_min_area,
                sensitivity=channel_sensitivity,
                mask_strength=channel_amount,
                background_removal=channel_background_removal,
            )
            spot_mask = spot_masked > 0
            alpha = channel_amount / 100.0
            background_keep = channel_background_removal / 100.0
            raw_float = image_yxc[..., index].astype(np.float32)
            spot_float = spot_masked.astype(np.float32)
            blended = (raw_float * (1.0 - alpha) * background_keep) + (spot_float * alpha)
            blended[spot_mask] = np.maximum(blended[spot_mask], spot_float[spot_mask])
            max_value = np.iinfo(image_yxc.dtype).max if np.issubdtype(image_yxc.dtype, np.integer) else float(np.nanmax(image_yxc) or 1.0)
            cleaned[..., index] = np.clip(blended, 0, max_value).astype(image_yxc.dtype)
        else:
            cleaned[..., index] = clean_single_channel(image_yxc[..., index], radius_px=10)
    cleaned_root = os.path.join(output_root, "channel_cleaned_analysis_inputs")
    cleaned_rgb_root = os.path.join(cleaned_root, "cleaned_rgb_crops")
    preview_root = os.path.join(cleaned_root, "raw_vs_cleaned_channel_previews")
    split_root = os.path.join(cleaned_root, "cleaned_split_channels")
    os.makedirs(cleaned_rgb_root, exist_ok=True)
    os.makedirs(preview_root, exist_ok=True)
    os.makedirs(split_root, exist_ok=True)
    cleaned_path = os.path.join(cleaned_rgb_root, base_name + "_channel_cleaned_analysis_input.tif")
    tiff.imwrite(cleaned_path, cleaned, photometric="rgb")
    for channel_name, index in CHANNELS.items():
        tiff.imwrite(os.path.join(split_root, base_name + f"_cleaned_{channel_name}.tif"), cleaned[..., index])
    preview_path = os.path.join(preview_root, base_name + "_raw_vs_channel_cleaned.png")
    save_raw_vs_cleaned_preview(image_yxc, cleaned, preview_path, preserve_raw_channels, puncta_cleaned_channels)
    params_path = os.path.join(cleaned_root, "channel_cleaning_parameters.txt")
    with open(params_path, "w", encoding="utf-8") as f:
        f.write("Channel-cleaned analysis input parameters\n")
        f.write("Original raw crop is preserved. Cleaned RGB crops are derived copies only.\n")
        f.write("Merged images are split into red, green, and blue internally, cleaned separately, then recombined.\n")
        f.write("Every channel marked condensate/puncta is puncta-cleaned; all non-condensate channels are preserved raw.\n")
        f.write(f"Condensate/puncta cleaning radius_px={radius_px}\n")
        f.write(f"Condensate/puncta cleaning sensitivity={sensitivity} (higher keeps more signal; lower is stricter)\n")
        f.write("Puncta boundary recovery=local hysteresis on original intensities within watershed territories\n")
        f.write(f"Condensate/puncta cleaning min_area_px={min_area_px}\n")
        f.write(f"Condensate/puncta background_kept={background_removal} percent (0=no diffuse background kept, 100=raw background allowed)\n")
        f.write(f"Condensate/puncta cleaning amount={amount} percent (0=raw/original, 100=fully cleaned)\n")
        for channel_name in CHANNELS:
            channel_params = per_channel_params[channel_name]
            f.write(f"{channel_name} channel: amount={channel_params['amount']} background_kept={channel_params.get('background_removal', background_removal)} radius_px={channel_params['radius_px']} sensitivity={channel_params['sensitivity']} min_area_px={channel_params['min_area_px']}\n")
    if preserve_raw_channels:
        preserved_note_path = os.path.join(cleaned_root, base_name + "_preserved_raw_channels.txt")
        with open(preserved_note_path, "w", encoding="utf-8") as f:
            f.write("Raw-preserved analysis channels\n")
            f.write("These channels were copied from the reviewed crop without background cleaning:\n")
            for channel_name in sorted(preserve_raw_channels):
                f.write(f"- {channel_name}\n")
    if puncta_cleaned_channels:
        puncta_note_path = os.path.join(cleaned_root, base_name + "_puncta_cleaned_channels.txt")
        with open(puncta_note_path, "w", encoding="utf-8") as f:
            f.write("Puncta-cleaned analysis channels\n")
            f.write("These channels used stricter puncta-only cleanup to suppress diffuse background:\n")
            for channel_name in sorted(puncta_cleaned_channels):
                f.write(f"- {channel_name}\n")
    return cleaned_path, preview_path, params_path, f"channel_cleaned_from_{layout}"


def auto_clean_base_name(rel, crop_choice):
    return bounded_output_stem("auto_cleaned__" + crop_choice + "__" + rel.replace(os.sep, "__"))


def generate_auto_cleaned_crop_for_item(state, rel, item):
    ensure_sibling_output_layout(state)
    src, crop_choice = selected_crop_path_for_item(state, rel, item)
    if not src or not os.path.isfile(src):
        return None
    output_root = state.get("output_root") or APP_ROOT
    os.makedirs(output_root, exist_ok=True)
    cleaning_params = cleaning_params_for_rel(state, rel)
    channel_params = channel_cleaning_params_for_rel(state, rel)
    analysis_cleaning_params = {**cleaning_params, "channel_params": channel_params}
    puncta_cleaned_channels = cleaning_channels_for_rel(state, rel)
    raw_preserved_channels = set(CHANNELS) - puncta_cleaned_channels
    base_name = auto_clean_base_name(rel, crop_choice)
    cleaned_path, preview_path, params_path, analysis_input = prepare_analysis_input_crop(
        src,
        output_root,
        base_name,
        True,
        preserve_raw_channels=raw_preserved_channels,
        puncta_cleaned_channels=puncta_cleaned_channels,
        cleaning_params=analysis_cleaning_params,
    )
    item["auto_cleaned_crop_path"] = cleaned_path
    item["auto_cleaned_preview_path"] = preview_path
    item["auto_cleaned_params_path"] = params_path
    item["auto_cleaned_source_path"] = src
    item["auto_cleaned_source_mtime"] = os.path.getmtime(src)
    item["auto_cleaned_input_mode"] = analysis_input
    item["auto_cleaned_channels"] = sorted(puncta_cleaned_channels)
    item["auto_cleaned_settings_signature"] = auto_cleaned_settings_signature_for_item(state, rel, item, src)
    item["auto_cleaned_generated_at"] = datetime.now().isoformat(timespec="seconds")
    return cleaned_path


def autoclean_one_crop_with_saved_settings(state, rel):
    """Regenerate one kept crop without changing any shared or other-crop settings."""
    ensure_sibling_output_layout(state)
    items = state.get("items", {})
    if rel not in items:
        raise ValueError(f"Crop is not loaded: {rel}")
    item = items[rel]
    if item.get("status", "keep") != "keep":
        raise ValueError("Only a kept crop can be autocleaned.")
    cleaned_path = generate_auto_cleaned_crop_for_item(state, rel, item)
    if not cleaned_path:
        raise FileNotFoundError("The selected crop image could not be found.")
    state["last_workflow_action"] = "autoclean_one_crop"
    state["last_workflow_note"] = f"Autoclean regenerated only the selected crop: {rel}"
    state["last_single_autoclean_rel"] = rel
    state["last_single_autoclean_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Match batch AutoClean so count/size immediately uses the regenerated input.
    state["analysis_input_mode"] = "auto_cleaned"
    state["pcc_analysis_input_mode"] = "auto_cleaned"
    state["count_size_analysis_input_mode"] = "auto_cleaned"
    save_state(state)
    return cleaned_path, item.get("auto_cleaned_preview_path", "")


def clear_auto_cleaned_item_fields(item):
    for key in (
        "auto_cleaned_crop_path",
        "auto_cleaned_preview_path",
        "auto_cleaned_params_path",
        "auto_cleaned_source_path",
        "auto_cleaned_source_mtime",
        "auto_cleaned_input_mode",
        "auto_cleaned_channels",
        "auto_cleaned_settings_signature",
        "auto_cleaned_generated_at",
    ):
        item.pop(key, None)


def generate_auto_cleaned_crops_for_loaded(state, only_kept=False):
    count = 0
    errors = []
    for rel, item in sorted(state.get("items", {}).items()):
        if only_kept and item.get("status", "keep") != "keep":
            continue
        try:
            if generate_auto_cleaned_crop_for_item(state, rel, item):
                count += 1
        except Exception as error:
            errors.append({"rel_path": rel, "error": str(error)})
    state["last_auto_cleaned_count"] = count
    state["last_auto_cleaned_errors"] = errors[-50:]
    state["last_auto_cleaned_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_state(state)
    return count, errors


def current_auto_cleaned_input_for_item(item, src):
    cleaned_path = item.get("auto_cleaned_crop_path", "")
    if not cleaned_path or not is_real_tiff_file(cleaned_path):
        return None
    if os.path.abspath(item.get("auto_cleaned_source_path", "") or "") != os.path.abspath(src):
        return None
    try:
        if float(item.get("auto_cleaned_source_mtime", -1)) != os.path.getmtime(src):
            return None
    except (OSError, TypeError, ValueError):
        return None
    # A manual recrop can replace the source quickly enough that its mtime is
    # unchanged on filesystems with coarse timestamp resolution.  Never reuse
    # a cached cleaned crop whose spatial dimensions differ from its source.
    try:
        source_image, _ = read_yxc(src)
        cleaned_image, _ = read_yxc(cleaned_path)
        if source_image.shape[:2] != cleaned_image.shape[:2]:
            return None
    except Exception:
        return None
    return (
        cleaned_path,
        item.get("auto_cleaned_preview_path", ""),
        item.get("auto_cleaned_params_path", ""),
        item.get("auto_cleaned_input_mode", "channel_cleaned_from_cached_autoclean"),
    )


def auto_cleaned_settings_signature_for_item(state, rel, item, src):
    try:
        src_mtime = os.path.getmtime(src)
    except OSError:
        src_mtime = None
    payload = {
        "source_path": os.path.abspath(src) if src else "",
        "source_mtime": src_mtime,
        "channel_roles": channel_roles_for_item(state, item),
        "cleaning_params": cleaning_params_for_rel(state, rel),
        "channel_cleaning_params": channel_cleaning_params_for_rel(state, rel),
        "cleaning_channels": sorted(cleaning_channels_for_rel(state, rel)),
    }
    return hashlib.sha256(json.dumps(payload, sort_keys=True, default=str).encode("utf-8")).hexdigest()


def current_auto_cleaned_input_for_item_settings(state, rel, item, src):
    cached = current_auto_cleaned_input_for_item(item, src)
    if not cached:
        return None
    if item.get("auto_cleaned_generated_at") == "reloaded_existing_file":
        return cached
    current_signature = auto_cleaned_settings_signature_for_item(state, rel, item, src)
    if item.get("auto_cleaned_settings_signature") != current_signature:
        return None
    return cached


def reload_auto_cleaned_inputs_for_loaded(state):
    output_root = ensure_sibling_output_layout(state) or state.get("output_root") or APP_ROOT
    cleaned_root = os.path.join(output_root, "channel_cleaned_analysis_inputs")
    cleaned_rgb_root = os.path.join(cleaned_root, "cleaned_rgb_crops")
    preview_root = os.path.join(cleaned_root, "raw_vs_cleaned_channel_previews")
    params_path = os.path.join(cleaned_root, "channel_cleaning_parameters.txt")
    if not os.path.isdir(cleaned_rgb_root):
        return 0
    recovered = 0
    for rel, item in sorted(state.get("items", {}).items()):
        src, crop_choice = selected_crop_path_for_item(state, rel, item)
        if not src or not os.path.isfile(src):
            continue
        # A manual recrop can have been saved after autocleaning.  Older
        # state files may still say crop_choice="raw", while the cached
        # output was correctly written with the manual_recrop prefix.  Try
        # all known prefixes before falling back to a filename-token match.
        base_names = []
        for choice in (crop_choice, "manual_recrop", "raw"):
            candidate = auto_clean_base_name(rel, choice)
            if candidate not in base_names:
                base_names.append(candidate)
        cleaned_path = ""
        preview_path = ""
        for candidate in base_names:
            path = os.path.join(cleaned_rgb_root, candidate + "_channel_cleaned_analysis_input.tif")
            if is_real_tiff_file(path):
                cleaned_path = path
                preview_path = os.path.join(preview_root, candidate + "_raw_vs_channel_cleaned.png")
                break
        if not cleaned_path:
            rel_token = safe_name(rel.replace(os.sep, "__"))
            matches = sorted(
                name for name in os.listdir(cleaned_rgb_root)
                if not name.startswith("._")
                and rel_token in name
                and name.endswith("_channel_cleaned_analysis_input.tif")
                and is_real_tiff_file(os.path.join(cleaned_rgb_root, name))
            )
            if matches:
                cleaned_path = os.path.join(cleaned_rgb_root, matches[0])
                preview_name = matches[0].replace("_channel_cleaned_analysis_input.tif", "_raw_vs_channel_cleaned.png")
                preview_path = os.path.join(preview_root, preview_name)
        if not is_real_tiff_file(cleaned_path):
            legacy_candidates = [
                os.path.join(cleaned_rgb_root, "preview__" + safe_name(crop_choice + "__" + rel.replace(os.sep, "__")) + "_channel_cleaned_analysis_input.tif"),
                os.path.join(cleaned_rgb_root, safe_name(crop_choice + "__" + rel.replace(os.sep, "__")) + "_channel_cleaned_analysis_input.tif"),
            ]
            cleaned_path = next((path for path in legacy_candidates if is_real_tiff_file(path)), "")
        if not cleaned_path:
            continue
        if not is_real_png_file(preview_path):
            legacy_previews = [
                os.path.join(preview_root, "preview__" + safe_name(crop_choice + "__" + rel.replace(os.sep, "__")) + "_raw_vs_channel_cleaned.png"),
                os.path.join(preview_root, safe_name(crop_choice + "__" + rel.replace(os.sep, "__")) + "_raw_vs_channel_cleaned.png"),
            ]
            preview_path = next((path for path in legacy_previews if is_real_png_file(path)), "")
        item["auto_cleaned_crop_path"] = cleaned_path
        item["auto_cleaned_preview_path"] = preview_path
        item["auto_cleaned_params_path"] = params_path if os.path.isfile(params_path) else ""
        item["auto_cleaned_source_path"] = src
        item["auto_cleaned_source_mtime"] = os.path.getmtime(src)
        item["auto_cleaned_input_mode"] = "channel_cleaned_from_reloaded_autoclean"
        item["auto_cleaned_channels"] = sorted(cleaning_channels_for_rel(state, rel))
        item["auto_cleaned_settings_signature"] = auto_cleaned_settings_signature_for_item(state, rel, item, src)
        item["auto_cleaned_generated_at"] = "reloaded_existing_file"
        recovered += 1
    state["last_auto_cleaned_reloaded_count"] = recovered
    state["last_auto_cleaned_reloaded_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return recovered


def autoclean_needed_for_item(state, rel, item):
    if item.get("status", "keep") != "keep":
        return False
    try:
        src, _ = selected_crop_path_for_item(state, rel, item)
    except Exception:
        return False
    if not src or not os.path.isfile(src):
        return False
    return current_auto_cleaned_input_for_item_settings(state, rel, item, src) is None


def analysis_input_for_reviewed_crop(state, rel, item, approved_raw_path, output_root, base_name, use_cleaned):
    if not use_cleaned:
        return approved_raw_path, "", "", "original_raw_crop"
    selected_src, _ = selected_crop_path_for_item(state, rel, item)
    cached = current_auto_cleaned_input_for_item_settings(state, rel, item, selected_src)
    if cached:
        return cached
    cleaning_params = cleaning_params_for_rel(state, rel)
    channel_cleaning_params = channel_cleaning_params_for_rel(state, rel)
    analysis_cleaning_params = {**cleaning_params, "channel_params": channel_cleaning_params}
    puncta_cleaned_channels = cleaning_channels_for_rel(state, rel)
    raw_preserved_channels = set(CHANNELS) - puncta_cleaned_channels
    cleaned_path, preview_path, params_path, analysis_input = prepare_analysis_input_crop(
        approved_raw_path,
        output_root,
        base_name,
        True,
        preserve_raw_channels=raw_preserved_channels,
        puncta_cleaned_channels=puncta_cleaned_channels,
        cleaning_params=analysis_cleaning_params,
    )
    item["auto_cleaned_crop_path"] = cleaned_path
    item["auto_cleaned_preview_path"] = preview_path
    item["auto_cleaned_params_path"] = params_path
    item["auto_cleaned_source_path"] = selected_src
    item["auto_cleaned_source_mtime"] = os.path.getmtime(selected_src)
    item["auto_cleaned_input_mode"] = analysis_input
    item["auto_cleaned_channels"] = sorted(puncta_cleaned_channels)
    item["auto_cleaned_settings_signature"] = auto_cleaned_settings_signature_for_item(state, rel, item, selected_src)
    item["auto_cleaned_generated_at"] = datetime.now().isoformat(timespec="seconds")
    return cleaned_path, preview_path, params_path, analysis_input


def ensure_auto_cleaned_inputs_current_for_analysis(state, analysis_keys):
    keys = {key for key in analysis_keys if use_auto_cleaned_input_for(state, key)}
    if not keys:
        state["last_auto_cleaned_refreshed_before_analysis_count"] = 0
        state["last_auto_cleaned_refreshed_before_analysis_errors"] = []
        return 0, []
    refreshed = 0
    errors = []
    for rel, item in sorted(state.get("items", {}).items()):
        if item.get("status", "keep") != "keep":
            continue
        try:
            if autoclean_needed_for_item(state, rel, item):
                if generate_auto_cleaned_crop_for_item(state, rel, item):
                    refreshed += 1
        except Exception as error:
            errors.append({"rel_path": rel, "error": str(error)})
    state["last_auto_cleaned_refreshed_before_analysis_count"] = refreshed
    state["last_auto_cleaned_refreshed_before_analysis_errors"] = errors[-50:]
    state["last_auto_cleaned_refreshed_before_analysis_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_state(state)
    return refreshed, errors


def auto_clean_refresh_message(refreshed, errors):
    parts = []
    if refreshed:
        parts.append(f"Auto-clean refreshed {refreshed} stale/missing cleaned crop(s) before analysis.")
    if errors:
        parts.append(f"Auto-clean had {len(errors)} refresh error(s); see the app state/results for details.")
    return " ".join(parts)


def paired_pixels(channel_a, channel_b, mask=None):
    a = np.asarray(channel_a, dtype=np.float64)
    b = np.asarray(channel_b, dtype=np.float64)
    if mask is not None:
        mask = np.asarray(mask, dtype=bool)
        a = a[mask]
        b = b[mask]
    else:
        a = a.ravel()
        b = b.ravel()
    valid = np.isfinite(a) & np.isfinite(b)
    return a[valid], b[valid]


def pearson_coefficient(channel_a, channel_b, mask=None):
    a, b = paired_pixels(channel_a, channel_b, mask)
    if a.size < 2:
        return np.nan
    aa = a - a.mean()
    bb = b - b.mean()
    denom = np.sqrt(np.sum(aa * aa) * np.sum(bb * bb))
    if denom == 0:
        return np.nan
    return float(np.sum(aa * bb) / denom)


def overlap_and_k(channel_a, channel_b, mask=None):
    a, b = paired_pixels(channel_a, channel_b, mask)
    numerator = np.sum(a * b)
    sum_a2 = np.sum(a * a)
    sum_b2 = np.sum(b * b)
    denom = np.sqrt(sum_a2 * sum_b2)
    return (
        float(numerator / denom) if denom != 0 else np.nan,
        float(numerator / sum_a2) if sum_a2 != 0 else np.nan,
        float(numerator / sum_b2) if sum_b2 != 0 else np.nan,
    )


def manders_original(channel_a, channel_b, mask=None):
    a, b = paired_pixels(channel_a, channel_b, mask)
    sum_a = np.sum(a)
    sum_b = np.sum(b)
    m1 = np.sum(a[b > 0]) / sum_a if sum_a > 0 else np.nan
    m2 = np.sum(b[a > 0]) / sum_b if sum_b > 0 else np.nan
    return float(m1), float(m2)


def pearson_for_mask(a, b, mask):
    if np.count_nonzero(mask) < 2:
        return np.nan
    aa = a[mask] - a[mask].mean()
    bb = b[mask] - b[mask].mean()
    denom = np.sqrt(np.sum(aa * aa) * np.sum(bb * bb))
    if denom == 0:
        return np.nan
    return float(np.sum(aa * bb) / denom)


def imagej_ij_isodata(histogram):
    data = np.asarray(histogram, dtype=np.int64).copy()
    max_value = len(data) - 1
    count_zero = data[0]
    data[0] = 0
    count_max = data[max_value]
    data[max_value] = 0
    min_bin = 0
    while min_bin < max_value and data[min_bin] == 0:
        min_bin += 1
    max_bin = max_value
    while max_bin > 0 and data[max_bin] == 0:
        max_bin -= 1
    if min_bin >= max_bin:
        data[0] = count_zero
        data[max_value] = count_max
        return len(data) // 2
    moving_index = min_bin
    result = 0.0
    while True:
        low_bins = np.arange(min_bin, moving_index + 1)
        low_counts = data[min_bin:moving_index + 1]
        high_bins = np.arange(moving_index + 1, max_bin + 1)
        high_counts = data[moving_index + 1:max_bin + 1]
        sum2 = np.sum(low_counts)
        sum4 = np.sum(high_counts)
        if sum2 == 0 or sum4 == 0:
            break
        result = (np.sum(low_bins * low_counts) / sum2 + np.sum(high_bins * high_counts) / sum4) / 2.0
        moving_index += 1
        if not ((moving_index + 1) <= result and moving_index < max_bin - 1):
            break
    data[0] = count_zero
    data[max_value] = count_max
    return int(np.floor(result + 0.5))


def imagej_default_threshold(channel, mask=None):
    pixels = np.asarray(channel, dtype=np.float64)
    if mask is not None:
        pixels = pixels[np.asarray(mask, dtype=bool)]
    pixels = pixels[np.isfinite(pixels)].astype(np.int64)
    if pixels.size == 0:
        return np.nan
    histogram = np.bincount(np.clip(pixels, 0, 255), minlength=256)
    data = histogram.copy()
    mode = int(np.argmax(data))
    max_count = int(data[mode])
    max_count_2 = max([int(v) for i, v in enumerate(data) if i != mode] or [0])
    if max_count > (max_count_2 * 2) and max_count_2 != 0:
        data[mode] = int(max_count_2 * 1.5)
    return float(imagej_ij_isodata(data))


def manders_thresholded(channel_a, channel_b, threshold_a, threshold_b, mask=None):
    a, b = paired_pixels(channel_a, channel_b, mask)
    if not np.isfinite(threshold_a) or not np.isfinite(threshold_b):
        return np.nan, np.nan
    denom_a = np.sum(a)
    denom_b = np.sum(b)
    m1 = np.sum(a[b > threshold_b]) / denom_a if denom_a > 0 else np.nan
    m2 = np.sum(b[a > threshold_a]) / denom_b if denom_b > 0 else np.nan
    return float(m1), float(m2)


def analyze_pair(image_name, crop_file, pair_name, channel_a, channel_b, mask):
    pcc = pearson_coefficient(channel_a, channel_b, mask)
    overlap, k1, k2 = overlap_and_k(channel_a, channel_b, mask)
    m1_original, m2_original = manders_original(channel_a, channel_b, mask)
    threshold_a = imagej_default_threshold(channel_a, mask)
    threshold_b = imagej_default_threshold(channel_b, mask)
    a, b = paired_pixels(channel_a, channel_b, mask)
    below = (a <= threshold_a) & (b <= threshold_b)
    r_below = pearson_for_mask(a, b, below)
    m1_thresholded, m2_thresholded = manders_thresholded(channel_a, channel_b, threshold_a, threshold_b, mask)
    return {
        "Use_For_Summary": True,
        "Image": image_name,
        "Crop_File": crop_file,
        "Pair": pair_name,
        "Pixels_Analyzed": int(a.size),
        "PCC": pcc,
        "Overlap": overlap,
        "k1": k1,
        "k2": k2,
        "ImageJ_threshold_A": threshold_a,
        "ImageJ_threshold_B": threshold_b,
        "Pearson_below_ImageJ_threshold": r_below,
        "M1_original": m1_original,
        "M2_original": m2_original,
        "M1_thresholded": m1_thresholded,
        "M2_thresholded": m2_thresholded,
    }


def reorder_pcc_manders_columns(raw_df):
    if raw_df.empty:
        return raw_df

    preferred = [
        "Use_For_Summary",
        "Image",
        "Pair",
        "PCC_Manders_Pair_Label",
        "PCC_Manders_Pair_Type",
        "Crop_File",
        "Review_Relative_Path",
        "Analysis_Mode",
        "Analysis_Engine",
        "Pixels_Analyzed",
        "Fiji_Channel_A",
        "Fiji_Channel_A_Label",
        "Fiji_Channel_B",
        "Fiji_Channel_B_Label",
        "Channel_A_Label",
        "Channel_B_Label",
        "Channel_A_Role_For_Crop",
        "Channel_B_Role_For_Crop",
        "Condensate_Channel_For_Crop",
        "Condensate_Channel_Label_For_Crop",
        "Marker_Channel_For_Crop",
        "Marker_Channel_Label_For_Crop",
        "Condensate_Channel_A_For_Crop",
        "Condensate_Channel_A_Label_For_Crop",
        "Condensate_Channel_B_For_Crop",
        "Condensate_Channel_B_Label_For_Crop",
        "Red_Channel_Role_For_Crop",
        "Green_Channel_Role_For_Crop",
        "Blue_Channel_Role_For_Crop",
        "PCC",
        "PCC_Interpretation",
        "Overlap",
        "Overlap_Interpretation",
        "k1",
        "k1_Interpretation",
        "k2",
        "k2_Interpretation",
        "JACoP_threshold_A",
        "JACoP_threshold_B",
        "M1_original",
        "M1_original_Interpretation",
        "M1_thresholded",
        "M1_thresholded_Interpretation",
        "M1_Channel_Fraction_Measured",
        "M1_Overlap_Channel",
        "M2_original",
        "M2_original_Interpretation",
        "M2_thresholded",
        "M2_thresholded_Interpretation",
        "M2_Channel_Fraction_Measured",
        "M2_Overlap_Channel",
        "Manders_Display_Order_Note",
        "Manders_Condensate_Overlap_Marker_original",
        "Manders_Marker_Overlap_Condensate_original",
        "Manders_Condensate_Overlap_Marker_thresholded",
        "Manders_Marker_Overlap_Condensate_thresholded",
        "Manders_Channel_A_Overlap_Channel_B_original",
        "Manders_Channel_B_Overlap_Channel_A_original",
        "Manders_Channel_A_Overlap_Channel_B_thresholded",
        "Manders_Channel_B_Overlap_Channel_A_thresholded",
        "Fiji_RGB_Crop_Input",
        "Fiji_Log",
    ]
    ordered = [col for col in preferred if col in raw_df.columns]
    remaining = [col for col in raw_df.columns if col not in ordered]
    return raw_df[ordered + remaining]


def reorder_summary_columns(summary_df):
    if summary_df.empty:
        return summary_df
    preferred = [
        "Pair",
        "PCC_Manders_Pair_Label",
        "PCC_Manders_Pair_Type",
        "Metric",
        "N",
        "Metric_Interpretation",
        "Average",
        "SD",
        "SEM",
    ]
    ordered = [col for col in preferred if col in summary_df.columns]
    remaining = [col for col in summary_df.columns if col not in ordered]
    return summary_df[ordered + remaining]


def build_summary(raw_df):
    metrics = ["PCC", "Overlap", "k1", "k2", "M1_original", "M2_original", "M1_thresholded", "M2_thresholded"]
    rows = []
    if raw_df.empty:
        return pd.DataFrame(rows)
    summary_df = raw_df[raw_df["Use_For_Summary"] == True].copy()
    group_columns = ["Pair"]
    if "PCC_Manders_Pair_Type" in summary_df.columns:
        group_columns.append("PCC_Manders_Pair_Type")
    for group_key, pair_df in summary_df.groupby(group_columns, dropna=False):
        if isinstance(group_key, tuple):
            pair = group_key[0]
            pair_type = group_key[1] if len(group_key) > 1 else ""
        else:
            pair = group_key
            pair_type = ""
        for metric in metrics:
            values = pair_df[metric].dropna() if metric in pair_df else pd.Series(dtype=float)
            n = int(values.count())
            sd = float(values.std(ddof=1)) if n > 1 else 0.0
            first_row = pair_df.iloc[0].to_dict() if not pair_df.empty else {}
            rows.append({
                "Pair": pair,
                "PCC_Manders_Pair_Label": first_row.get("PCC_Manders_Pair_Label", ""),
                "PCC_Manders_Pair_Type": pair_type,
                "Metric": metric,
                "Metric_Interpretation": metric_interpretation_for_summary(metric, first_row),
                "Average": float(values.mean()) if n else np.nan,
                "SD": sd,
                "SEM": sd / math.sqrt(n) if n > 1 else 0.0,
                "N": n,
            })
    return pd.DataFrame(rows)


def pcc_manders_pair_label(channel_a_label, channel_b_label, channel_a_role, channel_b_role, pair_type):
    role_a = role_display_name(channel_a_role).lower()
    role_b = role_display_name(channel_b_role).lower()
    if pair_type == "condensate_marker":
        return f"{channel_a_label} condensate vs {channel_b_label} diffuse marker"
    if pair_type == "condensate_condensate":
        return f"{channel_a_label} condensate vs {channel_b_label} condensate"
    if pair_type == "marker_marker":
        return f"{channel_a_label} diffuse marker vs {channel_b_label} diffuse marker"
    return f"{channel_a_label} {role_a} vs {channel_b_label} {role_b}"


def role_display_name(role):
    return {
        "condensate": "Condensate",
        "diffuse_marker": "Diffuse marker",
        "nucleus": "Nucleus/reference",
        "ignore": "Ignored",
    }.get(role, str(role or "ignore"))


def channel_role_counts_for_item(state, item):
    roles = channel_roles_for_item(state, item)
    counts = {role: 0 for role in CHANNEL_ROLE_OPTIONS}
    for channel_name in CHANNELS:
        role = valid_channel_role(roles.get(channel_name), "ignore")
        counts[role] = counts.get(role, 0) + 1
    non_ignored = sum(count for role, count in counts.items() if role != "ignore")
    percent_condensate = (counts.get("condensate", 0) / non_ignored * 100.0) if non_ignored else 0.0
    percent_marker = (counts.get("diffuse_marker", 0) / non_ignored * 100.0) if non_ignored else 0.0
    return counts, non_ignored, percent_condensate, percent_marker


def build_per_cell_channel_role_summary(state, channel_labels):
    rows = []
    for rel, item in sorted(state.get("items", {}).items()):
        if item.get("status", "keep") != "keep":
            continue
        roles = channel_roles_for_item(state, item)
        src, crop_choice = selected_crop_path_for_item(state, rel, item)
        counts, non_ignored, percent_condensate, percent_marker = channel_role_counts_for_item(state, item)
        rows.append({
            "Review_Relative_Path": rel,
            "Image": os.path.basename(src),
            "Crop_Choice": crop_choice,
            "Red_Channel_Label": channel_labels.get("red", "red"),
            "Red_Channel_Role": roles.get("red", "ignore"),
            "Red_Channel_Role_Display": role_display_name(roles.get("red", "ignore")),
            "Green_Channel_Label": channel_labels.get("green", "green"),
            "Green_Channel_Role": roles.get("green", "ignore"),
            "Green_Channel_Role_Display": role_display_name(roles.get("green", "ignore")),
            "Blue_Channel_Label": channel_labels.get("blue", "blue"),
            "Blue_Channel_Role": roles.get("blue", "ignore"),
            "Blue_Channel_Role_Display": role_display_name(roles.get("blue", "ignore")),
            "Condensate_Channel_Count": counts.get("condensate", 0),
            "Diffuse_Marker_Channel_Count": counts.get("diffuse_marker", 0),
            "Nucleus_Reference_Channel_Count": counts.get("nucleus", 0),
            "Ignored_Channel_Count": counts.get("ignore", 0),
            "Non_Ignored_Channel_Count": non_ignored,
            "Percent_Condensate_Among_Non_Ignored": percent_condensate,
            "Percent_Diffuse_Marker_Among_Non_Ignored": percent_marker,
        })
    return pd.DataFrame(rows)


def build_overall_channel_role_summary(state, channel_labels):
    kept_items = [(rel, item) for rel, item in sorted(state.get("items", {}).items()) if item.get("status", "keep") == "keep"]
    total_kept = len(kept_items)
    rows = []
    for channel_name in CHANNELS:
        role_counts = {role: 0 for role in CHANNEL_ROLE_OPTIONS}
        for _rel, item in kept_items:
            role = valid_channel_role(channel_roles_for_item(state, item).get(channel_name), "ignore")
            role_counts[role] = role_counts.get(role, 0) + 1
        non_ignored = total_kept - role_counts.get("ignore", 0)
        rows.append({
            "Channel": channel_name.title(),
            "Channel_Key": channel_name,
            "Protein_Label": channel_labels.get(channel_name, channel_name),
            "Total_Kept_Crops": total_kept,
            "Times_Condensate": role_counts.get("condensate", 0),
            "Times_Diffuse_Marker": role_counts.get("diffuse_marker", 0),
            "Times_Nucleus_Reference": role_counts.get("nucleus", 0),
            "Times_Ignored": role_counts.get("ignore", 0),
            "Percent_Condensate_Of_Kept_Crops": (role_counts.get("condensate", 0) / total_kept * 100.0) if total_kept else 0.0,
            "Percent_Diffuse_Marker_Of_Kept_Crops": (role_counts.get("diffuse_marker", 0) / total_kept * 100.0) if total_kept else 0.0,
            "Percent_Condensate_Among_Non_Ignored": (role_counts.get("condensate", 0) / non_ignored * 100.0) if non_ignored else 0.0,
            "Percent_Diffuse_Marker_Among_Non_Ignored": (role_counts.get("diffuse_marker", 0) / non_ignored * 100.0) if non_ignored else 0.0,
        })
    return pd.DataFrame(rows)


class NoCondensateDetected(ValueError):
    pass







def graph_font(size=22, bold=False):
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size)
        except Exception:
            continue
    return ImageFont.load_default()



























































REGION_COLUMNS = [
    ("Core", "Core_mean"),
    ("Ring_0_2", "Ring_0_2_mean"),
    ("Ring_2_4", "Ring_2_4_mean"),
    ("Ring_4_8", "Ring_4_8_mean"),
]
REGION_DISPLAY_LABELS = ["Core", "0-2", "2-4", "4-8"]


























REFINED_REGIONS = [
    "Core_0_1R",
    "Ring1_1_1.25R",
    "Ring2_1.25_1.5R",
    "Ring3_1.5_2R",
    "Ring4_2_3R",
]
REFINED_REGION_MEAN_COLS = [f"{region}_mean" for region in REFINED_REGIONS]
REFINED_X_LABELS = ["Core", "Ring 1", "Ring 2", "Ring 3", "Ring 4"]
REFINED_MARKS = ["H3K9ac", "H3K9me3", "H3K27ac", "H3K27me3"]
REFINED_SINGLE_MARK_COLORS = {"N1ICD": "#0072B2", "N4ICD": "#D55E00"}
REFINED_CATEGORY_COLORS = {
    "N1ICD - Activating (H3K9ac + H3K27ac)": "#0072B2",
    "N4ICD - Activating (H3K9ac + H3K27ac)": "#E69F00",
    "N1ICD - Repressive (H3K9me3 + H3K27me3)": "#009E73",
    "N4ICD - Repressive (H3K9me3 + H3K27me3)": "#CC79A7",
}



















































def fiji_pcc_manders_script_text(input_path, pair_name, channel_a, channel_b, result_path):
    def bsh_string(value):
        return '"' + str(value).replace('\\', '\\\\').replace('"', '\\"') + '"'

    return f"""
import ij.*;
import ij.process.*;
import ij.plugin.ChannelSplitter;
import java.net.*;
import java.lang.reflect.*;

String inputPath = {bsh_string(input_path)};
String pairName = {bsh_string(pair_name)};
String channelA = {bsh_string(channel_a)};
String channelB = {bsh_string(channel_b)};
String resultPath = {bsh_string(result_path)};
ImagePlus imp = IJ.openImage(inputPath);
if (imp == null) throw new RuntimeException("Could not open RGB crop sent to Fiji: " + inputPath);
ImagePlus[] channels = ChannelSplitter.split(imp);
if (channels == null || channels.length < 3) throw new RuntimeException("PCC/Manders analysis expected Fiji to split RGB into three channels; got " + (channels == null ? 0 : channels.length));
int idxA = channelA.equals("red") ? 0 : (channelA.equals("green") ? 1 : 2);
int idxB = channelB.equals("red") ? 0 : (channelB.equals("green") ? 1 : 2);
ImagePlus imgA = channels[idxA];
ImagePlus imgB = channels[idxB];
imgA.setTitle("A_channel");
imgB.setTitle("B_channel");

int w = imgA.getWidth();
int h = imgA.getHeight();
int slices = imgA.getStackSize();
double[] A = new double[w * h * slices];
double[] B = new double[w * h * slices];
double sumA = 0, sumB = 0, sumA2 = 0, sumB2 = 0, sumAB = 0;
double sumAWhenBPositive = 0, sumBWhenAPositive = 0, sumAWhenBThresholded = 0, sumBWhenAThresholded = 0;
double Amin = 1.0/0.0, Bmin = 1.0/0.0, Amax = -1.0/0.0, Bmax = -1.0/0.0;
int thresholdA = imgA.getProcessor().getAutoThreshold();
int thresholdB = imgB.getProcessor().getAutoThreshold();

for (int kk = 1; kk <= slices; kk++) {{
    imgA.setSlice(kk);
    imgB.setSlice(kk);
    ImageProcessor ipa = imgA.getProcessor();
    ImageProcessor ipb = imgB.getProcessor();
    for (int yy = 0; yy < h; yy++) {{
        for (int xx = 0; xx < w; xx++) {{
            int offset = (kk - 1) * w * h + yy * w + xx;
            double a = ipa.getPixel(xx, yy);
            double b = ipb.getPixel(xx, yy);
            A[offset] = a;
            B[offset] = b;
            sumA += a;
            sumB += b;
            sumA2 += a * a;
            sumB2 += b * b;
            sumAB += a * b;
            if (b > 0) sumAWhenBPositive += a;
            if (a > 0) sumBWhenAPositive += b;
            if (b > thresholdB) sumAWhenBThresholded += a;
            if (a > thresholdA) sumBWhenAThresholded += b;
            if (a < Amin) Amin = a;
            if (a > Amax) Amax = a;
            if (b < Bmin) Bmin = b;
            if (b > Bmax) Bmax = b;
        }}
    }}
}}

double pixels = w * h * slices;
double pccDen = Math.sqrt((pixels * sumA2 - sumA * sumA) * (pixels * sumB2 - sumB * sumB));
double pcc = pccDen > 0 ? (pixels * sumAB - sumA * sumB) / pccDen : Double.NaN;
double overlap = (sumA2 > 0 && sumB2 > 0) ? sumAB / Math.sqrt(sumA2 * sumB2) : Double.NaN;
double k1 = sumA2 > 0 ? sumAB / sumA2 : Double.NaN;
double k2 = sumB2 > 0 ? sumAB / sumB2 : Double.NaN;
double m1Original = sumA > 0 ? sumAWhenBPositive / sumA : Double.NaN;
double m2Original = sumB > 0 ? sumBWhenAPositive / sumB : Double.NaN;
double m1Thresholded = sumA > 0 ? sumAWhenBThresholded / sumA : Double.NaN;
double m2Thresholded = sumB > 0 ? sumBWhenAThresholded / sumB : Double.NaN;

String resultHeader = "JACOP_PCC_MANDERS\tPair\tPixels_Analyzed\tPCC\tOverlap\tk1\tk2\tJACoP_threshold_A\tJACoP_threshold_B\tM1_original\tM2_original\tM1_thresholded\tM2_thresholded";
String resultValues = "JACOP_PCC_MANDERS\t" + pairName + "\t" + ((int)pixels) + "\t" + pcc + "\t" + overlap + "\t" + k1 + "\t" + k2 + "\t" + thresholdA + "\t" + thresholdB + "\t" + m1Original + "\t" + m2Original + "\t" + m1Thresholded + "\t" + m2Thresholded;
print(resultHeader);
print(resultValues);
java.io.PrintWriter resultWriter = new java.io.PrintWriter(resultPath, "UTF-8");
resultWriter.println(resultHeader);
resultWriter.println(resultValues);
resultWriter.close();
IJ.run("Close All");
"""



def pcc_channel_labels_from_state(state):
    return {
        "red": str(state.get("pcc_red_label", "") or "").strip() or "red",
        "green": str(state.get("pcc_green_label", "") or "").strip() or "green",
        "blue": str(state.get("pcc_blue_label", "") or "").strip() or "blue",
    }


def channel_label_rows_from_state(state):
    labels = pcc_channel_labels_from_state(state)
    return pd.DataFrame([
        {"Channel": "Red", "Channel_Key": "red", "Inserted_Label": labels["red"]},
        {"Channel": "Green", "Channel_Key": "green", "Inserted_Label": labels["green"]},
        {"Channel": "Blue", "Channel_Key": "blue", "Inserted_Label": labels["blue"]},
    ])


def apply_label_post_fields(state, data):
    """Persist labels typed immediately before an analysis button is clicked."""
    for key in (
        "pcc_red_label",
        "pcc_green_label",
        "pcc_blue_label",
        "analysis_input_mode",
        "pcc_analysis_input_mode",
        "count_size_analysis_input_mode",
        "red_channel_role",
        "green_channel_role",
        "blue_channel_role",
        "pcc_pair_mode",
        "cleaning_radius_px",
        "cleaning_sensitivity",
        "cleaning_min_area_px",
        "cleaning_background_removal",
        "cleaning_amount",
        "red_cleaning_amount",
        "green_cleaning_amount",
        "blue_cleaning_amount",
        "fixed_cell_replicate_id",
        "fixed_cell_condition",
        "fixed_cell_condensate_label",
        "fixed_cell_ms2_label",
        "fixed_cell_modifier_label",
        "partition_target_label",
        "partition_red_label",
        "partition_green_label",
        "partition_blue_label",
        "partition_condensate_label",
        "partition_replicate_id",
        "partition_condition",
    ):
        if key in data:
            state[key] = data.get(key, state.get(key, "")).strip()
    state["analysis_input_mode"] = normalize_analysis_input_mode(state.get("analysis_input_mode", "raw"))
    if "analysis_input_mode" in data:
        for key in ("pcc_analysis_input_mode", "count_size_analysis_input_mode"):
            state[key] = state["analysis_input_mode"]
    for key in ("pcc_analysis_input_mode", "count_size_analysis_input_mode"):
        state[key] = normalize_analysis_input_mode(state.get(key, state["analysis_input_mode"]))
    state["red_channel_role"] = valid_channel_role(state.get("red_channel_role"), "diffuse_marker")
    state["green_channel_role"] = valid_channel_role(state.get("green_channel_role"), "condensate")
    state["blue_channel_role"] = valid_channel_role(state.get("blue_channel_role"), "nucleus")
    state["pcc_pair_mode"] = valid_pcc_pair_mode(state.get("pcc_pair_mode"), "red_green")
    for channel_name in CHANNELS:
        state[f"count_size_{channel_name}"] = f"count_size_{channel_name}" in data
    if "count_size_particle_size" in data:
        _, particle_size_range = normalize_count_size_particle_size(data.get("count_size_particle_size"))
        state["count_size_particle_size"] = particle_size_range
    for key, fallback in (
        ("fixed_cell_n1icd_channel", "green"),
        ("fixed_cell_ms2_channel", "red"),
        ("fixed_cell_dapi_channel", "blue"),
        ("fixed_cell_modifier_channel", "blue"),
    ):
        if key in data:
            state[key] = valid_channel(data.get(key), fallback)
    if "fixed_cell_color_mode" in data:
        mode = str(data.get("fixed_cell_color_mode", "three_color_dapi")).strip().lower()
        state["fixed_cell_color_mode"] = mode if mode in ("two_color", "three_color_dapi", "three_color_modifier") else "three_color_dapi"
    if "fixed_cell_annotation_mode" in data:
        annotation_mode = str(data.get("fixed_cell_annotation_mode", "manual_required")).strip().lower()
        state["fixed_cell_annotation_mode"] = annotation_mode if annotation_mode in ("manual_required", "prefer_manual", "automatic") else "manual_required"
    if "fixed_cell_intensity_source_mode" in data:
        state["fixed_cell_intensity_source_mode"] = normalize_fixed_cell_intensity_source_mode(
            data.get("fixed_cell_intensity_source_mode")
        )
    if "partition_target_channel" in data:
        state["partition_target_channel"] = valid_channel(data.get("partition_target_channel"), "red")
    if "partition_condensate_channel" in data:
        state["partition_condensate_channel"] = valid_channel(data.get("partition_condensate_channel"), "red")
    if "partition_intensity_source_mode" in data:
        state["partition_intensity_source_mode"] = normalize_fixed_cell_intensity_source_mode(
            data.get("partition_intensity_source_mode")
        )
    if "partition_boundary_exclusion_px" in data:
        state["partition_boundary_exclusion_px"] = clamp_int(
            data.get("partition_boundary_exclusion_px"), 1, 0, 10
        )
    if "fixed_cell_condensate_source" in data:
        condensate_source = str(data.get("fixed_cell_condensate_source", "count_size")).strip().lower()
        state["fixed_cell_condensate_source"] = (
            condensate_source if condensate_source in ("count_size", "manual") else "count_size"
        )
    if "fixed_cell_focus_rule" in data:
        rule = str(data.get("fixed_cell_focus_rule", "brightest")).strip().lower()
        state["fixed_cell_focus_rule"] = rule if rule in ("brightest", "sum") else "brightest"
    for key, fallback, minimum, maximum in (
        ("fixed_cell_ms2_min_area_px2", 3, 1, 10000),
        ("fixed_cell_background_inner_px", 3, 1, 100),
        ("fixed_cell_background_outer_px", 12, 2, 500),
    ):
        if key in data:
            state[key] = clamp_int(data.get(key), fallback, minimum, maximum)
    if state.get("fixed_cell_background_outer_px", 12) <= state.get("fixed_cell_background_inner_px", 3):
        state["fixed_cell_background_outer_px"] = state.get("fixed_cell_background_inner_px", 3) + 1
    cleaning_params = cleaning_params_from_state(state)
    state["cleaning_radius_px"] = cleaning_params["radius_px"]
    state["cleaning_sensitivity"] = cleaning_params["sensitivity"]
    state["cleaning_min_area_px"] = cleaning_params["min_area_px"]
    state["cleaning_background_removal"] = cleaning_params["background_removal"]
    state["cleaning_amount"] = cleaning_params["amount"]
    channel_params = channel_cleaning_params_from_mapping(state, cleaning_params)
    for channel_name in CHANNELS:
        state[f"{channel_name}_cleaning_amount"] = channel_params[channel_name]["amount"]
    if "analysis_mode" in data:
        state["analysis_mode"] = valid_analysis_mode(data.get("analysis_mode"), state.get("analysis_mode", "red_green"))
        state["stage3_mode"] = "dapi_red_green" if state["analysis_mode"] == "red_green" else "no_dapi_three_color"
    if "pixel_um" in data:
        try:
            state["pixel_um"] = float(data.get("pixel_um", state.get("pixel_um", 0.1290271)))
        except (TypeError, ValueError):
            pass
    save_state(state)


def channel_display_name(channel, channel_labels=None):
    key = str(channel).lower()
    labels = channel_labels or {}
    return labels.get(key) or {"red": "red", "green": "green", "blue": "blue"}.get(key, str(channel))


def pcc_pair_interpretation_columns(channel_a, channel_b, channel_labels=None):
    a_name = channel_display_name(channel_a, channel_labels)
    b_name = channel_display_name(channel_b, channel_labels)
    return {
        "Channel_A_Label": a_name,
        "Channel_B_Label": b_name,
        "PCC_Interpretation": f"Pearson correlation between {a_name} and {b_name} signals",
        "Overlap_Interpretation": f"JACoP overlap coefficient between {a_name} and {b_name} signals",
        "k1_Interpretation": f"JACoP k1 for {a_name} relative to {b_name}",
        "k2_Interpretation": f"JACoP k2 for {b_name} relative to {a_name}",
    }


def manders_interpretation_columns(channel_a, channel_b, values_swapped=False, channel_labels=None):
    # JACoP-style M1 is the fraction of channel A that overlaps channel B.
    # The app calls Fiji with the displayed channel order, so M1/M2 should
    # normally follow channel A/channel B directly.
    m1_channel = channel_b if values_swapped else channel_a
    m1_overlap = channel_a if values_swapped else channel_b
    m2_channel = channel_a if values_swapped else channel_b
    m2_overlap = channel_b if values_swapped else channel_a
    m1_channel_name = channel_display_name(m1_channel, channel_labels)
    m1_overlap_name = channel_display_name(m1_overlap, channel_labels)
    m2_channel_name = channel_display_name(m2_channel, channel_labels)
    m2_overlap_name = channel_display_name(m2_overlap, channel_labels)
    return {
        "M1_Channel_Fraction_Measured": m1_channel_name,
        "M1_Overlap_Channel": m1_overlap_name,
        "M1_original_Interpretation": f"fraction of {m1_channel_name} signal in pixels where {m1_overlap_name} signal is present (>0)",
        "M1_thresholded_Interpretation": f"fraction of {m1_channel_name} signal in pixels where {m1_overlap_name} signal is above the JACoP threshold",
        "M2_Channel_Fraction_Measured": m2_channel_name,
        "M2_Overlap_Channel": m2_overlap_name,
        "M2_original_Interpretation": f"fraction of {m2_channel_name} signal in pixels where {m2_overlap_name} signal is present (>0)",
        "M2_thresholded_Interpretation": f"fraction of {m2_channel_name} signal in pixels where {m2_overlap_name} signal is above the JACoP threshold",
        "Manders_Display_Order_Note": "M1/M2 values swapped after Fiji call to match this displayed interpretation" if values_swapped else "M1/M2 follow Fiji channel A/channel B order",
    }


def metric_interpretation_for_summary(metric, row):
    if metric == "PCC":
        return row.get("PCC_Interpretation", "")
    if metric == "Overlap":
        return row.get("Overlap_Interpretation", "")
    if metric == "k1":
        return row.get("k1_Interpretation", "")
    if metric == "k2":
        return row.get("k2_Interpretation", "")
    if metric == "M1_original":
        return row.get("M1_original_Interpretation", "")
    if metric == "M2_original":
        return row.get("M2_original_Interpretation", "")
    if metric == "M1_thresholded":
        return row.get("M1_thresholded_Interpretation", "")
    if metric == "M2_thresholded":
        return row.get("M2_thresholded_Interpretation", "")
    return ""

def parse_fiji_pcc_manders_output(stdout_text):
    header = None
    values = None
    for line in stdout_text.splitlines():
        if not line.startswith("JACOP_PCC_MANDERS\t"):
            continue
        parts = line.rstrip("\n").split("\t")
        if len(parts) > 1 and parts[1] == "Pair":
            header = parts[1:]
        else:
            values = parts[1:]
    if not header or not values:
        raise ValueError("Fiji did not return PCC/Manders output. Output was:\n" + stdout_text[-2000:])
    row = dict(zip(header, values))
    for key in ["Pixels_Analyzed", "PCC", "Overlap", "k1", "k2", "JACoP_threshold_A", "JACoP_threshold_B", "M1_original", "M2_original", "M1_thresholded", "M2_thresholded"]:
        try:
            row[key] = float(row[key])
        except Exception:
            row[key] = np.nan
    row["Pixels_Analyzed"] = int(row["Pixels_Analyzed"]) if np.isfinite(row["Pixels_Analyzed"]) else 0
    row["Analysis_Engine"] = "Fiji headless direct PCC/Manders calculation with ImageJ auto-thresholds"
    row["Use_For_Summary"] = True
    return row


def pcc_manders_pair_cache_fingerprint(rgb_crop_path, pair_name, channel_a, channel_b):
    digest = hashlib.sha256()
    digest.update(b"fiji-pcc-manders-pair-cache-v3-windows-direct-java-20261009")
    digest.update(str(pair_name).encode("utf-8"))
    digest.update(str(channel_a).encode("utf-8"))
    digest.update(str(channel_b).encode("utf-8"))
    with open(rgb_crop_path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def run_fiji_pcc_manders_pair(rgb_crop_path, pair_name, channel_a, channel_b, log_path):
    if not os.path.isfile(FIJI_EXECUTABLE):
        raise ValueError(f"Fiji executable was not found: {FIJI_EXECUTABLE}")
    fingerprint = pcc_manders_pair_cache_fingerprint(rgb_crop_path, pair_name, channel_a, channel_b)
    cache_dir = os.path.dirname(log_path)
    cache_name = hashlib.sha256(os.path.basename(log_path).encode("utf-8")).hexdigest()[:16]
    cache_path = os.path.join(cache_dir, f"pcc_{cache_name}.json")
    if os.path.isfile(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as handle:
                cached = json.load(handle)
            if cached.get("fingerprint") == fingerprint and isinstance(cached.get("row"), dict):
                return cached["row"]
        except Exception:
            pass
    with tempfile.NamedTemporaryFile("w", suffix=".tsv", delete=False) as result_handle:
        result_path = result_handle.name
    with tempfile.NamedTemporaryFile("w", suffix=".bsh", delete=False) as handle:
        handle.write(fiji_pcc_manders_script_text(rgb_crop_path, pair_name, channel_a, channel_b, result_path))
        script_path = handle.name
    try:
        command = (
            windows_direct_fiji_java_command(rgb_crop_path, result_path, pair_name, channel_a, channel_b)
            if os.name == "nt"
            else [FIJI_EXECUTABLE, "--headless", "--console", "--run", script_path]
        )
        completed = subprocess.run(
            command,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=180,
        )
        if os.name == "nt" and completed.returncode == 0:
            wait_deadline = time.monotonic() + 30
            while time.monotonic() < wait_deadline:
                if os.path.isfile(result_path) and os.path.getsize(result_path) > 0:
                    break
                time.sleep(0.1)
        result_text = ""
        if os.path.isfile(result_path) and os.path.getsize(result_path) > 0:
            with open(result_path, "r", encoding="utf-8", errors="replace") as result_file:
                result_text = result_file.read()
    finally:
        try:
            os.unlink(script_path)
        except OSError:
            pass
        try:
            os.unlink(result_path)
        except OSError:
            pass
    stdout_text = completed.stdout or ""
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"FIJI_EXECUTABLE: {FIJI_EXECUTABLE}\nRETURN_CODE: {completed.returncode}\nRESULT_FILE_RECEIVED: {bool(result_text)}\n\n")
        f.write(stdout_text)
        if result_text:
            f.write("\nRESULT_FILE_OUTPUT:\n")
            f.write(result_text)
    if completed.returncode != 0:
        raise RuntimeError(f"Fiji PCC/Manders script failed with code {completed.returncode}. See log: {log_path}\n\n{stdout_text[-2000:]}")
    if not result_text and not stdout_text.strip():
        raise RuntimeError(f"Fiji returned no PCC/Manders result file or console output. Executable: {FIJI_EXECUTABLE}. See log: {log_path}")
    row = parse_fiji_pcc_manders_output(result_text or stdout_text)
    if os.name == "nt":
        row["Analysis_Engine"] = "Bundled Fiji Java/ImageJ direct PCC/Manders calculation with ImageJ auto-thresholds"
    tmp_path = os.path.join(cache_dir, f"tmp_{uuid4().hex[:12]}.json")
    try:
        with open(tmp_path, "w", encoding="utf-8") as handle:
            json.dump({"fingerprint": fingerprint, "row": row}, handle)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp_path, cache_path)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    return row


def save_fiji_channel_inputs(src, output_dir, base_name):
    image_yxc, layout = read_yxc(src)
    os.makedirs(output_dir, exist_ok=True)
    paths = {}
    for channel_name, index in CHANNELS.items():
        channel = image_yxc[..., index]
        path = os.path.join(output_dir, f"{base_name}_{channel_name}.tif")
        tiff.imwrite(path, channel)
        paths[channel_name] = path
    return paths, layout


def count_size_channels_from_state(state):
    selected = []
    for channel_name in CHANNELS:
        key = f"count_size_{channel_name}"
        if key not in state:
            selected.append(channel_name)
        elif bool(state.get(key)):
            selected.append(channel_name)
    return selected


def normalize_count_size_particle_size(value, default="10-Infinity"):
    """Return Fiji's minimum particle area and a canonical N-Infinity label."""
    text = str(value if value is not None else default).strip()
    if not text:
        text = default
    match = re.fullmatch(
        r"\s*(\d+(?:\.\d+)?)\s*(?:(?:-|\N{EN DASH}|\N{EM DASH}|to)\s*(?:inf(?:inity)?|\N{INFINITY}))?\s*",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        raise ValueError("Particle size must look like 10-Infinity (or a non-negative minimum such as 10).")
    minimum = float(match.group(1))
    if not np.isfinite(minimum) or minimum < 0:
        raise ValueError("Particle-size minimum must be a finite non-negative number.")
    minimum_label = str(int(minimum)) if minimum.is_integer() else f"{minimum:g}"
    return minimum, f"{minimum_label}-Infinity"


def fiji_count_size_macro_text():
    return """args = getArgument();
parts = split(args, "|||");
input = parts[0];
results = parts[1];
summary = parts[2];
mask = parts[3];
minimum = parts[4];
manualThresholdMin = parseFloat(parts[5]);
manualThresholdMax = parseFloat(parts[6]);
display = "";
if (parts.length > 7) display = parts[7];
open(input);
run("8-bit");
if (manualThresholdMin >= 0 && manualThresholdMax >= 0) {
    setThreshold(manualThresholdMin, manualThresholdMax);
} else {
    setAutoThreshold("Default dark");
}
setOption("BlackBackground", false);
run("Convert to Mask");
run("Watershed");
saveAs("Tiff", mask);
if (display != "") {
    run("Duplicate...", "title=Fiji_Watershed_Display");
    selectWindow("Fiji_Watershed_Display");
    previewWidth = getWidth();
    previewHeight = getHeight();
    for (previewY = 0; previewY < previewHeight; previewY++) {
        for (previewX = 0; previewX < previewWidth; previewX++) {
            setPixel(previewX, previewY, 255 - getPixel(previewX, previewY));
        }
    }
    saveAs("PNG", display);
    close();
}
run("Set Measurements...", "area centroid bounding display redirect=None decimal=3");
run("Analyze Particles...", "size=" + minimum + "-Infinity display clear");
count = nResults;
totalArea = 0;
print("FIJI_COUNT_SIZE_RESULTS_BEGIN");
print("Particle\\tArea\\tX\\tY\\tBX\\tBY\\tWidth\\tHeight");
for (i = 0; i < count; i++) {
    area = getResult("Area", i);
    x = getResult("X", i);
    y = getResult("Y", i);
    bx = getResult("BX", i);
    by = getResult("BY", i);
    width = getResult("Width", i);
    height = getResult("Height", i);
    totalArea = totalArea + area;
    print((i + 1) + "\\t" + area + "\\t" + x + "\\t" + y + "\\t" + bx + "\\t" + by + "\\t" + width + "\\t" + height);
}
print("FIJI_COUNT_SIZE_RESULTS_END");
averageArea = 0;
if (count > 0) averageArea = totalArea / count;
percentArea = 0;
if (getWidth() * getHeight() > 0) percentArea = 100 * totalArea / (getWidth() * getHeight());
print("FIJI_COUNT_SIZE_SUMMARY\\tCount\\tTotal Area\\tAverage Size\\t%Area");
print("FIJI_COUNT_SIZE_SUMMARY\\t" + count + "\\t" + totalArea + "\\t" + averageArea + "\\t" + percentArea);
run("Close All");
"""


def normalize_threshold_range(threshold_min=None, threshold_max=None):
    if threshold_min is None or threshold_max is None:
        return None, None
    try:
        lower = max(0, min(255, int(float(threshold_min))))
        upper = max(0, min(255, int(float(threshold_max))))
    except (TypeError, ValueError):
        return None, None
    if upper < lower:
        lower, upper = upper, lower
    return lower, upper


def parse_fiji_count_size_output(stdout_text, results_path, summary_path):
    particles = []
    in_particles = False
    for line in str(stdout_text).splitlines():
        line = line.strip()
        if line == "FIJI_COUNT_SIZE_RESULTS_BEGIN":
            in_particles = True
            continue
        if line == "FIJI_COUNT_SIZE_RESULTS_END":
            in_particles = False
            continue
        if in_particles:
            if not line or line.startswith("Particle"):
                continue
            parts = line.split("\t")
            if len(parts) < 4:
                continue
            try:
                row = {
                    "Particle": int(float(parts[0])),
                    "Area_px2": float(parts[1]),
                    "Centroid_X": float(parts[2]),
                    "Centroid_Y": float(parts[3]),
                }
                if len(parts) >= 8:
                    row.update({
                        "BX": float(parts[4]),
                        "BY": float(parts[5]),
                        "Width": float(parts[6]),
                        "Height": float(parts[7]),
                    })
                particles.append(row)
            except ValueError:
                continue
    summary = {"Condensate_Count": 0, "Total_Area_px2": 0.0, "Average_Area_px2": 0.0, "Percent_Area": 0.0}
    for line in str(stdout_text).splitlines():
        line = line.strip()
        if not line.startswith("FIJI_COUNT_SIZE_SUMMARY\t"):
            continue
        parts = line.split("\t")
        if len(parts) >= 5 and parts[1] != "Count":
            try:
                summary = {
                    "Condensate_Count": int(float(parts[1])),
                    "Total_Area_px2": float(parts[2]),
                    "Average_Area_px2": float(parts[3]),
                    "Percent_Area": float(parts[4]),
                }
            except ValueError:
                pass
    os.makedirs(os.path.dirname(results_path), exist_ok=True)
    os.makedirs(os.path.dirname(summary_path), exist_ok=True)
    pd.DataFrame(particles).to_csv(results_path, index=False)
    pd.DataFrame([summary]).to_csv(summary_path, index=False)
    return summary, particles


def run_fiji_count_size_macro(input_path, results_path, summary_path, mask_path, min_area=0, threshold_min=None, threshold_max=None, display_path=None):
    if not os.path.isfile(FIJI_EXECUTABLE):
        raise ValueError(f"Fiji executable was not found: {FIJI_EXECUTABLE}")
    for path in (results_path, summary_path, mask_path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
    if display_path:
        os.makedirs(os.path.dirname(display_path), exist_ok=True)
    with tempfile.NamedTemporaryFile("w", suffix=".ijm", delete=False, encoding="utf-8") as handle:
        handle.write(fiji_count_size_macro_text())
        macro_path = handle.name
    lower, upper = normalize_threshold_range(threshold_min, threshold_max)
    try:
        minimum_value = float(min_area)
    except (TypeError, ValueError):
        raise ValueError("Fiji particle-size minimum must be numeric.")
    if not np.isfinite(minimum_value) or minimum_value < 0:
        raise ValueError("Fiji particle-size minimum must be a finite non-negative number.")
    minimum_argument = str(int(minimum_value)) if minimum_value.is_integer() else f"{minimum_value:g}"
    threshold_min_argument = "-1" if lower is None else str(int(lower))
    threshold_max_argument = "-1" if upper is None else str(int(upper))
    completed = subprocess.run(
        [FIJI_EXECUTABLE, "--headless", "--console", "-macro", macro_path, "|||".join([input_path, results_path, summary_path, mask_path, minimum_argument, threshold_min_argument, threshold_max_argument, display_path or ""])],
        text=True,
        capture_output=True,
        check=False,
        timeout=120,
    )
    try:
        os.unlink(macro_path)
    except OSError:
        pass
    if completed.returncode != 0:
        raise RuntimeError((completed.stdout or "") + "\n" + (completed.stderr or ""))
    summary, particles = parse_fiji_count_size_output(completed.stdout, results_path, summary_path)
    return {"summary": summary, "particles": particles, "completed": completed}



COUNT_SIZE_BATCH_CACHE_VERSION = "fiji-count-size-batch-v1-20261009"
ANALYSIS_PROGRESS_LOCK = threading.Lock()
ANALYSIS_PROGRESS = {
    "active": False,
    "stage": "idle",
    "done": 0,
    "total": 0,
    "percent": 0,
    "message": "No count/size analysis is running.",
}


def set_analysis_progress(active, stage, done=0, total=0, message=""):
    total = max(0, int(total or 0))
    done = max(0, min(int(done or 0), total)) if total else max(0, int(done or 0))
    percent = int(round(100.0 * done / total)) if total else (100 if stage == "complete" else 0)
    with ANALYSIS_PROGRESS_LOCK:
        ANALYSIS_PROGRESS.update({
            "active": bool(active),
            "stage": str(stage or "idle"),
            "done": done,
            "total": total,
            "percent": percent,
            "message": str(message or ""),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        })


def get_analysis_progress():
    with ANALYSIS_PROGRESS_LOCK:
        return dict(ANALYSIS_PROGRESS)


def count_size_job_fingerprint(input_path, min_area, threshold_min, threshold_max):
    stat = os.stat(input_path)
    payload = {
        "version": COUNT_SIZE_BATCH_CACHE_VERSION,
        "input": os.path.abspath(input_path),
        "size": int(stat.st_size),
        "mtime_ns": int(stat.st_mtime_ns),
        "min_area": int(min_area),
        "threshold_min": threshold_min,
        "threshold_max": threshold_max,
    }
    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()


def count_size_array_fingerprint(channel, min_area, threshold_min, threshold_max):
    array = np.ascontiguousarray(channel)
    digest = hashlib.sha256()
    digest.update(COUNT_SIZE_BATCH_CACHE_VERSION.encode("utf-8"))
    digest.update(str(array.shape).encode("utf-8"))
    digest.update(str(array.dtype).encode("utf-8"))
    digest.update(array.tobytes())
    digest.update(json.dumps({
        "min_area": int(min_area),
        "threshold_min": threshold_min,
        "threshold_max": threshold_max,
    }, sort_keys=True).encode("utf-8"))
    return digest.hexdigest()


def load_cached_count_size_job(job):
    required = (job["results_path"], job["summary_path"], job["mask_path"], job["cache_path"])
    if not all(os.path.isfile(path) for path in required):
        return None
    try:
        with open(job["cache_path"], "r", encoding="utf-8") as handle:
            cache = json.load(handle)
        if cache.get("fingerprint") != job["fingerprint"]:
            return None
        summary_frame = pd.read_csv(job["summary_path"])
        if summary_frame.empty:
            return None
        row = summary_frame.iloc[0]
        summary = {
            "Condensate_Count": int(float(row["Condensate_Count"])),
            "Total_Area_px2": float(row["Total_Area_px2"]),
            "Average_Area_px2": float(row["Average_Area_px2"]),
            "Percent_Area": float(row["Percent_Area"]),
        }
        particle_frame = pd.read_csv(job["results_path"])
        particles = particle_frame.to_dict("records") if not particle_frame.empty else []
        return {"summary": summary, "particles": particles, "cached": True}
    except Exception:
        return None


def save_count_size_job_cache(job):
    cache = {
        "fingerprint": job["fingerprint"],
        "version": COUNT_SIZE_BATCH_CACHE_VERSION,
        "completed_at": datetime.now().isoformat(timespec="seconds"),
    }
    tmp = job["cache_path"] + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(cache, handle, indent=2, sort_keys=True)
    os.replace(tmp, job["cache_path"])


def fiji_count_size_batch_macro_text():
    return r'''manifest = getArgument();
text = File.openAsString(manifest);
lines = split(text, "\n");
setBatchMode(true);
completedJobs = 0;
for (jobIndex = 0; jobIndex < lines.length; jobIndex++) {
    line = replace(lines[jobIndex], "\r", "");
    if (lengthOf(line) == 0) continue;
    parts = split(line, "\t");
    input = parts[0];
    results = parts[1];
    summary = parts[2];
    mask = parts[3];
    minimum = parts[4];
    manualThresholdMin = parseFloat(parts[5]);
    manualThresholdMax = parseFloat(parts[6]);
    open(input);
    run("8-bit");
    if (manualThresholdMin >= 0 && manualThresholdMax >= 0) {
        setThreshold(manualThresholdMin, manualThresholdMax);
    } else {
        setAutoThreshold("Default dark");
    }
    setOption("BlackBackground", false);
    run("Convert to Mask");
    run("Watershed");
    imageArea = getWidth() * getHeight();
    saveAs("Tiff", mask);
    run("Set Measurements...", "area centroid bounding display redirect=None decimal=3");
    run("Analyze Particles...", "size=" + minimum + "-Infinity display clear");
    count = nResults;
    totalArea = 0;
    File.saveString("Particle,Area_px2,Centroid_X,Centroid_Y,BX,BY,Width,Height\n", results);
    for (i = 0; i < count; i++) {
        area = getResult("Area", i);
        x = getResult("X", i);
        y = getResult("Y", i);
        bx = getResult("BX", i);
        by = getResult("BY", i);
        width = getResult("Width", i);
        height = getResult("Height", i);
        totalArea = totalArea + area;
        File.append((i + 1) + "," + area + "," + x + "," + y + "," + bx + "," + by + "," + width + "," + height + "\n", results);
    }
    averageArea = 0;
    if (count > 0) averageArea = totalArea / count;
    percentArea = 0;
    if (imageArea > 0) percentArea = 100 * totalArea / imageArea;
    File.saveString("Condensate_Count,Total_Area_px2,Average_Area_px2,Percent_Area\n" + count + "," + totalArea + "," + averageArea + "," + percentArea + "\n", summary);
    run("Clear Results");
    close();
    completedJobs++;
    print("FIJI_COUNT_SIZE_BATCH_PROGRESS\t" + completedJobs);
}
setBatchMode(false);
print("FIJI_COUNT_SIZE_BATCH_COMPLETE\t" + completedJobs);
'''


def read_count_size_job_outputs(job):
    summary_frame = pd.read_csv(job["summary_path"])
    if summary_frame.empty:
        raise RuntimeError("Fiji batch produced an empty summary: " + job["summary_path"])
    row = summary_frame.iloc[0]
    summary = {
        "Condensate_Count": int(float(row["Condensate_Count"])),
        "Total_Area_px2": float(row["Total_Area_px2"]),
        "Average_Area_px2": float(row["Average_Area_px2"]),
        "Percent_Area": float(row["Percent_Area"]),
    }
    particle_frame = pd.read_csv(job["results_path"])
    particles = particle_frame.to_dict("records") if not particle_frame.empty else []
    return {"summary": summary, "particles": particles, "cached": False}


def run_fiji_count_size_batch(jobs, progress_callback=None):
    if not jobs:
        return {}
    if not os.path.isfile(FIJI_EXECUTABLE):
        raise ValueError(f"Fiji executable was not found: {FIJI_EXECUTABLE}")
    for job in jobs:
        for path in (job["results_path"], job["summary_path"], job["mask_path"]):
            os.makedirs(os.path.dirname(path), exist_ok=True)
    with tempfile.NamedTemporaryFile("w", suffix=".ijm", delete=False, encoding="utf-8") as macro_handle:
        macro_handle.write(fiji_count_size_batch_macro_text())
        macro_path = macro_handle.name
    with tempfile.NamedTemporaryFile("w", suffix=".tsv", delete=False, encoding="utf-8", newline="") as manifest_handle:
        manifest_path = manifest_handle.name
        for job in jobs:
            lower, upper = normalize_threshold_range(job.get("threshold_min"), job.get("threshold_max"))
            fields = [
                job["input_path"], job["results_path"], job["summary_path"], job["mask_path"],
                str(int(job.get("min_area", 0))), "-1" if lower is None else str(lower), "-1" if upper is None else str(upper),
            ]
            if any("\t" in value or "\n" in value or "\r" in value for value in fields):
                raise ValueError("A count/size path contains a tab or newline and cannot be batched.")
            manifest_handle.write("\t".join(fields) + "\n")
    output_lines = []
    try:
        process = subprocess.Popen(
            [FIJI_EXECUTABLE, "--headless", "--console", "-macro", macro_path, manifest_path],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            bufsize=1,
        )
        completed = 0
        assert process.stdout is not None
        for line in process.stdout:
            output_lines.append(line)
            if line.strip().startswith("FIJI_COUNT_SIZE_BATCH_PROGRESS\t"):
                completed += 1
                if progress_callback:
                    progress_callback(completed, len(jobs))
        return_code = process.wait()
        if return_code != 0:
            raise RuntimeError("".join(output_lines))
        results = {}
        for job in jobs:
            result = read_count_size_job_outputs(job)
            save_count_size_job_cache(job)
            results[job["job_id"]] = result
        return results
    finally:
        for path in (macro_path, manifest_path):
            try:
                os.unlink(path)
            except OSError:
                pass

def fiji_count_size_threshold_probe_macro_text():
    return """input = getArgument();
open(input);
run("8-bit");
setAutoThreshold("Default dark");
getThreshold(lower, upper);
print("FIJI_THRESHOLD\\t" + lower + "\\t" + upper);
getHistogram(values, counts, 256);
print("FIJI_HISTOGRAM_BEGIN");
for (i = 0; i < counts.length; i++) {
    print(i + "\\t" + counts[i]);
}
print("FIJI_HISTOGRAM_END");
run("Close All");
"""


def parse_fiji_count_size_threshold_probe_output(stdout_text):
    threshold_min = None
    threshold_max = None
    histogram = [0] * 256
    in_histogram = False
    for line in str(stdout_text).splitlines():
        line = line.strip()
        if line == "FIJI_HISTOGRAM_BEGIN":
            in_histogram = True
            continue
        if line == "FIJI_HISTOGRAM_END":
            in_histogram = False
            continue
        parts = line.split("\t")
        if len(parts) >= 3 and parts[0] == "FIJI_THRESHOLD":
            threshold_min, threshold_max = normalize_threshold_range(parts[1], parts[2])
            continue
        if in_histogram and len(parts) >= 2:
            try:
                index = int(float(parts[0]))
                count = int(float(parts[1]))
            except ValueError:
                continue
            if 0 <= index < 256:
                histogram[index] = count
    if threshold_min is None or threshold_max is None:
        raise RuntimeError("Fiji did not report a Default dark threshold.\n" + str(stdout_text or ""))
    return threshold_min, threshold_max, histogram


def run_fiji_count_size_threshold_probe(input_path):
    if not os.path.isfile(FIJI_EXECUTABLE):
        raise ValueError(f"Fiji executable was not found: {FIJI_EXECUTABLE}")
    with tempfile.NamedTemporaryFile("w", suffix=".ijm", delete=False, encoding="utf-8") as handle:
        handle.write(fiji_count_size_threshold_probe_macro_text())
        macro_path = handle.name
    completed = subprocess.run(
        [FIJI_EXECUTABLE, "--headless", "--console", "-macro", macro_path, input_path],
        text=True,
        capture_output=True,
        check=False,
        timeout=120,
    )
    try:
        os.unlink(macro_path)
    except OSError:
        pass
    if completed.returncode != 0:
        raise RuntimeError((completed.stdout or "") + "\n" + (completed.stderr or ""))
    threshold_min, threshold_max, histogram = parse_fiji_count_size_threshold_probe_output(completed.stdout)
    return {"threshold_min": threshold_min, "threshold_max": threshold_max, "histogram": histogram}


def fiji_count_size_threshold_display_macro_text():
    return """args = getArgument();
parts = split(args, "|||");
input = parts[0];
display = parts[1];
manualThresholdMin = parseFloat(parts[2]);
manualThresholdMax = parseFloat(parts[3]);
open(input);
run("8-bit");
setThreshold(manualThresholdMin, manualThresholdMax);
run("Create Selection");
run("RGB Color");
setForegroundColor(255, 0, 0);
run("Fill", "slice");
saveAs("PNG", display);
run("Close All");
"""


def run_fiji_count_size_threshold_display(input_path, display_path, threshold_min, threshold_max):
    if not os.path.isfile(FIJI_EXECUTABLE):
        raise ValueError(f"Fiji executable was not found: {FIJI_EXECUTABLE}")
    os.makedirs(os.path.dirname(display_path), exist_ok=True)
    lower, upper = normalize_threshold_range(threshold_min, threshold_max)
    if lower is None or upper is None:
        lower, upper = 0, 255
    with tempfile.NamedTemporaryFile("w", suffix=".ijm", delete=False, encoding="utf-8") as handle:
        handle.write(fiji_count_size_threshold_display_macro_text())
        macro_path = handle.name
    completed = subprocess.run(
        [FIJI_EXECUTABLE, "--headless", "--console", "-macro", macro_path, "|||".join([input_path, display_path, str(int(lower)), str(int(upper))])],
        text=True,
        capture_output=True,
        check=False,
        timeout=120,
    )
    try:
        os.unlink(macro_path)
    except OSError:
        pass
    if completed.returncode != 0:
        raise RuntimeError((completed.stdout or "") + "\n" + (completed.stderr or ""))
    return display_path


def count_size_thresholds_from_state(state):
    thresholds = state.get("count_size_thresholds", {})
    thresholds = thresholds if isinstance(thresholds, dict) else {}
    path = count_size_thresholds_file_path(state)
    if path and os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
            saved = payload.get("thresholds", payload) if isinstance(payload, dict) else {}
            if isinstance(saved, dict):
                merged = dict(saved)
                merged.update(thresholds)
                thresholds = merged
        except Exception:
            pass
    state["count_size_thresholds"] = thresholds
    return thresholds


def count_size_thresholds_file_path(state):
    raw_output_root = str(state.get("output_root", "") or "").strip()
    if not raw_output_root:
        return ""
    output_root = os.path.abspath(raw_output_root)
    return os.path.join(output_root, COUNT_SIZE_THRESHOLDS_FILENAME)


def save_count_size_thresholds_file(state):
    path = count_size_thresholds_file_path(state)
    if not path:
        raise ValueError("Choose or load an experiment before saving crop thresholds.")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    payload = {
        "version": 1,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "thresholds": state.get("count_size_thresholds", {}),
    }
    tmp_path = path + "." + uuid4().hex + ".tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    return path


def count_size_threshold_key(rel, channel_name):
    portable_rel = str(rel or "").replace("\\", "/")
    return f"{portable_rel}|||{valid_channel(channel_name, 'red')}"


def saved_count_size_threshold(state, rel, channel_name):
    thresholds = count_size_thresholds_from_state(state)
    row = thresholds.get(count_size_threshold_key(rel, channel_name), {})
    if not row:
        legacy_keys = (
            f"{rel}|||{valid_channel(channel_name, 'red')}",
            f"{str(rel or '').replace('/', os.sep)}|||{valid_channel(channel_name, 'red')}",
        )
        row = next((thresholds.get(key, {}) for key in legacy_keys if thresholds.get(key)), {})
    lower, upper = normalize_threshold_range(row.get("min"), row.get("max")) if isinstance(row, dict) else (None, None)
    return lower, upper


def save_count_size_threshold(state, rel, channel_name, threshold_min, threshold_max):
    channel_name = valid_channel(channel_name, "red")
    rel = str(rel or "").strip()
    if not rel or rel not in state.get("items", {}):
        raise ValueError("The selected crop is not loaded in this experiment. Refresh the crop list and try again.")
    lower, upper = normalize_threshold_range(threshold_min, threshold_max)
    if lower is None or upper is None:
        raise ValueError("Threshold must be a valid 0-255 Min/Max range.")
    state.setdefault("count_size_thresholds", {})
    state["count_size_thresholds"][count_size_threshold_key(rel, channel_name)] = {
        "rel": rel.replace("\\", "/"),
        "channel": channel_name,
        "min": lower,
        "max": upper,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    threshold_file = save_count_size_thresholds_file(state)
    save_state(state)
    return lower, upper, threshold_file


def count_size_analysis_source_for_rel(state, rel):
    item = state.get("items", {}).get(rel)
    if not item:
        raise FileNotFoundError(f"Crop is not loaded: {rel}")
    src, crop_choice = selected_crop_path_for_item(state, rel, item)
    if not os.path.isfile(src):
        raise FileNotFoundError(src)
    output_root = ensure_sibling_output_layout(state) or APP_ROOT
    cleaning_params = cleaning_params_for_rel(state, rel)
    channel_cleaning_params = channel_cleaning_params_for_rel(state, rel)
    analysis_cleaning_params = {**cleaning_params, "channel_params": channel_cleaning_params}
    puncta_cleaned_channels = cleaning_channels_for_rel(state, rel)
    raw_preserved_channels = set(CHANNELS) - puncta_cleaned_channels
    base_name = bounded_output_stem("count_size_threshold__" + crop_choice + "__" + rel.replace(os.sep, "__"))
    if use_auto_cleaned_input_for(state, "count_size"):
        cached = current_auto_cleaned_input_for_item_settings(state, rel, item, src)
        if cached:
            return cached
        generated = generate_auto_cleaned_crop_for_item(state, rel, item)
        if generated:
            cached = current_auto_cleaned_input_for_item_settings(state, rel, item, src)
            if cached:
                return cached
    return prepare_analysis_input_crop(
        src,
        output_root,
        base_name,
        use_auto_cleaned_input_for(state, "count_size"),
        preserve_raw_channels=raw_preserved_channels,
        puncta_cleaned_channels=puncta_cleaned_channels,
        cleaning_params=analysis_cleaning_params,
    )


def count_size_channel_for_threshold(state, rel, channel_name):
    channel_name = valid_channel(channel_name, "red")
    analysis_src, cleaning_preview, cleaning_params_file, analysis_input = count_size_analysis_source_for_rel(state, rel)
    image_yxc, layout = read_yxc(analysis_src)
    channel = np.asarray(image_yxc[..., CHANNELS[channel_name]])
    if channel.size == 0 or np.count_nonzero(channel) == 0:
        raise ValueError(f"{channel_name.title()} channel is blank or absent for this crop.")
    return channel, analysis_src, analysis_input, layout


def count_size_threshold_channel_input(state, rel, channel_name):
    channel, analysis_src, analysis_input, layout = count_size_channel_for_threshold(state, rel, channel_name)
    tmpdir = tempfile.mkdtemp(prefix="count_size_threshold_source_")
    channel_input = os.path.join(tmpdir, f"{valid_channel(channel_name, 'red')}_channel_input.tif")
    tiff.imwrite(channel_input, channel)
    return channel, channel_input, tmpdir, analysis_src, analysis_input, layout


def scale_plane_for_png(plane):
    arr = np.asarray(plane, dtype=np.float64)
    if arr.size == 0 or np.nanmax(arr) <= 0:
        return np.zeros(arr.shape, dtype=np.uint8)
    positive = arr[arr > 0]
    if positive.size:
        low, high = np.percentile(positive, [1, 99.8])
    else:
        low, high = np.percentile(arr, [1, 99.8])
    if high <= low:
        high = float(np.nanmax(arr)) or 1.0
        low = float(np.nanmin(arr))
    return np.clip((arr - low) / max(high - low, 1e-9) * 255, 0, 255).astype(np.uint8)


def imagej_8bit_for_preview(plane):
    arr = np.asarray(plane, dtype=np.float64)
    if arr.size == 0:
        return np.zeros(arr.shape, dtype=np.uint8)
    min_value = float(np.nanmin(arr))
    max_value = float(np.nanmax(arr))
    if not np.isfinite(min_value) or not np.isfinite(max_value) or max_value <= min_value:
        return np.zeros(arr.shape, dtype=np.uint8)
    return np.clip((arr - min_value) / (max_value - min_value) * 255, 0, 255).astype(np.uint8)


def colorize_threshold_channel(plane, channel_name):
    gray = scale_plane_for_png(plane)
    rgb = np.zeros((*gray.shape, 3), dtype=np.uint8)
    rgb[..., CHANNELS[valid_channel(channel_name, "red")]] = gray
    return rgb


def png_bytes_from_rgb_array(arr):
    out = io.BytesIO()
    Image.fromarray(np.asarray(arr, dtype=np.uint8), mode="RGB").save(out, format="PNG")
    return out.getvalue()


def fiji_png_as_black_particles_display(png_bytes):
    image = Image.open(io.BytesIO(png_bytes)).convert("L")
    arr = np.asarray(image)
    if arr.size and float(np.mean(arr)) < 127.5:
        arr = 255 - arr
        image = Image.fromarray(arr.astype(np.uint8), mode="L")
    out = io.BytesIO()
    image.convert("RGB").save(out, format="PNG")
    return out.getvalue()


THRESHOLD_PREVIEW_CACHE_VERSION = "fiji-threshold-preview-v1-20261009"
THRESHOLD_PREVIEW_FIJI_LOCK = threading.Lock()


def threshold_preview_fingerprint(channel, channel_name, analysis_input):
    array = np.ascontiguousarray(channel)
    digest = hashlib.sha256()
    digest.update(THRESHOLD_PREVIEW_CACHE_VERSION.encode("utf-8"))
    digest.update(valid_channel(channel_name, "red").encode("utf-8"))
    digest.update(str(analysis_input or "").encode("utf-8"))
    digest.update(str(array.shape).encode("utf-8"))
    digest.update(str(array.dtype).encode("utf-8"))
    digest.update(array.tobytes())
    return digest.hexdigest()


def threshold_preview_cache_path(state, fingerprint, suffix):
    output_root = ensure_sibling_output_layout(state) or APP_ROOT
    cache_root = os.path.join(output_root, ".count_size_threshold_preview_cache")
    os.makedirs(cache_root, exist_ok=True)
    return os.path.join(cache_root, fingerprint + suffix)


def atomic_write_bytes(path, data):
    tmp_path = path + "." + uuid4().hex + ".tmp"
    try:
        with open(tmp_path, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def fast_8bit_threshold_preview(channel, lower, upper):
    gray = imagej_8bit_for_preview(channel)
    rgb = np.repeat(gray[..., np.newaxis], 3, axis=2)
    selected = (gray >= int(lower)) & (gray <= int(upper))
    rgb[selected] = np.array([255, 0, 0], dtype=np.uint8)
    return png_bytes_from_rgb_array(rgb)


def fast_watershed_threshold_preview(channel, lower, upper):
    """Return an interactive watershed approximation without launching Fiji.

    The exact Fiji watershed is still requested after the slider is released and
    is still used for all saved measurements.
    """
    gray = imagej_8bit_for_preview(channel)
    binary = (gray >= int(lower)) & (gray <= int(upper))
    display = np.full(gray.shape, 255, dtype=np.uint8)
    if not np.any(binary):
        return png_bytes_from_rgb_array(np.repeat(display[..., np.newaxis], 3, axis=2))

    distance = ndi.distance_transform_edt(binary)
    # Label connected local-maximum plateaus instead of making every plateau
    # pixel a marker. This remains responsive even for a nearly uniform crop.
    local_maxima = (distance == ndi.maximum_filter(distance, size=3)) & binary
    markers, marker_count = ndi.label(local_maxima)
    if marker_count == 0:
        markers, _ = ndi.label(binary)
    separated = watershed(-distance, markers, mask=binary, watershed_line=True) > 0
    display[separated] = 0
    return png_bytes_from_rgb_array(np.repeat(display[..., np.newaxis], 3, axis=2))


def count_size_threshold_preview_png(
    state, rel, channel_name, threshold_min, threshold_max, view, fast_watershed=False
):
    channel_name = valid_channel(channel_name, "red")
    lower, upper = normalize_threshold_range(threshold_min, threshold_max)
    if lower is None or upper is None:
        lower, upper = 0, 255
    channel, analysis_src, analysis_input, layout = count_size_channel_for_threshold(state, rel, channel_name)
    if view == "original":
        return png_bytes_from_rgb_array(colorize_threshold_channel(channel, channel_name))
    if view == "8bit":
        return fast_8bit_threshold_preview(channel, lower, upper)
    if view != "watershed":
        raise ValueError(f"Unknown threshold preview view: {view}")
    if fast_watershed:
        return fast_watershed_threshold_preview(channel, lower, upper)

    fingerprint = threshold_preview_fingerprint(channel, channel_name, analysis_input)
    cache_path = threshold_preview_cache_path(state, fingerprint, f"__watershed_{int(lower)}_{int(upper)}.png")
    if os.path.isfile(cache_path):
        with open(cache_path, "rb") as handle:
            return handle.read()

    with THRESHOLD_PREVIEW_FIJI_LOCK:
        if os.path.isfile(cache_path):
            with open(cache_path, "rb") as handle:
                return handle.read()
        tmpdir = tempfile.mkdtemp(prefix="count_size_threshold_watershed_")
        try:
            channel_input = os.path.join(tmpdir, f"{channel_name}_channel_input.tif")
            results_path = os.path.join(tmpdir, "particles.csv")
            summary_path = os.path.join(tmpdir, "summary.csv")
            mask_path = os.path.join(tmpdir, "mask.tif")
            display_path = os.path.join(tmpdir, "watershed_preview.png")
            tiff.imwrite(channel_input, channel)
            run_fiji_count_size_macro(
                channel_input, results_path, summary_path, mask_path,
                threshold_min=lower, threshold_max=upper, display_path=display_path,
            )
            with open(display_path, "rb") as handle:
                png_data = fiji_png_as_black_particles_display(handle.read())
            atomic_write_bytes(cache_path, png_data)
            return png_data
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


def count_size_threshold_default_json(state, rel, channel_name):
    channel_name = valid_channel(channel_name, "red")
    saved_min, saved_max = saved_count_size_threshold(state, rel, channel_name)
    channel, analysis_src, analysis_input, layout = count_size_channel_for_threshold(state, rel, channel_name)
    fingerprint = threshold_preview_fingerprint(channel, channel_name, analysis_input)
    cache_path = threshold_preview_cache_path(state, fingerprint, "__fiji_default_threshold.json")
    result = None
    cache_hit = False
    try:
        if os.path.isfile(cache_path):
            with open(cache_path, "r", encoding="utf-8") as handle:
                result = json.load(handle)
            cache_hit = True
        if result is None:
            with THRESHOLD_PREVIEW_FIJI_LOCK:
                if os.path.isfile(cache_path):
                    with open(cache_path, "r", encoding="utf-8") as handle:
                        result = json.load(handle)
                    cache_hit = True
                else:
                    tmpdir = tempfile.mkdtemp(prefix="count_size_threshold_probe_")
                    try:
                        channel_input = os.path.join(tmpdir, f"{channel_name}_channel_input.tif")
                        tiff.imwrite(channel_input, channel)
                        result = run_fiji_count_size_threshold_probe(channel_input)
                        tmp_path = cache_path + "." + uuid4().hex + ".tmp"
                        try:
                            with open(tmp_path, "w", encoding="utf-8") as handle:
                                json.dump(result, handle)
                                handle.flush()
                                os.fsync(handle.fileno())
                            os.replace(tmp_path, cache_path)
                        finally:
                            if os.path.exists(tmp_path):
                                os.unlink(tmp_path)
                    finally:
                        shutil.rmtree(tmpdir, ignore_errors=True)
    except Exception as error:
        result = {
            "threshold_min": saved_min if saved_min is not None else 0,
            "threshold_max": saved_max if saved_max is not None else 255,
            "histogram": [],
            "probe_error": str(error),
        }
    result.update({
        "saved_min": saved_min,
        "saved_max": saved_max,
        "analysis_input_file": analysis_src,
        "analysis_input_mode": analysis_input,
        "preview_cache_hit": cache_hit,
    })
    return result


def build_count_size_channel_summary(result_df, status_df, channel_labels):
    rows = []
    for channel_name in CHANNELS:
        channel_results = result_df[result_df["Channel_Key"] == channel_name] if not result_df.empty and "Channel_Key" in result_df else pd.DataFrame()
        channel_status = status_df[status_df["Channel_Key"] == channel_name] if not status_df.empty and "Channel_Key" in status_df else pd.DataFrame()
        count_values = pd.to_numeric(channel_results.get("Condensate_Count", pd.Series(dtype=float)), errors="coerce").dropna()
        avg_area_px = pd.to_numeric(channel_results.get("Average_Area_px2", pd.Series(dtype=float)), errors="coerce").dropna()
        avg_area_um = pd.to_numeric(channel_results.get("Average_Area_um2", pd.Series(dtype=float)), errors="coerce").dropna()
        avg_diameter_um = pd.to_numeric(channel_results.get("Average_Equivalent_Diameter_um", pd.Series(dtype=float)), errors="coerce").dropna()
        total_area_px = pd.to_numeric(channel_results.get("Total_Area_px2", pd.Series(dtype=float)), errors="coerce").dropna()
        total_area_um = pd.to_numeric(channel_results.get("Total_Area_um2", pd.Series(dtype=float)), errors="coerce").dropna()
        minimum_areas = pd.to_numeric(channel_results.get("Minimum_Particle_Area_px2", pd.Series(dtype=float)), errors="coerce").dropna()
        size_ranges = channel_results.get("Particle_Size_Range_px2", pd.Series(dtype=str)).dropna()
        processed = channel_status[channel_status["Status"] == "Processed"] if not channel_status.empty and "Status" in channel_status else pd.DataFrame()
        rows.append({
            "Channel": channel_name.title(),
            "Channel_Key": channel_name,
            "Channel_Label": channel_labels.get(channel_name, channel_name),
            "Processed_Crop_Channel_Rows": int(len(processed)),
            "Total_Condensate_Count": int(count_values.sum()) if not count_values.empty else 0,
            "Mean_Count_Per_Crop": float(count_values.mean()) if not count_values.empty else 0.0,
            "Mean_Average_Area_px2_Per_Crop": float(avg_area_px.mean()) if not avg_area_px.empty else 0.0,
            "Mean_Average_Area_um2_Per_Crop": float(avg_area_um.mean()) if not avg_area_um.empty else 0.0,
            "Mean_Average_Equivalent_Diameter_um_Per_Crop": float(avg_diameter_um.mean()) if not avg_diameter_um.empty else 0.0,
            "Total_Area_px2_All_Crops": float(total_area_px.sum()) if not total_area_px.empty else 0.0,
            "Total_Area_um2_All_Crops": float(total_area_um.sum()) if not total_area_um.empty else 0.0,
            "Minimum_Particle_Area_px2": float(minimum_areas.iloc[0]) if not minimum_areas.empty else 0.0,
            "Particle_Size_Range_px2": str(size_ranges.iloc[0]) if not size_ranges.empty else "",
        })
    return pd.DataFrame(rows)



def run_count_size_analysis_on_kept_crops(state):
    ensure_sibling_output_layout(state)
    state["last_workflow_action"] = "run_count_size_analysis"
    state["last_workflow_note"] = "Count/size analysis was run because the count/size analysis button was clicked."
    crop_root = state["crop_root"]
    output_root = state["output_root"]
    os.makedirs(output_root, exist_ok=True)
    approved_root = os.path.join(output_root, "approved_crops_for_condensate_count_size")
    count_root = os.path.join(output_root, "real_fiji_condensate_count_size_analysis")
    channel_input_root = os.path.join(count_root, "channel_inputs_sent_to_fiji")
    mask_root = os.path.join(count_root, "real_fiji_binary_watershed_masks")
    table_root = os.path.join(count_root, "real_fiji_tables")
    for folder in (approved_root, channel_input_root, mask_root, table_root):
        os.makedirs(folder, exist_ok=True)
    selected_channels = count_size_channels_from_state(state)
    if not selected_channels:
        raise ValueError("Choose at least one channel for condensate count/size.")
    channel_labels = pcc_channel_labels_from_state(state)
    use_cleaned = use_auto_cleaned_input_for(state, "count_size")
    pixel_um = float(state.get("pixel_um", 0.1290271))
    minimum_particle_area_px2, particle_size_range_px2 = normalize_count_size_particle_size(
        state.get("count_size_particle_size", "10-Infinity")
    )
    state["count_size_particle_size"] = particle_size_range_px2
    rows, status_rows, approved_rows, error_rows, jobs = [], [], [], [], []
    kept_items = [(rel, item) for rel, item in sorted(state.get("items", {}).items()) if item.get("status", "keep") == "keep"]
    set_analysis_progress(True, "preparing", 0, len(kept_items), "Preparing kept crops and condensate channels for one Fiji batch.")
    for crop_number, (rel, item) in enumerate(kept_items, start=1):
        if item.get("status", "keep") != "keep":
            continue
        channel_roles = channel_roles_for_item(state, item)
        channels_for_crop = [channel for channel in selected_channels if channel_roles.get(channel) == "condensate"]
        src, crop_choice = selected_crop_path_for_item(state, rel, item)
        if not os.path.exists(src):
            continue
        cleaning_params = cleaning_params_for_rel(state, rel)
        channel_cleaning_params = channel_cleaning_params_for_rel(state, rel)
        analysis_cleaning_params = {**cleaning_params, "channel_params": channel_cleaning_params}
        puncta_cleaned_channels = cleaning_channels_for_rel(state, rel)
        raw_preserved_channels = set(CHANNELS) - puncta_cleaned_channels
        dst_name = safe_name(crop_choice + "__" + rel.replace(os.sep, "__")) + ".tiff"
        dst = os.path.join(approved_root, dst_name)
        copy2_unless_same(src, dst)
        base_name = os.path.splitext(dst_name)[0]
        analysis_src = ""
        try:
            analysis_src, cleaning_preview, cleaning_params_file, analysis_input = analysis_input_for_reviewed_crop(
                state, rel, item, dst, output_root, base_name, use_cleaned
            )
            image_yxc, layout = read_yxc(analysis_src)
            approved_rows.append({
                "Status": "Included", "Crop_Choice": crop_choice, "Original_Crop_File": src,
                "Approved_Crop_File": dst, "Analysis_Input_File": analysis_src,
                "Analysis_Input_Mode": analysis_input, "Channel_Cleaning_Preview_PNG": cleaning_preview,
                "Review_Relative_Path": rel, "Layout_Read": layout,
                "Red_Channel_Role_For_Crop": channel_roles.get("red", "ignore"),
                "Green_Channel_Role_For_Crop": channel_roles.get("green", "ignore"),
                "Blue_Channel_Role_For_Crop": channel_roles.get("blue", "ignore"),
            })
            for channel_name in selected_channels:
                channel_label = channel_labels.get(channel_name, channel_name)
                if channel_name not in channels_for_crop:
                    status_rows.append({
                        "Review_Relative_Path": rel, "Channel": channel_name.title(), "Channel_Key": channel_name,
                        "Channel_Label": channel_label, "Channel_Role_For_Crop": channel_roles.get(channel_name, "ignore"),
                        "Status": "Skipped_Not_Condensate_For_This_Crop", "Analysis_Input_File": analysis_src,
                        "Fiji_Channel_Input": "", "Error": "Per-crop role is not condensate.",
                    })
                    continue
                channel = np.asarray(image_yxc[..., CHANNELS[channel_name]])
                safe_base = safe_name(rel)
                channel_input = os.path.join(channel_input_root, f"{safe_base}_{channel_name}_count_size_input.tif")
                mask_path = os.path.join(mask_root, f"{safe_base}_{channel_name}_count_size_mask.tif")
                results_path = os.path.join(table_root, f"{safe_base}_{channel_name}_particles.csv")
                summary_path = os.path.join(table_root, f"{safe_base}_{channel_name}_summary.csv")
                tiff.imwrite(channel_input, channel)
                if channel.size == 0 or np.count_nonzero(channel) == 0:
                    status_rows.append({
                        "Review_Relative_Path": rel, "Channel": channel_name.title(), "Channel_Key": channel_name,
                        "Channel_Label": channel_label, "Channel_Role_For_Crop": channel_roles.get(channel_name, "ignore"),
                        "Status": "Skipped_Blank_or_Absent", "Analysis_Input_File": analysis_src,
                        "Fiji_Channel_Input": channel_input, "Error": "Selected channel was blank or absent.",
                    })
                    continue
                threshold_min, threshold_max = saved_count_size_threshold(state, rel, channel_name)
                threshold_source = "Saved_Manual_Threshold_Range_After_8bit" if threshold_min is not None and threshold_max is not None else "Fiji_Default_Dark_After_8bit"
                job_id = f"{rel}|||{channel_name}"
                jobs.append({
                    "job_id": job_id, "rel": rel, "src": src, "dst": dst, "crop_choice": crop_choice,
                    "channel_name": channel_name, "channel_label": channel_label, "channel_roles": channel_roles,
                    "analysis_src": analysis_src, "analysis_input": analysis_input, "cleaning_preview": cleaning_preview,
                    "input_path": channel_input, "mask_path": mask_path, "results_path": results_path,
                    "summary_path": summary_path, "cache_path": summary_path + ".cache.json",
                    "min_area": minimum_particle_area_px2,
                    "particle_size_range_px2": particle_size_range_px2,
                    "threshold_min": threshold_min, "threshold_max": threshold_max, "threshold_source": threshold_source,
                    "fingerprint": count_size_array_fingerprint(channel, minimum_particle_area_px2, threshold_min, threshold_max),
                })
        except Exception as error:
            error_rows.append({
                "Review_Relative_Path": rel, "Approved_Crop_File": dst,
                "Analysis_Input_File": analysis_src, "Error": str(error),
            })
        set_analysis_progress(True, "preparing", crop_number, len(kept_items), f"Prepared crop {crop_number} of {len(kept_items)}.")

    results_by_job = {}
    pending_jobs = []
    for job in jobs:
        cached = load_cached_count_size_job(job)
        if cached is None:
            pending_jobs.append(job)
        else:
            results_by_job[job["job_id"]] = cached
    cache_hits = len(jobs) - len(pending_jobs)
    if pending_jobs:
        set_analysis_progress(True, "fiji_batch", cache_hits, len(jobs), f"Starting one Fiji batch for {len(pending_jobs)} changed crop-channel jobs; {cache_hits} reused from cache.")
        try:
            batch_results = run_fiji_count_size_batch(
                pending_jobs,
                progress_callback=lambda done, total: set_analysis_progress(
                    True, "fiji_batch", cache_hits + done, len(jobs),
                    f"Fiji batch processed {done} of {total} changed jobs; {cache_hits} cached."
                ),
            )
            results_by_job.update(batch_results)
        except Exception as batch_error:
            # Preserve correctness and recoverability if a Fiji installation rejects the batch macro.
            set_analysis_progress(True, "safe_fallback", cache_hits, len(jobs), "Batch failed; using the original single-job Fiji fallback.")
            for fallback_number, job in enumerate(pending_jobs, start=1):
                try:
                    single = run_fiji_count_size_macro(
                        job["input_path"], job["results_path"], job["summary_path"], job["mask_path"],
                        min_area=job["min_area"], threshold_min=job["threshold_min"], threshold_max=job["threshold_max"],
                    )
                    result = {"summary": single["summary"], "particles": single.get("particles", []), "cached": False}
                    save_count_size_job_cache(job)
                    results_by_job[job["job_id"]] = result
                except Exception as error:
                    error_rows.append({
                        "Review_Relative_Path": job["rel"], "Approved_Crop_File": job["dst"],
                        "Analysis_Input_File": job["analysis_src"],
                        "Error": f"Batch error: {batch_error}\nFallback error: {error}",
                    })
                set_analysis_progress(True, "safe_fallback", cache_hits + fallback_number, len(jobs), f"Fallback processed {fallback_number} of {len(pending_jobs)} changed jobs.")
    else:
        set_analysis_progress(True, "cache", len(jobs), len(jobs), f"Reused all {len(jobs)} unchanged Fiji results from cache.")

    for job in jobs:
        result = results_by_job.get(job["job_id"])
        if result is None:
            status_rows.append({
                "Review_Relative_Path": job["rel"], "Channel": job["channel_name"].title(),
                "Channel_Key": job["channel_name"], "Channel_Label": job["channel_label"],
                "Channel_Role_For_Crop": job["channel_roles"].get(job["channel_name"], "ignore"),
                "Status": "Error", "Analysis_Input_File": job["analysis_src"],
                "Fiji_Channel_Input": job["input_path"], "Error": "Fiji did not produce a result for this job.",
            })
            continue
        summary = result["summary"]
        particles = result.get("particles", [])
        particle_df = pd.DataFrame(particles)
        if not particle_df.empty and "Area_px2" in particle_df:
            particle_area_px2 = pd.to_numeric(particle_df["Area_px2"], errors="coerce")
            particle_df["Area_um2"] = particle_area_px2 * (pixel_um ** 2)
            particle_df["Equivalent_Diameter_um"] = 2.0 * np.sqrt(particle_area_px2 / np.pi) * pixel_um
            particle_df.to_csv(job["results_path"], index=False)
            diameters = pd.to_numeric(particle_df["Equivalent_Diameter_um"], errors="coerce").dropna()
            average_equivalent_diameter_um = float(diameters.mean()) if not diameters.empty else 0.0
        else:
            average_equivalent_diameter_um = 0.0
        average_area_px2 = float(summary["Average_Area_px2"])
        total_area_px2 = float(summary["Total_Area_px2"])
        rows.append({
            "Image": os.path.basename(job["src"]), "Review_Relative_Path": job["rel"],
            "Channel": job["channel_name"].title(), "Channel_Key": job["channel_name"],
            "Channel_Label": job["channel_label"],
            "Channel_Role": job["channel_roles"].get(job["channel_name"], "ignore"),
            "Channel_Role_For_Crop": job["channel_roles"].get(job["channel_name"], "ignore"),
            "Red_Channel_Role_For_Crop": job["channel_roles"].get("red", "ignore"),
            "Green_Channel_Role_For_Crop": job["channel_roles"].get("green", "ignore"),
            "Blue_Channel_Role_For_Crop": job["channel_roles"].get("blue", "ignore"),
            "Condensate_Count": int(summary["Condensate_Count"]),
            "Average_Equivalent_Diameter_um": average_equivalent_diameter_um,
            "Average_Area_px2": average_area_px2, "Average_Area_um2": average_area_px2 * (pixel_um ** 2),
            "Equivalent_Diameter_From_Average_Area_um": 2.0 * float(np.sqrt(average_area_px2 / np.pi)) * pixel_um if average_area_px2 > 0 else 0.0,
            "Total_Area_px2": total_area_px2, "Total_Area_um2": total_area_px2 * (pixel_um ** 2),
            "Percent_Area": float(summary["Percent_Area"]), "Pixel_Size_um": pixel_um,
            "Minimum_Particle_Area_px2": job["min_area"],
            "Particle_Size_Range_px2": job["particle_size_range_px2"],
            "Measurement_Source": "Real_Fiji_Analyze_Particles_On_Selected_Channel",
            "Threshold_8bit_Min": job["threshold_min"] if job["threshold_min"] is not None else "Fiji Default",
            "Threshold_8bit_Max": job["threshold_max"] if job["threshold_max"] is not None else "Fiji Default",
            "Threshold_Source": job["threshold_source"],
            "Watershed_Source": "Real_Fiji_run_Watershed_after_Convert_to_Mask",
            "Approved_Crop_File": job["dst"], "Analysis_Input_File": job["analysis_src"],
            "Analysis_Input_Mode": job["analysis_input"], "Channel_Cleaning_Preview_PNG": job["cleaning_preview"],
            "Fiji_Channel_Input": job["input_path"], "Fiji_Mask_File": job["mask_path"],
            "Fiji_Particle_Table": job["results_path"], "Fiji_Summary_Table": job["summary_path"],
            "Result_Cache": "Reused" if result.get("cached") else "Fresh Fiji batch",
        })
        status_rows.append({
            "Review_Relative_Path": job["rel"], "Channel": job["channel_name"].title(),
            "Channel_Key": job["channel_name"], "Channel_Label": job["channel_label"],
            "Channel_Role_For_Crop": job["channel_roles"].get(job["channel_name"], "ignore"),
            "Status": "Processed",
            "Analysis_Input_File": job["analysis_src"], "Fiji_Channel_Input": job["input_path"], "Error": "",
        })

    set_analysis_progress(True, "workbook", len(jobs), len(jobs), "Writing the count/size workbook.")
    result_df = pd.DataFrame(rows)
    status_df = pd.DataFrame(status_rows)
    approved_df = pd.DataFrame(approved_rows)
    errors_df = pd.DataFrame(error_rows)
    channel_summary_df = build_count_size_channel_summary(result_df, status_df, channel_labels)
    output_excel = os.path.join(count_root, "Real_Fiji_Condensate_Count_Size_Reviewed_ROIs.xlsx")
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        approved_df.to_excel(writer, sheet_name="Approved Crops", index=False)
        result_df.to_excel(writer, sheet_name="Count Size Results", index=False)
        channel_summary_df.to_excel(writer, sheet_name="Channel Summary", index=False)
        channel_label_rows_from_state(state).to_excel(writer, sheet_name="Channel Labels", index=False)
        status_df.to_excel(writer, sheet_name="Status", index=False)
        errors_df.to_excel(writer, sheet_name="Analysis Errors", index=False)
    style_workbook(output_excel)
    state["last_count_size_excel"] = output_excel
    state["last_count_size_approved_root"] = approved_root
    state["last_count_size_output_root"] = count_root
    state["last_count_size_cache_hits"] = cache_hits
    state["last_count_size_fresh_jobs"] = len(pending_jobs)
    save_state(state)
    set_analysis_progress(False, "complete", len(jobs), len(jobs), f"Count/size complete: {len(jobs)} crop-channel jobs, {cache_hits} cached, {len(pending_jobs)} changed.")
    return output_excel, approved_root, count_root, len(approved_rows), len(result_df)


def run_pcc_manders_analysis_on_kept_crops(state):
    ensure_sibling_output_layout(state)
    state["last_workflow_action"] = "run_pcc_manders_analysis"
    state["last_workflow_note"] = "PCC/Manders analysis was run because the PCC/Manders analysis button was clicked."
    crop_root = state["crop_root"]
    output_root = state["output_root"]
    mode = valid_analysis_mode(state.get("analysis_mode"), "red_green")
    state["analysis_mode"] = mode
    state["stage3_mode"] = "dapi_red_green" if mode == "red_green" else "no_dapi_three_color"
    save_state(state)
    channel_labels = pcc_channel_labels_from_state(state)
    channel_labels_df = channel_label_rows_from_state(state)
    use_cleaned = use_auto_cleaned_input_for(state, "pcc")
    state["pcc_pair_mode"] = valid_pcc_pair_mode(state.get("pcc_pair_mode"), "red_green")
    os.makedirs(output_root, exist_ok=True)
    approved_root = os.path.join(output_root, "approved_crops_for_real_fiji_pcc_manders")
    pcc_root = os.path.join(output_root, "real_fiji_pcc_manders_analysis")
    final_preview_root = os.path.join(pcc_root, "preview_images_sent_to_fiji")
    rgb_input_root = os.path.join(pcc_root, "rgb_crops_sent_to_fiji")
    log_root = os.path.join(pcc_root, "fiji_jacop_logs")
    os.makedirs(approved_root, exist_ok=True)
    os.makedirs(pcc_root, exist_ok=True)
    os.makedirs(final_preview_root, exist_ok=True)
    os.makedirs(rgb_input_root, exist_ok=True)
    os.makedirs(log_root, exist_ok=True)

    rows = []
    approved_rows = []
    error_rows = []
    kept_items = [(rel, item) for rel, item in sorted(state.get("items", {}).items()) if item.get("status", "keep") == "keep"]
    set_analysis_progress(True, "pcc_manders", 0, len(kept_items), f"Starting Fiji PCC/Manders for {len(kept_items)} kept crops.")
    for crop_number, (rel, item) in enumerate(kept_items, start=1):
        channel_roles = channel_roles_for_item(state, item)
        pair_specs = pcc_pair_specs_for_item(state, item)
        src, crop_choice = selected_crop_path_for_item(state, rel, item)
        if not os.path.exists(src):
            continue
        cleaning_params = cleaning_params_for_rel(state, rel)
        channel_cleaning_params = channel_cleaning_params_for_rel(state, rel)
        analysis_cleaning_params = {**cleaning_params, "channel_params": channel_cleaning_params}
        puncta_cleaned_channels = cleaning_channels_for_rel(state, rel)
        raw_preserved_channels = set(CHANNELS) - puncta_cleaned_channels
        dst_name = bounded_output_stem(crop_choice + "__" + rel.replace(os.sep, "__")) + ".tiff"
        dst = os.path.join(approved_root, dst_name)
        copy2_unless_same(src, dst)
        preview_name = os.path.splitext(dst_name)[0] + "_sent_to_fiji_preview.png"
        preview_path = os.path.join(final_preview_root, preview_name)
        base_name = os.path.splitext(dst_name)[0]
        analysis_src, cleaning_preview, cleaning_params_file, analysis_input = analysis_input_for_reviewed_crop(
            state, rel, item, dst, output_root, base_name, use_cleaned
        )
        try:
            image_yxc, layout = read_yxc(analysis_src)
            pcc_image_yxc = pcc_image_yxc_for_selected_pair(state, image_yxc)
            save_display_png_for_rgb(display_rgb(pcc_image_yxc), preview_path)
            rgb_input_path = os.path.join(rgb_input_root, base_name + "_rgb_crop_sent_to_fiji.tif")
            tiff.imwrite(rgb_input_path, pcc_image_yxc)
            image_name = os.path.basename(src)
            expected_pairs = ", ".join(pair for pair, _a, _b, _pair_type in pair_specs) or "No PCC/Manders pair selected for this crop"
            if not pair_specs:
                error_rows.append({
                    "Review_Relative_Path": rel,
                    "Approved_Crop_File": dst,
                    "Analysis_Input_File": analysis_src,
                    "Analysis_Input_Mode": analysis_input,
                    "Channel_Cleaning_Preview_PNG": cleaning_preview,
                    "Error": "Skipped PCC/Manders because this crop does not have a valid condensate-marker pair or at least two condensate channels.",
                    "Red_Channel_Role_For_Crop": channel_roles.get("red", "ignore"),
                    "Green_Channel_Role_For_Crop": channel_roles.get("green", "ignore"),
                    "Blue_Channel_Role_For_Crop": channel_roles.get("blue", "ignore"),
                })
            for pair_name, channel_a, channel_b, pair_type in pair_specs:
                log_path = os.path.join(log_root, f"{base_name}_{pair_name}_fiji_log.txt")
                row = run_fiji_pcc_manders_pair(rgb_input_path, pair_name, channel_a, channel_b, log_path)
                row.update(pcc_pair_interpretation_columns(channel_a, channel_b, channel_labels=channel_labels))
                row.update(manders_interpretation_columns(channel_a, channel_b, channel_labels=channel_labels))
                channel_a_label = channel_labels.get(channel_a, channel_a)
                channel_b_label = channel_labels.get(channel_b, channel_b)
                channel_a_role = channel_roles.get(channel_a, "ignore")
                channel_b_role = channel_roles.get(channel_b, "ignore")
                pair_label = pcc_manders_pair_label(channel_a_label, channel_b_label, channel_a_role, channel_b_role, pair_type)
                condensate_label = channel_a_label if pair_type == "condensate_marker" else ""
                marker_label = channel_b_label if pair_type == "condensate_marker" else ""
                row.update({
                    "Image": image_name,
                    "Crop_File": dst,
                    "Analysis_Input_File": analysis_src,
                    "Analysis_Input_Mode": analysis_input,
                    "Channel_Cleaning_Preview_PNG": cleaning_preview,
                    "Review_Relative_Path": rel,
                    "Fiji_RGB_Crop_Input": rgb_input_path,
                    "Fiji_Channel_A": channel_a,
                    "Fiji_Channel_B": channel_b,
                    "Fiji_Channel_A_Label": channel_a_label,
                    "Fiji_Channel_B_Label": channel_b_label,
                    "Fiji_Channel_A_Role": channel_a_role,
                    "Fiji_Channel_B_Role": channel_b_role,
                    "PCC_Manders_Pair_Label": pair_label,
                    "PCC_Manders_Pair_Type": pair_type,
                    "Channel_A_Role_For_Crop": channel_a_role,
                    "Channel_B_Role_For_Crop": channel_b_role,
                    "Condensate_Channel_For_Crop": channel_a if pair_type == "condensate_marker" else "",
                    "Marker_Channel_For_Crop": channel_b if pair_type == "condensate_marker" else "",
                    "Condensate_Channel_Label_For_Crop": condensate_label,
                    "Marker_Channel_Label_For_Crop": marker_label,
                    "Condensate_Channel_A_For_Crop": channel_a if pair_type == "condensate_condensate" else "",
                    "Condensate_Channel_B_For_Crop": channel_b if pair_type == "condensate_condensate" else "",
                    "Condensate_Channel_A_Label_For_Crop": channel_a_label if pair_type == "condensate_condensate" else "",
                    "Condensate_Channel_B_Label_For_Crop": channel_b_label if pair_type == "condensate_condensate" else "",
                    "Manders_Condensate_Overlap_Marker_original": row.get("M1_original") if pair_type == "condensate_marker" else "",
                    "Manders_Marker_Overlap_Condensate_original": row.get("M2_original") if pair_type == "condensate_marker" else "",
                    "Manders_Condensate_Overlap_Marker_thresholded": row.get("M1_thresholded") if pair_type == "condensate_marker" else "",
                    "Manders_Marker_Overlap_Condensate_thresholded": row.get("M2_thresholded") if pair_type == "condensate_marker" else "",
                    "Manders_Channel_A_Overlap_Channel_B_original": row.get("M1_original"),
                    "Manders_Channel_B_Overlap_Channel_A_original": row.get("M2_original"),
                    "Manders_Channel_A_Overlap_Channel_B_thresholded": row.get("M1_thresholded"),
                    "Manders_Channel_B_Overlap_Channel_A_thresholded": row.get("M2_thresholded"),
                    "Red_Channel_Role_For_Crop": channel_roles.get("red", "ignore"),
                    "Green_Channel_Role_For_Crop": channel_roles.get("green", "ignore"),
                    "Blue_Channel_Role_For_Crop": channel_roles.get("blue", "ignore"),
                    "Fiji_Log": log_path,
                    "Analysis_Mode": mode,
                    "PCC_Manders_Pair_Mode": "per_crop_channel_roles",
                    "Cleaning_Radius_px": cleaning_params["radius_px"],
                    "Cleaning_Sensitivity": cleaning_params["sensitivity"],
                    "Cleaning_Min_Area_px": cleaning_params["min_area_px"],
                    "Cleaning_Amount_Percent": cleaning_params["amount"],
                    "Cleaned_Channels": ", ".join(sorted(puncta_cleaned_channels)),
                    "Cleaning_Settings_Source": "per-crop manual adjustment" if has_per_crop_cleaning_params(state, rel) else "global settings",
                })
                rows.append(row)
            approved_rows.append({
                "Status": "Included",
                "Crop_Choice": crop_choice,
                "Original_Crop_File": src,
                "Approved_Crop_File": dst,
                "Analysis_Input_File": analysis_src,
                "Analysis_Input_Mode": analysis_input,
                "Channel_Cleaning_Preview_PNG": cleaning_preview,
                "Final_Analysis_Preview_PNG": preview_path,
                "Review_Relative_Path": rel,
                "Layout_Read": layout,
                "Fiji_RGB_Input_Folder": rgb_input_root,
                "Analysis_Mode": mode,
                "Expected_Fiji_Pairs": expected_pairs,
                "Red_Channel_Label": channel_labels.get("red", "red"),
                "Green_Channel_Label": channel_labels.get("green", "green"),
                "Blue_Channel_Label": channel_labels.get("blue", "blue"),
                "Red_Channel_Role": channel_roles.get("red", "ignore"),
                "Green_Channel_Role": channel_roles.get("green", "ignore"),
                "Blue_Channel_Role": channel_roles.get("blue", "ignore"),
                "Red_Channel_Role_For_Crop": channel_roles.get("red", "ignore"),
                "Green_Channel_Role_For_Crop": channel_roles.get("green", "ignore"),
                "Blue_Channel_Role_For_Crop": channel_roles.get("blue", "ignore"),
                "Expected_Fiji_Pairs_From_Per_Crop_Roles": expected_pairs,
                "PCC_Manders_Pair_Mode": "per_crop_channel_roles",
                "Cleaning_Radius_px": cleaning_params["radius_px"],
                "Cleaning_Sensitivity": cleaning_params["sensitivity"],
                "Cleaning_Min_Area_px": cleaning_params["min_area_px"],
                "Cleaning_Amount_Percent": cleaning_params["amount"],
                "Cleaned_Channels": ", ".join(sorted(puncta_cleaned_channels)),
                "Cleaning_Settings_Source": "per-crop manual adjustment" if has_per_crop_cleaning_params(state, rel) else "global settings",
            })
        except Exception as error:
            error_rows.append({"Review_Relative_Path": rel, "Approved_Crop_File": dst, "Analysis_Input_File": analysis_src, "Analysis_Input_Mode": analysis_input, "Channel_Cleaning_Preview_PNG": cleaning_preview, "Error": str(error)})
        set_analysis_progress(True, "pcc_manders", crop_number, len(kept_items), f"PCC/Manders processed crop {crop_number} of {len(kept_items)}.")

    set_analysis_progress(True, "workbook", len(kept_items), len(kept_items), "Writing the PCC/Manders workbook.")
    raw_df = reorder_pcc_manders_columns(pd.DataFrame(rows))
    review_df = pd.DataFrame([
        {
            "Status": item.get("status", "keep"),
            "Crop_Choice": item.get("crop_choice", "raw"),
            "Selected_Crop_File": selected_crop_path_for_item(state, rel, item)[0],
            "Original_Review_Item": item.get("path", os.path.join(crop_root, rel)),
            "Red_Channel_Role_For_Crop": channel_roles_for_item(state, item).get("red", "ignore"),
            "Green_Channel_Role_For_Crop": channel_roles_for_item(state, item).get("green", "ignore"),
            "Blue_Channel_Role_For_Crop": channel_roles_for_item(state, item).get("blue", "ignore"),
            "Note": item.get("note", ""),
        }
        for rel, item in sorted(state.get("items", {}).items())
    ])
    approved_df = pd.DataFrame(approved_rows)
    summary_df = reorder_summary_columns(build_summary(raw_df))
    per_cell_role_df = build_per_cell_channel_role_summary(state, channel_labels)
    overall_role_df = build_overall_channel_role_summary(state, channel_labels)
    errors_df = pd.DataFrame(error_rows)
    output_excel = os.path.join(pcc_root, "Real_Fiji_PCC_Manders_Reviewed_ROIs.xlsx")
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        review_df.to_excel(writer, sheet_name="Review Decisions", index=False)
        approved_df.to_excel(writer, sheet_name="Approved Crops", index=False)
        raw_df.to_excel(writer, sheet_name="Real Fiji PCC Manders", index=False)
        summary_df.to_excel(writer, sheet_name="Average SEM", index=False)
        per_cell_role_df.to_excel(writer, sheet_name="Per Cell Role Summary", index=False)
        overall_role_df.to_excel(writer, sheet_name="Overall Role Summary", index=False)
        channel_labels_df.to_excel(writer, sheet_name="Channel Labels", index=False)
        errors_df.to_excel(writer, sheet_name="Analysis Errors", index=False)

    style_workbook(output_excel)
    state["final_preview_root"] = final_preview_root
    state["last_final_excel"] = output_excel
    state["last_real_fiji_pcc_manders_excel"] = output_excel
    state["last_pcc_manders_preview_root"] = final_preview_root
    state["last_pcc_manders_approved_root"] = approved_root
    save_state(state)
    if kept_items and raw_df.empty:
        first_errors = [str(row.get("Error", "")) for row in error_rows[:3] if row.get("Error")]
        error_detail = " | ".join(first_errors) or "No valid PCC/Manders channel pair was produced."
        failure_message = (
            f"PCC/Manders FAILED: 0 result rows from {len(kept_items)} kept crops. "
            f"The Analysis Errors worksheet was written to {output_excel}. First error(s): {error_detail}"
        )
        set_analysis_progress(False, "failed", 0, len(kept_items), failure_message)
        raise RuntimeError(failure_message)
    set_analysis_progress(True, "pcc_manders_complete", len(kept_items), len(kept_items), "PCC/Manders finished; finalizing the response.")
    return output_excel, approved_root, final_preview_root, len(approved_rows), len(raw_df)


def preview_channel_cleaning_on_kept_crops(state, force_global=False, max_count=None):
    refresh_derived_paths(state)
    ensure_sibling_output_layout(state)
    state["last_workflow_action"] = "autoclean_kept_crops"
    state["last_workflow_note"] = "Autoclean generated raw-vs-cleaned previews and cleaned analysis inputs for kept crops only."
    output_root = state["output_root"]
    os.makedirs(output_root, exist_ok=True)
    channel_roles = channel_roles_from_state(state)
    global_params = cleaning_params_from_state(state)
    global_channel_params = channel_cleaning_params_from_mapping({}, global_params)
    # Global settings are defaults for crops that do not have an individual
    # override.  Never erase per-crop work just because the global preview is
    # regenerated; individual settings must survive restarts and global runs.
    if max_count is None:
        count, errors = generate_auto_cleaned_crops_for_loaded(state, only_kept=True)
        state["analysis_input_mode"] = "auto_cleaned"
        state["pcc_analysis_input_mode"] = "auto_cleaned"
        state["count_size_analysis_input_mode"] = "auto_cleaned"
        state["last_workflow_action"] = "autoclean_kept_crops"
        state["last_workflow_note"] = "Autoclean generated raw-vs-cleaned previews and cleaned analysis inputs for kept crops only."
        state["last_channel_cleaning_preview_count"] = count
        state["last_channel_cleaning_preview_roles"] = channel_roles
        state["last_channel_cleaning_preview_params"] = cleaning_params_from_state(state)
        state["last_channel_cleaning_preview_errors"] = errors[-50:]
        save_state(state)
        return count
    preview_jobs = []
    for rel, item in sorted(state.get("items", {}).items()):
        if item.get("status", "keep") != "keep":
            continue
        if max_count is not None and len(preview_jobs) >= max(0, int(max_count)):
            break
        src, crop_choice = selected_crop_path_for_item(state, rel, item)
        if not os.path.exists(src):
            continue
        cleaning_params = global_params if force_global else cleaning_params_for_rel(state, rel)
        channel_cleaning_params = global_channel_params if force_global else channel_cleaning_params_for_rel(state, rel)
        analysis_cleaning_params = {**cleaning_params, "channel_params": channel_cleaning_params}
        puncta_cleaned_channels = cleaning_channels_for_rel(state, rel)
        raw_preserved_channels = set(CHANNELS) - puncta_cleaned_channels
        base_name = cleaning_preview_base_name(rel, crop_choice)
        preview_jobs.append((
            src,
            output_root,
            base_name,
            raw_preserved_channels,
            puncta_cleaned_channels,
            analysis_cleaning_params,
        ))

    def run_preview_job(job):
        src, output_root, base_name, raw_preserved_channels, puncta_cleaned_channels, analysis_cleaning_params = job
        prepare_analysis_input_crop(
            src,
            output_root,
            base_name,
            True,
            preserve_raw_channels=raw_preserved_channels,
            puncta_cleaned_channels=puncta_cleaned_channels,
            cleaning_params=analysis_cleaning_params,
        )
        return 1

    preview_count = 0
    if preview_jobs:
        workers = max(1, min(CHANNEL_CLEANING_PREVIEW_WORKERS, len(preview_jobs)))
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = [pool.submit(run_preview_job, job) for job in preview_jobs]
            for future in as_completed(futures):
                preview_count += future.result()
    state["analysis_input_mode"] = "auto_cleaned"
    state["pcc_analysis_input_mode"] = "auto_cleaned"
    state["count_size_analysis_input_mode"] = "auto_cleaned"
    state["last_channel_cleaning_preview_count"] = preview_count
    state["last_channel_cleaning_preview_roles"] = channel_roles
    state["last_channel_cleaning_preview_params"] = cleaning_params_from_state(state)
    save_state(state)
    return preview_count


def persist_channel_cleaning_settings_for_one_crop(state, rel, params):
    """Save per-crop controls without waiting for image regeneration."""
    if rel not in state.get("items", {}):
        raise ValueError(f"Crop is not loaded: {rel}")
    item = state["items"][rel]
    if isinstance(params, dict) and any(f"{channel_name}_role" in params for channel_name in CHANNELS):
        set_item_channel_roles(item, params)
    src, crop_choice = selected_crop_path_for_item(state, rel, item)
    if not os.path.exists(src):
        raise FileNotFoundError(f"Selected crop file does not exist: {src}")
    per_crop = state.setdefault("per_crop_cleaning_params", {})
    if not isinstance(per_crop, dict):
        per_crop = {}
        state["per_crop_cleaning_params"] = per_crop
    global_params = cleaning_params_from_state(state)
    cleaned_params = cleaning_params_from_mapping(params, global_params)
    channel_params = channel_cleaning_params_from_mapping(params, cleaned_params)
    cleaned_channels = channel_set_from_mapping(params, condensate_channels_from_state(state))
    per_crop[rel] = {
        **cleaned_params,
        "channel_params": channel_params,
        "channels": sorted(cleaned_channels),
    }
    save_state(state)
    return {**cleaned_params, "channel_params": channel_params}, cleaned_channels


def preview_channel_cleaning_for_one_crop(state, rel, params):
    ensure_sibling_output_layout(state)
    state["last_workflow_action"] = "autoclean_one_crop_with_adjusted_settings"
    state["last_workflow_note"] = "Autoclean regenerated only the selected crop using its per-crop settings."
    analysis_cleaning_params, cleaned_channels = persist_channel_cleaning_settings_for_one_crop(
        state, rel, params
    )
    item = state["items"][rel]
    src, crop_choice = selected_crop_path_for_item(state, rel, item)
    if not os.path.exists(src):
        raise FileNotFoundError(f"Selected crop file does not exist: {src}")
    output_root = state["output_root"]
    os.makedirs(output_root, exist_ok=True)
    puncta_cleaned_channels = cleaned_channels
    raw_preserved_channels = set(CHANNELS) - puncta_cleaned_channels
    base_name = auto_clean_base_name(rel, crop_choice)
    cleaned_path, preview_path, params_path, analysis_input = prepare_analysis_input_crop(
        src,
        output_root,
        base_name,
        True,
        preserve_raw_channels=raw_preserved_channels,
        puncta_cleaned_channels=puncta_cleaned_channels,
        cleaning_params=analysis_cleaning_params,
    )
    item["auto_cleaned_crop_path"] = cleaned_path
    item["auto_cleaned_preview_path"] = preview_path
    item["auto_cleaned_params_path"] = params_path
    item["auto_cleaned_source_path"] = src
    item["auto_cleaned_source_mtime"] = os.path.getmtime(src)
    item["auto_cleaned_input_mode"] = analysis_input
    item["auto_cleaned_channels"] = sorted(puncta_cleaned_channels)
    item["auto_cleaned_settings_signature"] = auto_cleaned_settings_signature_for_item(state, rel, item, src)
    item["auto_cleaned_generated_at"] = datetime.now().isoformat(timespec="seconds")
    # Keep individual AutoClean consistent with batch AutoClean in the shared
    # PCC/Manders and count/size input dropdown.
    state["analysis_input_mode"] = "auto_cleaned"
    state["pcc_analysis_input_mode"] = "auto_cleaned"
    state["count_size_analysis_input_mode"] = "auto_cleaned"
    save_state(state)
    return preview_path, analysis_cleaning_params


def reset_channel_cleaning_for_one_crop(state, rel):
    ensure_sibling_output_layout(state)
    state["last_workflow_action"] = "reset_one_channel_cleaning"
    state["last_workflow_note"] = "A per-crop cleaning preview was reset because the reset cleaning button was clicked."
    if rel not in state.get("items", {}):
        raise ValueError(f"Crop is not loaded: {rel}")
    per_crop = state.get("per_crop_cleaning_params", {})
    if isinstance(per_crop, dict):
        per_crop.pop(rel, None)
    item = state["items"][rel]
    if isinstance(item, dict):
        item.pop("channel_roles", None)
    src, crop_choice = selected_crop_path_for_item(state, rel, item)
    if not os.path.exists(src):
        raise FileNotFoundError(f"Selected crop file does not exist: {src}")
    output_root = state["output_root"]
    os.makedirs(output_root, exist_ok=True)
    cleaned_params = cleaning_params_for_rel(state, rel)
    channel_params = channel_cleaning_params_for_rel(state, rel)
    analysis_cleaning_params = {**cleaned_params, "channel_params": channel_params}
    puncta_cleaned_channels = cleaning_channels_for_rel(state, rel)
    raw_preserved_channels = set(CHANNELS) - puncta_cleaned_channels
    base_name = cleaning_preview_base_name(rel, crop_choice)
    _, preview_path, _, _ = prepare_analysis_input_crop(
        src,
        output_root,
        base_name,
        True,
        preserve_raw_channels=raw_preserved_channels,
        puncta_cleaned_channels=puncta_cleaned_channels,
        cleaning_params=analysis_cleaning_params,
    )
    save_state(state)
    return preview_path, analysis_cleaning_params, puncta_cleaned_channels


def style_workbook(path):
    wb = load_workbook(path)
    if "Pooled Summary" in wb.sheetnames:
        ws = wb["Pooled Summary"]
        ws.auto_filter.ref = ws.dimensions
        for cell in ws["C"][1:]:
            cell.number_format = "0"
        for column in "DEFGHI":
            for cell in ws[column][1:]:
                cell.number_format = "0.0000"
        for row in ws.iter_rows(min_row=2):
            metric_name = str(row[0].value or "")
            if metric_name == "Corrected_MS2_Intensity" or metric_name.endswith(
                "_Corrected_Intensity_At_Selected_Focus"
            ) or metric_name.endswith("_Partition_Coefficient"):
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor="FFF2CC")
                    cell.font = Font(bold=True)
        try:
            wb.calculation.calcMode = "auto"
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
        except AttributeError:
            pass
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
        for column_cells in ws.columns:
            max_len = 0
            letter = get_column_letter(column_cells[0].column)
            for cell in column_cells[:80]:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, min(len(value), 55))
            ws.column_dimensions[letter].width = max(10, max_len + 2)
        for cell in ws[1]:
            if str(cell.value or "").endswith("_Partition_Coefficient") or str(cell.value or "") == "Partition_Coefficient":
                for data_cell in ws[get_column_letter(cell.column)][1:]:
                    data_cell.number_format = "0.0000"
                    data_cell.fill = PatternFill("solid", fgColor="E2F0D9")
    wb.save(path)



def run_original_script_stage(state, stage):
    os.makedirs(RUNNER_DIR, exist_ok=True)
    ensure_sibling_output_layout(state)
    state["last_workflow_action"] = f"run_stage_{stage}"
    state["last_workflow_note"] = f"Pipeline stage '{stage}' was run because its stage button was clicked."
    raw_folder, preprocessor_name, crop_name, output_name, preprocessor_root, crop_folder, manders_folder = named_pipeline_paths(state)
    crop_base = os.path.join(preprocessor_root, crop_name)
    raw_crop_folder = os.path.join(crop_base, "01_raw_masked_roi_crops_for_jacop")
    enhanced_crop_folder = os.path.join(crop_base, "02_enhanced_roi_crops_for_optional_analysis")

    if stage == "preprocess":
        source = resolve_analysis_script("preprocess")
        runner_name = "run_stage1_enhanced_preprocess.py"
        overrides = {
            "INPUT_FOLDER": raw_folder,
            "OUTPUT_FOLDER_NAME": preprocessor_name,
            "OUTPUT_FOLDER": preprocessor_root,
            "OUTPUT_EXCEL": os.path.join(preprocessor_root, "Enhanced_Preprocessor_QC.xlsx"),
            "RAW_MASKED_JACOP_FOLDER": os.path.join(preprocessor_root, "01_raw_masked_for_jacop"),
            "ENHANCED_JACOP_FOLDER": os.path.join(preprocessor_root, "02_enhanced_masked_for_visual_qc"),
            "MASK_FOLDER": os.path.join(preprocessor_root, "masks"),
            "PREVIEW_FOLDER": os.path.join(preprocessor_root, "previews"),
        }
    elif stage == "crop":
        source = resolve_analysis_script("crop")
        runner_name = "run_stage2_raw_enhanced_crop.py"
        overrides = {
            "EXPERIMENT_ROOT": raw_folder,
            "PREPROCESSOR_FOLDER": preprocessor_root,
            "QC_EXCEL": os.path.join(preprocessor_root, "Enhanced_Preprocessor_QC.xlsx"),
            "RAW_MASKED_IMAGE_FOLDER": os.path.join(preprocessor_root, "01_raw_masked_for_jacop"),
            "ENHANCED_IMAGE_FOLDER": os.path.join(preprocessor_root, "01_raw_masked_for_jacop"),
            "MASK_FOLDER": os.path.join(preprocessor_root, "masks"),
            "OUTPUT_FOLDER": crop_base,
            "RAW_CROP_FOLDER": raw_crop_folder,
            "ENHANCED_CROP_FOLDER": enhanced_crop_folder,
            "OUTPUT_EXCEL": os.path.join(crop_base, "Raw_Cropped_ROI_Index.xlsx"),
        }
    elif stage == "stage3":
        mode = state.get("stage3_mode", "dapi_red_green")
        script_key = "stage3_dapi_red_green" if mode == "dapi_red_green" else "stage3_no_dapi_three_color"
        source = resolve_analysis_script(script_key)
        runner_name = "run_stage3_pcc_manders.py"
        output_excel = "Python_JACoP_Like_Red_Green_Positive_ROI_Results.xlsx" if mode == "dapi_red_green" else "Python_JACoP_Like_Triple_Positive_ROI_Results.xlsx"
        overrides = {
            "IMAGE_FOLDER": crop_folder,
            "OUTPUT_FOLDER": manders_folder,
            "OUTPUT_EXCEL": os.path.join(manders_folder, output_excel),
            "MASK_FOLDER": os.path.join(manders_folder, "per_cell_masks"),
        }
        if mode == "dapi_red_green":
            overrides["ROI_PREVIEW_FOLDER"] = os.path.join(manders_folder, "red_green_roi_preview_images")
        else:
            overrides["ROI_PREVIEW_FOLDER"] = os.path.join(manders_folder, "roi_preview_images")
    else:
        raise ValueError(f"Unknown stage: {stage}")

    runner_path = os.path.join(RUNNER_DIR, runner_name)
    with open(runner_path, "w", encoding="utf-8") as f:
        f.write("source = " + repr(source) + "\n")
        f.write("with open(source, 'r', encoding='utf-8') as _f:\n")
        f.write("    _code = _f.read()\n")
        f.write("_code = _code.replace('bool(s elected_by_marker_rule)', 'bool(selected_by_marker_rule)')\n")
        if stage == "preprocess":
            f.write("_code = _code.replace('    os.makedirs(ENHANCED_JACOP_FOLDER, exist_ok=True)\\n', '')\n")
            f.write("_code = _code.replace('    enhanced_masked = np.array(enhanced_image, copy=True)\\n    enhanced_masked[~output_mask, :] = BACKGROUND_VALUE\\n\\n', '')\n")
            f.write("_code = _code.replace('    enhanced_out = os.path.join(ENHANCED_JACOP_FOLDER, f\"{base}_enhanced_masked_visual_QC.tiff\")\\n', '    enhanced_out = \"\"\\n')\n")
            f.write("_code = _code.replace('    tiff.imwrite(enhanced_out, enhanced_masked, photometric=\"rgb\")\\n', '')\n")
            f.write("_code = _code.replace('    print(f\"Enhanced visual-QC images: {ENHANCED_JACOP_FOLDER}\")\\n', '    print(\"Enhanced visual-QC image output disabled in channel-cleaned app copy.\")\\n')\n")
        if stage == "crop":
            f.write("_code = _code.replace('def find_raw_and_enhanced_images(image_name):\\n    raw_path = find_file(RAW_MASKED_IMAGE_FOLDER, raw_masked_image_name(image_name))\\n    enhanced_path = find_file(ENHANCED_IMAGE_FOLDER, enhanced_image_name(image_name))\\n    return raw_path, enhanced_path\\n', 'def find_raw_and_enhanced_images(image_name):\\n    raw_path = find_file(RAW_MASKED_IMAGE_FOLDER, raw_masked_image_name(image_name))\\n    return raw_path, raw_path\\n')\n")
            f.write("_code = _code.replace('    image_enhanced_folder = os.path.join(ENHANCED_CROP_FOLDER, make_safe_name(image_name))\\n', '    image_enhanced_folder = \"\"\\n')\n")
            f.write("_code = _code.replace('    os.makedirs(image_enhanced_folder, exist_ok=True)\\n', '')\n")
            f.write("_code = _code.replace('    enhanced_out = os.path.join(image_enhanced_folder, enhanced_name)\\n', '    enhanced_out = \"\"\\n')\n")
            f.write("_code = _code.replace('    tiff.imwrite(enhanced_out, enhanced_crop, photometric=\"rgb\")\\n', '')\n")
            f.write("_code = _code.replace('    os.makedirs(ENHANCED_CROP_FOLDER, exist_ok=True)\\n', '')\n")
            f.write("_code = _code.replace('        crop_df.to_excel(writer, sheet_name=\"Raw Enhanced Crop Index\", index=False)\\n', '        crop_df.to_excel(writer, sheet_name=\"Raw Crop Index\", index=False)\\n')\n")
            f.write("_code = _code.replace('    print(f\"Enhanced ROI crops saved here: {ENHANCED_CROP_FOLDER}\")\\n', '    print(\"Enhanced ROI crop output disabled in channel-cleaned app copy.\")\\n')\n")
        f.write("ns = {'__file__': source, '__name__': '__enhanced_lab_app_runner__'}\n")
        f.write("exec(compile(_code, source, 'exec'), ns)\n")
        for key, value in overrides.items():
            f.write(f"ns[{key!r}] = {value!r}\n")
        f.write("ns['main']()\n")

    result = subprocess.run(
        [sys.executable, runner_path],
        cwd=APP_ROOT,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    state["last_stage_output"] = result.stdout[-12000:]
    state["last_stage_returncode"] = result.returncode
    state["crop_root"] = crop_folder
    state["output_root"] = manders_folder
    save_state(state)
    if result.returncode != 0:
        raise RuntimeError(f"Stage failed with code {result.returncode}.\n\n{result.stdout}")
    return result.stdout, overrides

def html_escape(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def render_recovered_crop_panel(state, items=None):
    items = ensure_crop_numbers(state) if items is None else items
    kept = sum(1 for x in items if x.get("status", "keep") == "keep")
    rejected = sum(1 for x in items if x.get("status", "keep") == "reject")
    fixed_cell_annotated_kept = sum(
        1
        for x in items
        if x.get("status", "keep") == "keep" and isinstance(x.get("fixed_cell_manual_annotation"), dict)
    )
    recovered_gallery_cards = {"keep": [], "reject": []}
    recovered_gallery_seen = {"keep": 0, "reject": 0}
    for idx, item in enumerate(items):
        status = item.get("status", "keep")
        if status not in recovered_gallery_cards:
            continue
        recovered_gallery_seen[status] += 1
        extra_class = " auxiliary-extra" if recovered_gallery_seen[status] > AUXILIARY_PREVIEW_LIMIT else ""
        rel = item["rel_path"]
        crop_number = int(item.get("crop_number", idx + 1) or idx + 1)
        selected_path, _ = selected_crop_path_for_item(state, rel, item)
        selected_exists = bool(selected_path and os.path.isfile(selected_path))
        recovered_autoclean = (
            current_auto_cleaned_input_for_item_settings(state, rel, item, selected_path)
            if status == "keep" and selected_exists
            else None
        )
        display_path = recovered_autoclean[0] if recovered_autoclean else selected_path
        display_kind = "auto_cleaned" if recovered_autoclean else "selected"
        thumb_html = (
            f'<img src="/recovered_crop_thumb?rel={quote(rel)}&kind={display_kind}&v={int(os.path.getmtime(display_path))}" loading="lazy" />'
            if display_path and os.path.isfile(display_path)
            else '<div class="raw-missing">Recovered crop image path is missing</div>'
        )
        selected_source_label = "manual corrected crop" if item.get("manual_crop_path") and selected_exists else "raw cropped ROI"
        source_label = (
            f"recovered auto-cleaned image (source: {selected_source_label})"
            if recovered_autoclean
            else selected_source_label
        )
        note = item.get("note", "")
        autoclean_notice = (
            '<div class="small autoclean-needed">Autoclean needed</div>'
            if autoclean_needed_for_item(state, rel, item)
            else ""
        )
        recovered_gallery_cards[status].append(f"""
        <div class="recovered-crop-card{extra_class}" data-recovery-status="{status}" data-rel="{html_escape(rel)}">
          <div class="crop-number-badge">Crop #{crop_number}</div>
          {thumb_html}
          <div class="meta"><b>{html_escape(os.path.basename(rel))}</b></div>
          <div class="small"><b>{html_escape(source_label)}</b><br>{html_escape(note)}</div>
          {autoclean_notice}
        </div>
        """)
    kept_gallery_html = "\n".join(recovered_gallery_cards["keep"]) or '<p data-empty-message="1">No kept crop images are currently loaded.</p>'
    rejected_gallery_html = "\n".join(recovered_gallery_cards["reject"]) or '<p data-empty-message="1">No rejected crop images are currently loaded.</p>'
    kept_more_html = f'<button class="btn secondary auxiliary-show-all" type="button" data-target="recoveredKeptGrid">Show all kept recovered crops ({kept})</button>' if kept > AUXILIARY_PREVIEW_LIMIT else ""
    rejected_more_html = f'<button class="btn secondary auxiliary-show-all" type="button" data-target="recoveredRejectedGrid">Show all rejected recovered crops ({rejected})</button>' if rejected > AUXILIARY_PREVIEW_LIMIT else ""
    return f"""
<div id="recoveredCropsSection" class="panel">
<h2>Recovered kept and rejected cropped images</h2>
<div class="small">These are the actual crop images currently loaded from the previous job state. Kept crops display the recovered auto-cleaned TIFF when one exists; otherwise they display the selected raw or manual crop and show “Autoclean needed.” Rejected crops show the original crop that was excluded.</div>
<div class="notice"><b>Recovered kept:</b> <span id="recoveredKeptCount">{kept}</span> &nbsp; <b>Recovered rejected:</b> <span id="recoveredRejectedCount">{rejected}</span></div>
<details open>
<summary id="recoveredKeptSummary">Recovered kept cropped images ({kept})</summary>
<div id="recoveredKeptGrid" class="recovered-gallery-grid">{kept_gallery_html}</div>
{kept_more_html}
</details>
<details>
<summary id="recoveredRejectedSummary">Recovered rejected cropped images ({rejected})</summary>
<div id="recoveredRejectedGrid" class="recovered-gallery-grid">{rejected_gallery_html}</div>
{rejected_more_html}
</details>
</div>
"""


def render_cleaning_cards(state, items=None):
    items = ensure_crop_numbers(state) if items is None else items
    cleaning_items, cleaning_preview_root = channel_cleaning_preview_items(state)
    cleaning_preview_by_rel = {item["rel_path"]: item for item in cleaning_items}
    cleaning_cards = []
    kept_seen = 0
    for item in items:
        if item.get("status", "keep") != "keep":
            continue
        kept_seen += 1
        extra_class = " auxiliary-extra" if kept_seen > AUXILIARY_PREVIEW_LIMIT else ""
        rel = item["rel_path"]
        crop_number = int(item.get("crop_number", 0) or 0)
        src, crop_choice = selected_crop_path_for_item(state, rel, item)
        preview_rel = cleaning_preview_relpath(rel, crop_choice)
        preview_item = None
        for candidate_rel in cleaning_preview_candidate_relpaths(rel, crop_choice):
            if candidate_rel in cleaning_preview_by_rel:
                preview_rel = candidate_rel
                preview_item = cleaning_preview_by_rel[candidate_rel]
                break
        params = cleaning_params_for_rel(state, rel)
        channel_params = channel_cleaning_params_for_rel(state, rel)
        cleaned_channels = cleaning_channels_for_rel(state, rel)
        needs_autoclean = autoclean_needed_for_item(state, rel, item)
        source_label = "AUTOCLEAN NEEDED" if needs_autoclean else ("PER-CROP ADJUSTED" if has_per_crop_cleaning_params(state, rel) else "GLOBAL SETTINGS")
        crop_roles = channel_roles_for_item(state, item)
        role_source = "PER-CROP ROLES" if isinstance(item.get("channel_roles"), dict) else "GLOBAL ROLES"
        if needs_autoclean:
            preview_html = '<div class="raw-missing autoclean-needed">Autoclean needed: click Autoclean this crop only to regenerate just this crop.</div>'
        elif preview_item:
            preview_html = f'<img src="/channel_cleaning_preview_png?rel={quote(preview_rel)}&v={os.stat(preview_item["path"]).st_mtime_ns}" loading="lazy" />'
        else:
            preview_html = '<div class="raw-missing">No preview yet. Use Autoclean this crop only.</div>'
        channel_controls = []
        for channel_name in ("red", "green", "blue"):
            cp = channel_params[channel_name]
            title = channel_name.title()
            role_select = role_select_html(channel_name, crop_roles.get(channel_name, "ignore"))
            role_cleaning_text = "Will be puncta-cleaned" if crop_roles.get(channel_name) == "condensate" else "Will be preserved raw"
            channel_controls.append(f"""
            <div class="notice channel-role-panel">
              <div class="channel-role-controls">
                <div class="compare-label">{title} channel role</div>
                {role_select}
                <div class="small" data-role="{channel_name}-cleaning-rule"><b>{role_cleaning_text}</b></div>
                <label>{title} raw to lowest background: <b><span data-value-for="{channel_name}_amount">{cp['amount']}</span>%</b><br><input type="range" name="{channel_name}_amount" min="0" max="100" step="1" value="{cp['amount']}" /></label>
                <label>{title} background kept: <b><span data-value-for="{channel_name}_background_removal">{cp.get('background_removal', 0)}</span>%</b><br><input type="range" name="{channel_name}_background_removal" min="0" max="100" step="1" value="{cp.get('background_removal', 0)}" /></label>
                <label>{title} radius: <b><span data-value-for="{channel_name}_radius_px">{cp['radius_px']}</span> px</b><br><input type="range" name="{channel_name}_radius_px" min="4" max="40" step="1" value="{cp['radius_px']}" /></label>
                <label>{title} sensitivity: <b><span data-value-for="{channel_name}_sensitivity">{cp['sensitivity']}</span></b><br><input type="range" name="{channel_name}_sensitivity" min="1" max="100" step="1" value="{cp['sensitivity']}" /></label>
                <label>{title} minimum puncta size: <b><span data-value-for="{channel_name}_min_area_px">{cp['min_area_px']}</span> px</b><br><input type="range" name="{channel_name}_min_area_px" min="1" max="100" step="1" value="{cp['min_area_px']}" /></label>
              </div>
            </div>
            """)
        channel_controls_html = "\n".join(channel_controls)
        cleaning_cards.append(f"""
        <div class="card analyzed-card cleaning-card{extra_class}" data-rel="{html_escape(rel)}">
          <div class="crop-number-badge">Crop #{crop_number}</div>
          <div class="cleaning-preview-panel" data-role="cleaning-preview-panel">{preview_html}</div>
          <div class="meta"><b>{html_escape(os.path.basename(rel))}</b></div>
          <div class="small">{html_escape(rel)}</div>
          <div class="badge" style="background:#315c2b">RAW VS CLEANED</div>
          <div class="small" data-role="cleaning-status"><b>{source_label}</b></div>
          <div class="small" data-role="role-status"><b>{role_source}</b></div>
          <form class="live-cleaning-form" method="post" action="/preview_one_channel_cleaning">
            <input type="hidden" name="rel" value="{html_escape(rel)}" />
            <div class="small"><b>Choose roles, then adjust from raw/no cleaning to lowest-background condensate spots</b></div>
            {channel_controls_html}
            <button class="btn secondary" type="submit">Autoclean this crop only</button>
            <div class="small">Only this crop and its per-crop settings are updated. Other crops and global settings stay unchanged.</div>
          </form>
          <form class="live-reset-cleaning-form" method="post" action="/reset_one_channel_cleaning">
            <input type="hidden" name="rel" value="{html_escape(rel)}" />
            <button class="btn secondary" type="submit">Reset this crop to global settings</button>
          </form>
        </div>
        """)
    return "\n".join(cleaning_cards) or '<p data-empty-message="1">No kept crops are available for raw-vs-cleaned preview yet.</p>', cleaning_preview_root


def render_cleaning_preview_section_parts(state, items=None, show_all=False):
    items = ensure_crop_numbers(state) if items is None else items
    kept_items = [item for item in items if item.get("status", "keep") == "keep"]
    display_items = kept_items if show_all else kept_items[:AUXILIARY_PREVIEW_LIMIT]
    cards_html, cleaning_preview_root = render_cleaning_cards(state, display_items)
    button_html = ""
    if not show_all and len(kept_items) > AUXILIARY_PREVIEW_LIMIT:
        button_html = f'<button class="btn secondary auxiliary-show-all" type="button" data-section="cleaning" data-target="cleaningPreviewGrid">Show all raw-vs-cleaned previews ({len(kept_items)})</button>'
    return cards_html, cleaning_preview_root, button_html, len(kept_items)


def count_size_threshold_crop_rows(state, items=None):
    items = ensure_crop_numbers(state) if items is None else items
    return [
        {
            "rel": item["rel_path"],
            "label": f'{int(item.get("crop_number", idx + 1) or idx + 1)} - {os.path.basename(item["rel_path"])}',
        }
        for idx, item in enumerate(items)
        if item.get("status", "keep") == "keep"
    ]


def page(state, message=""):
    raw_folder, preprocessor_name, crop_name, output_name, preprocessor_root, derived_crop_root, derived_output_root = named_pipeline_paths(state)
    items = ensure_crop_numbers(state)
    kept = sum(1 for x in items if x.get("status", "keep") == "keep")
    rejected = sum(1 for x in items if x.get("status", "keep") == "reject")
    fixed_cell_annotated_kept = sum(
        1
        for x in items
        if x.get("status", "keep") == "keep" and isinstance(x.get("fixed_cell_manual_annotation"), dict)
    )
    partition_annotated_kept = sum(
        1
        for x in items
        if x.get("status", "keep") == "keep" and isinstance(x.get("partition_manual_annotation"), dict)
    )
    fixed_cell_pending_qc = sum(
        1
        for x in items
        if x.get("status", "keep") == "keep" and bool(x.get("fixed_cell_annotation_requires_analysis"))
    )
    persistent_notice_html = ""
    last_load_summary = state.get("last_previous_job_load_summary", {})
    if (
        last_load_summary.get("status") == "preprocessor_ready"
        and raw_folder
        and not items
        and not os.path.isdir(derived_crop_root)
    ):
        preprocessor_ready_root = last_load_summary.get("preprocessor_ready_root") or os.path.join(preprocessor_root, "01_raw_masked_for_jacop")
        preprocessor_ready_tiffs = last_load_summary.get("preprocessor_ready_tiffs")
        if preprocessor_ready_tiffs is None:
            preprocessor_ready_tiffs = len(recovery_find_tiffs(preprocessor_ready_root))
        persistent_notice_html = f"""
        <div class="notice">
          <b>Preprocessor output is ready.</b><br>
          {int(preprocessor_ready_tiffs or 0)} raw-masked TIFF(s) were found in: {html_escape(preprocessor_ready_root)}<br>
          No reviewed ROI crops exist yet, and no workflow step was run automatically. Click <b>2. Crop raw-masked ROIs for final choice</b> below when you want to create: {html_escape(derived_crop_root)}
        </div>
        """
    threshold_crop_options = "\n".join(
        f'<option value="{html_escape(row["rel"])}">{html_escape(row["label"])}</option>'
        for row in count_size_threshold_crop_rows(state, items)
    )
    if not threshold_crop_options:
        threshold_crop_options = '<option value="">No kept crops yet</option>'
    cards = []
    for idx, item in enumerate(items):
        rel = item["rel_path"]
        crop_number = int(item.get("crop_number", idx + 1) or idx + 1)
        status = item.get("status", "keep")
        border = "#188038" if status == "keep" else "#b3261e"
        label = "KEEP" if status == "keep" else "REJECTED"
        toggle = "reject" if status == "keep" else "keep"
        toggle_text = "Reject this ROI" if status == "keep" else "Keep this ROI"
        card_id = f"roi-card-{idx}"
        badge_id = f"roi-badge-{idx}"
        button_id = f"roi-button-{idx}"
        chosen_label_id = f"chosen-label-{idx}"
        chosen_panel_id = f"chosen-panel-{idx}"
        rel_js = json.dumps(rel)
        card_js = json.dumps(card_id)
        badge_js = json.dumps(badge_id)
        button_js = json.dumps(button_id)
        context = crop_context_for_rel(state, rel)
        raw_m_path = find_raw_m_image(state, rel)
        crop_choice = item.get("crop_choice", "raw")
        has_manual = bool(item.get("manual_crop_path") and os.path.isfile(item.get("manual_crop_path", "")))
        has_fixed_annotation = isinstance(item.get("fixed_cell_manual_annotation"), dict)
        has_partition_annotation = isinstance(item.get("partition_manual_annotation"), dict)
        raw_selected = "selected-choice" if crop_choice == "raw" and not has_manual else ""
        if has_manual:
            chosen_label = "Manual corrected crop"
            manual_version = str(int(os.path.getmtime(item.get("manual_crop_path", ""))))
            chosen_panel = '<img src="/thumb?rel=' + quote(rel) + '&v=' + manual_version + '" loading="lazy" />'
        else:
            chosen_label = "Raw cropped ROI"
            chosen_kind = "raw"
            chosen_path, _ = selected_crop_path_for_item(state, rel, item)
            chosen_panel = '<img src="/index_crop_thumb?kind=' + chosen_kind + '&rel=' + quote(rel) + '" loading="lazy" />' if os.path.isfile(chosen_path) else '<div class="raw-missing">Selected crop image not found</div>'
        raw_m_panel = '<img src="/raw_m_context_thumb?rel=' + quote(rel) + '" loading="lazy" />' if raw_m_path else '<div class="raw-missing">No raw M image found</div>'
        raw_m_note = html_escape(raw_m_path) if raw_m_path else "No matching raw M image found"
        source_key = crop_source_key(rel)
        explicit_keep = bool(item.get("explicit_keep"))
        autoclean_notice = (
            '<div class="autoclean-needed" data-role="autoclean-needed">Autoclean needed for this crop. Use Autoclean only this crop below or its controls in Raw vs cleaned image previews.</div>'
            if autoclean_needed_for_item(state, rel, item)
            else ""
        )
        cards.append(f"""
        <div id="{card_id}" class="card review-card" data-rel="{html_escape(rel)}" data-source-key="{html_escape(source_key)}" data-number="{crop_number}" data-status="{status}" data-explicit-keep="{str(explicit_keep).lower()}" style="border-color:{border}">
          <div class="crop-number-badge">Crop #{crop_number}</div>
          <div class="choice-strip">
            <form method="post" action="/set_crop_choice">
              <input type="hidden" name="rel" value="{html_escape(rel)}" />
              <input type="hidden" name="crop_choice" value="raw" />
              <button class="choice-btn {raw_selected}" type="submit">Use raw cropped ROI</button>
            </form>
          </div>
          <div class="compare-pair">
            <div class="compare-panel"><div id="{chosen_label_id}" data-role="chosen-label" class="compare-label">Selected: {chosen_label}</div><div id="{chosen_panel_id}" data-role="chosen-panel">{chosen_panel}</div></div>
            <div class="compare-panel"><div class="compare-label">Raw M merged image with crop box</div>{raw_m_panel}</div>
          </div>
          <div class="meta"><b>{html_escape(os.path.basename(rel))}</b></div>
          <div class="small">Crop: {html_escape(rel)}</div>
          <div class="small">Raw M: {raw_m_note}</div>
          <div id="{badge_id}" class="status-banner {status}">
            <span>{label}</span>
            <span>{"Included in final analysis" if status == "keep" else "Excluded from final analysis"}</span>
          </div>
          {autoclean_notice}
          <div class="review-actions">
            <button class="btn secondary manual-recrop-button" type="button" data-rel="{html_escape(rel)}">Manual recrop this ROI</button>
            {f'<button class="btn secondary undo-manual-button" type="button" data-rel="{html_escape(rel)}">Undo manual recrop</button>' if has_manual else ''}
            <button class="btn fixed-cell-annotation-button" type="button" data-rel="{html_escape(rel)}">{'Edit saved ROI + MCP focus + condensates' if has_fixed_annotation else 'Draw ROI + MCP focus + condensates'}</button>
            <button class="btn partition-annotation-button" type="button" data-rel="{html_escape(rel)}">{'Edit measurement ROI + condensates' if has_partition_annotation else 'Draw measurement ROI + condensates'}</button>
            <button class="btn secondary autoclean-one-crop-button" type="button" data-rel="{html_escape(rel)}">Autoclean only this crop</button>
            <form class="status-form" method="post" action="/set_status" target="review_sink" data-card-id="{card_id}" data-badge-id="{badge_id}" data-status="keep">
              <input type="hidden" name="rel" value="{html_escape(rel)}" />
              <input class="status-input" type="hidden" name="status" value="keep" />
              <button class="btn keep-btn" data-choice-button="keep" type="submit">Keep this ROI</button>
            </form>
            <form class="status-form" method="post" action="/set_status" target="review_sink" data-card-id="{card_id}" data-badge-id="{badge_id}" data-status="reject">
              <input type="hidden" name="rel" value="{html_escape(rel)}" />
              <input class="status-input" type="hidden" name="status" value="reject" />
              <button class="btn reject-btn" data-choice-button="reject" type="submit">Reject this ROI</button>
            </form>
            <button class="btn reject-btn reject-image-rest-button" type="button" data-rel="{html_escape(rel)}">Reject other crops from this image</button>
          </div>
        </div>
        """)
    cards_html = "\n".join(cards) or "<p>No crops loaded yet. Set the crop folder and click Scan crops.</p>"
    recovered_panel_html = render_recovered_crop_panel(state, items)
    approved_items, approved_root = approved_crop_items(state)
    final_png_items, final_preview_root = final_preview_png_items(state)
    approved_cards = []
    if final_png_items:
        for item in final_png_items:
            rel = item["rel_path"]
            approved_cards.append(f"""
            <div class="card analyzed-card">
              <img src="/final_preview_png?rel={quote(rel)}" loading="lazy" />
              <div class="meta"><b>{html_escape(os.path.basename(rel))}</b></div>
              <div class="small">{html_escape(rel)}</div>
              <div class="badge" style="background:#1957a6">ANALYZED PREVIEW</div>
            </div>
            """)
    else:
        for item in approved_items:
            rel = item["rel_path"]
            approved_cards.append(f"""
            <div class="card analyzed-card">
              <img src="/approved_thumb?rel={quote(rel)}" loading="lazy" />
              <div class="meta"><b>{html_escape(os.path.basename(rel))}</b></div>
              <div class="small">{html_escape(rel)}</div>
              <div class="badge" style="background:#1957a6">ANALYZED</div>
            </div>
            """)
    approved_cards_html = "\n".join(approved_cards) or "<p>No analysis previews yet. Run <b>PCC/Manders</b> or <b>count/size</b> on kept crops first.</p>"
    fixed_cell_png_items, fixed_cell_preview_root = fixed_cell_preview_png_items(state)
    fixed_cell_cards = []
    for item in fixed_cell_png_items:
        rel = item["rel_path"]
        try:
            preview_version = str(os.stat(item["path"]).st_mtime_ns)
        except OSError:
            preview_version = "0"
        fixed_cell_cards.append(f"""
        <div class="card analyzed-card">
          <img src="/fixed_cell_preview_png?rel={quote(rel)}&v={preview_version}" loading="lazy" />
          <div class="meta"><b>{html_escape(os.path.basename(rel))}</b></div>
          <div class="small">{html_escape(rel)}</div>
          <div class="badge" style="background:#1957a6">FIXED-CELL QC OVERLAY</div>
        </div>
        """)
    fixed_cell_cards_html = "\n".join(fixed_cell_cards) or "<p>No fixed-cell QC overlays yet. Run <b>fixed-cell MS2/condensate measurements</b> first.</p>"
    cleaning_cards_html, cleaning_preview_root, cleaning_more_html, _ = render_cleaning_preview_section_parts(state, items)
    raw_crop_root, _ = raw_and_enhanced_crop_roots(state)
    raw_crop_items = crop_preview_items(raw_crop_root)

    def build_choice_cards(items, kind, root):
        cards_out = []
        for idx, item in enumerate(items, start=1):
            rel = item["rel_path"]
            extra_class = " auxiliary-extra" if idx > AUXILIARY_PREVIEW_LIMIT else ""
            cards_out.append(f"""
            <div class="card choice-card{extra_class}">
              <img src="/choice_thumb?kind={kind}&rel={quote(rel)}" loading="lazy" />
              <div class="meta"><b>{html_escape(os.path.basename(rel))}</b></div>
              <div class="small">{html_escape(rel)}</div>
            </div>
            """)
        return "\n".join(cards_out) or f"<p>No TIFF crops found yet in:<br>{html_escape(root)}</p>"

    raw_choice_cards_html = build_choice_cards(raw_crop_items, "raw", raw_crop_root)
    raw_choice_more_html = f'<button class="btn secondary auxiliary-show-all" type="button" data-target="rawChoiceGrid">Show all raw crops used for review ({len(raw_crop_items)})</button>' if len(raw_crop_items) > AUXILIARY_PREVIEW_LIMIT else ""
    current_choice_label = "01 raw masked for JACoP, original intensities"
    pcc_excel = state.get("last_real_fiji_pcc_manders_excel") or state.get("last_pcc_manders_excel", "")
    count_size_excel = state.get("last_count_size_excel", "")
    fixed_cell_excel = state.get("last_fixed_cell_ms2_excel", "")
    partition_excel = state.get("last_partition_excel", "")
    pcc_preview_root = state.get("last_pcc_manders_preview_root", "")
    count_size_output_root = state.get("last_count_size_output_root", "")
    latest_outputs_html = ""
    if pcc_excel or count_size_excel or fixed_cell_excel or partition_excel:
        latest_outputs_html = f"""
        <div class="panel">
        <h2>Latest analysis outputs</h2>
        <div class="small">
        PCC/Manders Excel: {html_escape(pcc_excel or 'not run yet')}<br>
        PCC/Manders previews: {html_escape(pcc_preview_root or 'not run yet')}<br>
        Condensate count/size Excel: {html_escape(count_size_excel or 'not run yet')}<br>
        Condensate count/size output folder: {html_escape(count_size_output_root or 'not run yet')}<br>
        Fixed-cell MS2/condensate Excel: {html_escape(fixed_cell_excel or 'not run yet')}<br>
        Fixed-cell MS2/condensate output folder: {html_escape(state.get('last_fixed_cell_ms2_output_root', '') or 'not run yet')}
        <br>Manual nuclear partition-coefficient Excel: {html_escape(partition_excel or 'not run yet')}<br>
        Manual nuclear partition-coefficient output folder: {html_escape(state.get('last_partition_output_root', '') or 'not run yet')}
        </div>
        </div>
        """

    return f"""<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>Condensate Analysis App</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, Helvetica, Arial, sans-serif; margin: 22px; background:#f7f7f5; color:#1f1f1f; }}
h1 {{ margin-bottom: 4px; }}
.boundary-recovery-banner {{ display:flex; flex-wrap:wrap; align-items:center; gap:8px 16px; margin:10px 0 14px; padding:12px 15px; border:2px solid #15803d; border-radius:8px; background:#ecfdf5; color:#14532d; font-weight:800; box-shadow:0 0 0 3px rgba(21,128,61,0.10); }}
.boundary-recovery-banner .active-label {{ font-size:16px; }}
.boundary-recovery-banner .build-label {{ padding:4px 7px; border-radius:5px; background:#14532d; color:white; font-family:ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size:12px; overflow-wrap:anywhere; }}
.boundary-recovery-banner .method-label {{ font-size:12px; font-weight:700; color:#166534; }}
.panel {{ background:white; border:1px solid #ddd; border-radius:8px; padding:16px; margin:14px 0; }}
input[type=text] {{ width: 96%; padding:9px; margin:5px 0 12px; border:1px solid #bbb; border-radius:5px; }}
input[type=file] {{ width: 96%; padding:9px; margin:5px 0 8px; border:1px dashed #777; border-radius:5px; background:#fafafa; }}
select {{ padding:8px; }}
.form-grid {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:12px 14px; align-items:start; }}
.form-grid label {{ display:block; min-width:0; }}
.form-grid input[type=text], .form-grid select {{ width:100%; box-sizing:border-box; margin:5px 0 0; }}
.notice h3 {{ margin:18px 0 6px; }}
.btn {{ background:#1957a6; color:white; border:0; border-radius:5px; padding:9px 12px; cursor:pointer; margin-top:8px; }}
.btn.secondary {{ background:#555; }}
.grid {{ display:grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap:14px; }}
.auxiliary-extra {{ display:none; }}
.choice-grid {{ display:grid; grid-template-columns: 1fr 1fr; gap:18px; align-items:start; }}
.choice-column {{ border:1px solid #ddd; border-radius:8px; padding:12px; background:#fbfbfb; }}
.choice-column.selected-choice {{ border:3px solid #1957a6; background:#eef6ff; }}
.choice-card {{ border-color:#777; }}
.card {{ background:white; border:3px solid #188038; border-radius:8px; padding:10px; }}
.analyzed-card {{ border-color:#1957a6; }}
.card img {{ width:100%; max-height:260px; object-fit:contain; background:#000; border-radius:4px; }}
.summary-result {{ border:1px solid #d6dde8; border-radius:8px; padding:12px; background:#fbfdff; margin-top:12px; }}
.summary-result img {{ width:100%; max-height:520px; object-fit:contain; background:white; border:1px solid #e5e7eb; border-radius:6px; }}
.meta {{ font-size:13px; margin-top:8px; overflow-wrap:anywhere; }}
.small {{ font-size:11px; color:#555; overflow-wrap:anywhere; min-height:32px; }}
.badge {{ display:inline-block; color:white; font-weight:bold; font-size:12px; padding:4px 8px; border-radius:4px; margin-top:6px; }}
.crop-number-badge {{ display:inline-flex; align-items:center; justify-content:center; min-width:84px; padding:6px 10px; border-radius:6px; background:#111827; color:white; font-weight:900; margin-bottom:8px; }}
.number-select-form {{ display:flex; gap:10px; flex-wrap:wrap; align-items:end; margin:12px 0; padding:10px; background:#f8fafc; border:1px solid #cbd5e1; border-radius:8px; }}
.number-select-form input {{ min-width:280px; }}
.notice {{ background:#eef6ff; border-left:5px solid #1957a6; padding:10px; margin:10px 0; }}
.message-notice b {{ display:block; white-space:pre-wrap; overflow-wrap:anywhere; line-height:1.35; }}
details {{ margin-top:12px; }}
summary {{ cursor:pointer; font-weight:bold; color:#333; }}

.compare-pair {{ display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:10px; align-items:start; }}
.compare-panel {{ min-width:0; }}
.compare-panel img {{ width:100%; height:auto; max-height:330px; object-fit:contain; background:#050505; border-radius:8px; border:1px solid #d9e2f3; }}
.compare-label {{ font-size:12px; font-weight:700; margin:0 0 6px; color:#334155; }}
.raw-missing {{ min-height:180px; display:grid; place-items:center; background:#fff7ed; border:1px solid #fdba74; border-radius:8px; color:#9a3412; padding:10px; font-size:12px; }}
.card.review-card {{ max-width:980px; }}
.choice-strip {{ display:flex; gap:8px; flex-wrap:wrap; margin-bottom:10px; }}
.choice-btn {{ border:1px solid #cbd5e1; border-radius:8px; padding:8px 10px; background:#f8fafc; cursor:pointer; font-weight:700; }}
.choice-btn.selected-choice {{ background:#1957a6; color:white; border-color:#1957a6; }}
.role-grid {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(180px, 1fr)); gap:8px; align-items:end; }}
.cleaning-preview-panel {{ position:sticky; top:8px; z-index:4; background:white; padding-bottom:6px; }}
.cleaning-preview-panel img {{ width:100%; max-height:330px; object-fit:contain; background:#050505; border-radius:8px; border:1px solid #d9e2f3; }}
.channel-role-panel {{ margin-top:8px; padding:8px 10px; }}
.channel-role-controls {{ display:grid; gap:6px; }}
.channel-role-controls label {{ display:block; min-width:0; }}
.channel-role-controls select {{ width:100%; box-sizing:border-box; }}
.channel-role-controls input[type=range] {{ width:min(230px, 100%); }}
.review-filter-bar {{ display:flex; gap:8px; flex-wrap:wrap; align-items:center; margin:12px 0; }}
.review-filter-bar button {{ border:1px solid #cbd5e1; border-radius:8px; padding:8px 10px; background:#f8fafc; color:#111827; cursor:pointer; font-weight:800; }}
.review-filter-bar button.active {{ background:#1957a6; border-color:#1957a6; color:white; }}
.recovered-gallery-grid {{ display:grid; grid-template-columns:repeat(auto-fill, minmax(190px, 1fr)); gap:12px; margin-top:10px; }}
.recovered-crop-card {{ border:2px solid #cbd5e1; border-radius:8px; background:white; padding:9px; min-width:0; }}
.recovered-crop-card[data-recovery-status="keep"] {{ border-color:#15803d; background:#f0fdf4; }}
.recovered-crop-card[data-recovery-status="reject"] {{ border-color:#b91c1c; background:#fef2f2; }}
.recovered-crop-card img {{ width:100%; height:170px; object-fit:contain; background:#050505; border-radius:6px; }}
@media (max-width: 1100px) {{ .compare-pair {{ grid-template-columns:repeat(2, minmax(0, 1fr)); }} }}
@media (max-width: 700px) {{ .compare-pair, .channel-role-panel {{ grid-template-columns:1fr; }} }}

.card[data-status="keep"] {{ border-color:#15803d !important; background:#f0fdf4; box-shadow:0 0 0 4px rgba(21,128,61,0.14); }}
.card[data-status="reject"] {{ border-color:#b91c1c !important; background:#fef2f2; box-shadow:0 0 0 4px rgba(185,28,28,0.18); opacity:0.86; }}
.status-banner {{ display:flex; align-items:center; justify-content:space-between; gap:10px; padding:9px 11px; border-radius:8px; color:white; font-weight:800; margin:10px 0; letter-spacing:0.02em; }}
.status-banner.keep {{ background:#15803d; }}
.status-banner.reject {{ background:#b91c1c; }}
.autoclean-needed {{ margin:8px 0; padding:8px 10px; border-left:4px solid #b45309; border-radius:6px; background:#fff7ed; color:#7c2d12; font-weight:800; }}
.cleaning-card .autoclean-needed, .recovered-crop-card .autoclean-needed {{ font-size:12px; }}
.cleaning-preview-panel .autoclean-needed {{ min-height:120px; display:flex; align-items:center; justify-content:center; text-align:center; }}
.review-actions {{ display:flex; gap:8px; flex-wrap:wrap; align-items:center; margin-top:8px; }}
.keep-btn {{ background:#15803d !important; color:white !important; }}
.reject-btn {{ background:#b91c1c !important; color:white !important; }}
.card[data-status="keep"] .keep-btn,
.card[data-status="reject"] .reject-btn {{ outline:4px solid rgba(15,23,42,0.18); transform:scale(1.02); }}
.status-clicked {{ box-shadow:0 0 0 5px rgba(15,23,42,0.22) !important; filter:saturate(1.25); }}
.manual-modal {{ position:fixed; inset:0; display:none; z-index:20; background:rgba(15,23,42,0.72); padding:24px; overflow:auto; }}
.manual-modal.open {{ display:block; }}
.manual-box {{ background:white; border-radius:10px; padding:16px; max-width:1180px; margin:0 auto; }}
.manual-toolbar {{ display:flex; gap:8px; flex-wrap:wrap; align-items:end; margin:10px 0; }}
.manual-toolbar input {{ width:95px; padding:8px; border:1px solid #bbb; border-radius:5px; }}
.manual-canvas-wrap {{ max-height:70vh; overflow:auto; background:#0f172a; border:1px solid #cbd5e1; border-radius:8px; }}
#manualCropCanvas {{ display:block; max-width:100%; height:auto; cursor:crosshair; }}
.fixed-cell-split-grid {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; }}
.fixed-cell-split-pane {{ min-width:0; background:#0f172a; border:1px solid #cbd5e1; border-radius:8px; overflow:hidden; }}
.fixed-cell-split-pane h3 {{ margin:0; padding:7px 9px; background:#172033; color:white; font-size:13px; }}
.fixed-cell-annotation-canvas {{ display:block; width:100%; max-width:100%; height:auto; cursor:default; touch-action:none; }}
@media (max-width:900px) {{ .fixed-cell-split-grid {{ grid-template-columns:1fr; }} }}
.annotation-tool-active {{ outline:4px solid rgba(25,87,166,0.28); }}
.threshold-panel {{ border:1px solid #d6dde8; border-radius:8px; padding:14px; background:#fbfdff; margin-top:12px; display:grid; gap:10px; }}
.threshold-controls {{ display:flex; align-items:center; gap:10px; flex-wrap:wrap; }}
.threshold-controls input[type=range] {{ width:min(360px, 100%); accent-color:#1957a6; }}
.threshold-controls input[type=number] {{ width:82px; font-weight:700; text-align:center; }}
.threshold-crop-display {{ min-width:min(520px, 100%); max-width:820px; padding:9px 12px; border:1px solid #d6dde8; border-radius:8px; background:white; color:#223; font-weight:650; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.threshold-histogram-wrap {{ display:grid; gap:4px; max-width:720px; }}
.threshold-histogram-wrap canvas {{ width:100%; height:96px; background:#050506; border:1px solid #cbd5e1; border-radius:6px; }}
.threshold-scale {{ display:flex; gap:14px; align-items:center; flex-wrap:wrap; color:#555; font-size:12px; }}
.threshold-scale span {{ display:inline-block; min-width:34px; text-align:center; }}
.threshold-preview-grid {{ display:grid; grid-template-columns:repeat(3, minmax(0,1fr)); gap:12px; }}
.threshold-preview-pane {{ border:1px solid #d6dde8; background:white; border-radius:8px; min-width:0; display:grid; grid-template-rows:auto minmax(0,1fr); overflow:hidden; }}
.threshold-preview-pane h3 {{ font-size:13px; margin:0; padding:9px 10px; border-bottom:1px solid #d6dde8; color:#555; font-weight:650; }}
.threshold-preview-box {{ background:#050506; min-height:260px; display:grid; place-items:center; overflow:auto; padding:10px; color:#94a3b8; text-align:center; }}
.threshold-preview-box img {{ max-width:100%; max-height:52vh; object-fit:contain; }}
@media (max-width: 900px) {{ .threshold-preview-grid {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>
<h1>Condensate Analysis App</h1>
<div class="boundary-recovery-banner" id="boundaryRecoveryBuildBanner">
  <span class="active-label">✓ UPDATED PUNCTA BOUNDARY RECOVERY IS ACTIVE</span>
  <span class="build-label">Build: {html_escape(APP_BUILD_ID)}</span>
  <span class="method-label">Local hysteresis reconstruction · original pixel intensities preserved</span>
</div>
<div id="manualCropModal" class="manual-modal">
  <div class="manual-box">
    <h2>Manual Recrop Cell</h2>
    <div class="small" id="manualCropLabel">Drag a rectangle inside the already-cropped cell image.</div>
    <div class="manual-canvas-wrap"><canvas id="manualCropCanvas"></canvas></div>
    <div class="manual-toolbar">
      <label>crop from<br><select id="manualSource"><option value="crop">Already-cropped ROI</option><option value="original">Original source image</option></select></label>
      <label>x<br><input id="manualX" type="number" min="0"></label>
      <label>y<br><input id="manualY" type="number" min="0"></label>
      <label>width<br><input id="manualW" type="number" min="1"></label>
      <label>height<br><input id="manualH" type="number" min="1"></label>
      <button id="saveManualCropButton" class="btn" type="button">Save corrected crop</button>
      <button class="btn secondary" type="button" onclick="closeManualCrop()">Cancel</button>
    </div>
    <div class="small">Saving a manual recrop updates the same review card as KEEP. You can crop from the already-cropped ROI or from the original source image.</div>
  </div>
</div>
<div id="fixedCellAnnotationModal" class="manual-modal">
  <div class="manual-box">
    <h2 id="fixedCellAnnotationTitle">Draw measurement ROI, MCP/MS2 focus, and condensates</h2>
    <div class="small" id="fixedCellAnnotationLabel">Draw one ROI, at least one MCP/MS2 focus, and optional condensate polygons.</div>
    <div class="manual-toolbar" style="margin:8px 0">
      <label>Image used for drawing and measurement<br>
        <select id="fixedCellAnnotationSource">
          <option value="crop">Already-cropped ROI</option>
          <option value="original">Original raw image</option>
          <option value="untouched">Matching untouched acquisition TIFF + reviewed crop box</option>
        </select>
      </label>
      <label>Split-channel display only<br>
        <select id="fixedCellChannelDisplay">
          <option value="pseudocolor" selected>Channel colors (red / green / blue)</option>
          <option value="grayscale">Grayscale</option>
        </select>
      </label>
      <label>Display contrast only<br>
        <select id="fixedCellDisplayContrast">
          <option value="original" selected>Original raw brightness</option>
          <option value="auto">Auto-enhance faint signal for drawing</option>
        </select>
      </label>
      <label>Drawing view<br>
        <select id="fixedCellDrawingView">
          <option value="reviewed" selected>Enlarge reviewed cell</option>
          <option value="full">Full acquisition</option>
        </select>
      </label>
    </div>
    <div class="small" style="margin:-2px 0 8px"><b>Display controls do not change measurements.</b> For an 8-bit acquisition, the merged panel uses the matched TIFF's direct RGB values. Split panels may be colorized or optionally brightened only on screen; analysis still uses untouched grayscale channel values.</div>
    <div class="fixed-cell-split-grid">
      <div class="fixed-cell-split-pane"><h3>Merged reference</h3><canvas id="fixedCellAnnotationCanvas" class="fixed-cell-annotation-canvas" data-channel="merged"></canvas></div>
      <div class="fixed-cell-split-pane"><h3>Red channel</h3><canvas class="fixed-cell-annotation-canvas" data-channel="red"></canvas></div>
      <div class="fixed-cell-split-pane"><h3>Green channel</h3><canvas class="fixed-cell-annotation-canvas" data-channel="green"></canvas></div>
      <div class="fixed-cell-split-pane"><h3>Blue channel</h3><canvas class="fixed-cell-annotation-canvas" data-channel="blue"></canvas></div>
    </div>
    <div class="manual-toolbar">
      <button id="drawMeasurementRoiButton" class="btn" type="button">Draw/replace measurement ROI</button>
      <button id="addMcpFocusButton" class="btn secondary" type="button">Add an MCP/MS2 focus</button>
      <button id="addCondensateButton" class="btn secondary" type="button">Add a condensate ROI</button>
      <button id="undoCondensateButton" class="btn secondary" type="button">Undo last condensate</button>
      <button id="drawMs2BackgroundButton" class="btn secondary" type="button">Draw/replace MS2 background ROI (optional)</button>
      <button id="removeMs2BackgroundButton" class="btn secondary" type="button">Remove manual background — use automatic</button>
      <label id="partitionNoCondensateWrap" style="display:none; align-items:center; gap:6px"><input id="partitionNoCondensate" type="checkbox"> No condensate in this cell</label>
      <button id="finishAnnotationShapeButton" class="btn secondary" type="button">Finish current shape</button>
      <button id="undoMcpFocusButton" class="btn secondary" type="button">Undo last focus</button>
      <button id="clearFixedCellAnnotationButton" class="btn secondary" type="button">Clear all drawings</button>
      <button id="saveFixedCellAnnotationButton" class="btn" type="button">Save ROI + focus + condensates + background</button>
      <button class="btn secondary" type="button" onclick="closeFixedCellAnnotation()">Cancel</button>
    </div>
    <div id="fixedCellSaveStatus" class="notice" style="display:none; margin-top:8px; font-weight:700"></div>
    <div class="small" id="fixedCellAnnotationHelp">Choose the untouched acquisition option to see the exact reviewed PCC/Count-Size crop as a yellow box. Draw only inside that box. The same coordinates appear on all four synchronized channel panels. Cyan is the measurement ROI, magenta is MCP/MS2 focus, green is condensates, and orange is optional background.</div>
  </div>
</div>
<div class="notice">Independent PCC/Manders and count/size-only app copy. Review crops here, then run either retained analysis on the same kept crops.</div>
{f'<div class="notice message-notice"><b>{html_escape(message)}</b></div>' if message else ''}
{persistent_notice_html}
<div class="panel">
<h2>Full Pipeline</h2>
<form id="uploadRawFolderForm" method="post" action="/upload_raw_folder" enctype="multipart/form-data" style="margin-bottom:16px">
<label>Copy/upload a SMALL experiment folder into this app (maximum 256 MB)</label><br>
<input id="rawFolderInput" type="file" name="files" webkitdirectory directory multiple />
<button class="btn" type="submit">Upload and use selected folder</button>
<div id="uploadStatus" class="small">Two steps: choose the folder, then click the blue Upload button. For thousands of files or anything over 256 MB, use the original-location path box below instead.</div>
</form>
<form method="post" action="/use_local_raw_folder_path" style="margin-bottom:12px">
<label>Open a NEW experiment from its original location (recommended for large folders)</label><br>
<input type="text" name="raw_image_folder_path" value="{html_escape(state.get('raw_image_folder',''))}" placeholder="Choose or enter the experiment raw-image folder" />
<button class="btn" type="submit">Use pasted raw image folder</button>
<div class="small">Mac: select the folder in Finder and press Option-Command-C. Windows: use Copy as path. Results stay beside the original images.</div>
</form>
<form method="post" action="/load_previous_job_folder" style="margin-bottom:16px">
<label>Reopen and correct ANY experiment that contains previous app data/results</label><br>
<input type="text" name="previous_job_folder" value="{html_escape(state.get('last_previous_job_folder',''))}" placeholder="Paste the experiment, results, raw_masked_analysis, or crop folder path" />
<button class="btn" type="submit">Load previous experiment and recover its work</button>
<div class="small">The app auto-detects crops, KEEP/REJECT decisions, manual recrops, thresholds, and prior outputs. It does not rerun analysis until you click an analysis button.</div>
</form>
<hr>
<form method="post" action="/pipeline_config">
<div class="small"><b>Current active experiment folder:</b> {html_escape(raw_folder)}</div>
<div class="notice"><b>All generated results will stay inside this uploaded folder.</b><br>
Preprocessor outputs: {html_escape(preprocessor_root)}<br>
Currently selected crop set: {html_escape(derived_crop_root)}<br>
Real Fiji PCC/Manders results: {html_escape(derived_output_root)}</div>
<label>Preprocessor folder name</label><br>
<input type="text" name="preprocessor_folder_name" value="{html_escape(state.get('preprocessor_folder_name','ENHANCED_JACOP_PREPROCESSOR_RESULTS'))}" />
<label>Crop review folder name</label><br>
<input type="text" name="crop_folder_name" value="{html_escape(state.get('crop_folder_name','03_cropped_rois_for_final_choice'))}" />
<label>Combined results folder name</label><br>
<input type="text" name="final_output_folder_name" value="{html_escape(state.get('final_output_folder_name','REAL FIJI PCC MANDERS RESULTS'))}" />
	<div class="notice"><b>Review crop set:</b> 01 raw masked for JACoP, original intensities.</div>
<label>Pixel size, um</label><br>
<input type="number" step="0.0000001" name="pixel_um" value="{html_escape(state.get('pixel_um', 0.1290271))}" />
<div class="small">PCC/Manders is generated headlessly in real Fiji/ImageJ using the same exported PCC, overlap, k1/k2, and Manders formulas with ImageJ auto-thresholds. Count/size uses Fiji Analyze Particles.</div>
<input class="btn" type="submit" value="Save pipeline settings" />
</form>
<form method="post" action="/run_stage" style="display:inline"><input type="hidden" name="stage" value="preprocess" /><input class="btn" type="submit" value="1. Run preprocessing and ROI selection" /></form>
<form method="post" action="/run_stage" style="display:inline"><input type="hidden" name="stage" value="crop" /><input class="btn" type="submit" value="2. Crop selected ROIs" /></form>
<form method="post" action="/configure" style="display:inline"><input type="hidden" name="crop_root" value="{html_escape(state.get('crop_root',''))}" /><input type="hidden" name="output_root" value="{html_escape(state.get('output_root',''))}" /><input class="btn" type="submit" value="3. Load crops for review" /></form>
<div class="small" style="margin-top:8px">Protected scripts are read only. The app runs generated copies from this app folder's generated_runners directory.</div>
</div>
<div class="panel">
<h2>Review Existing Crops</h2>
<form method="post" action="/configure_names">
<div class="small">Use this if preprocessing/cropping has already been done and you only want to load crops for review.</div>
<label>Crop review folder name inside the preprocessor folder</label><br>
<input type="text" name="crop_folder_name" value="{html_escape(state.get('crop_folder_name','03_cropped_rois_for_final_choice'))}" />
<label>Combined results folder name inside the raw image folder</label><br>
<input type="text" name="final_output_folder_name" value="{html_escape(state.get('final_output_folder_name','REAL FIJI PCC MANDERS RESULTS'))}" />
	<div class="notice"><b>Review crop set:</b> 01 raw masked for JACoP, original intensities.</div>
<input class="btn" type="submit" value="Use these folder names and scan crops" />
</form>
<form method="post" action="/recover_review_state" style="margin-top:10px">
<input class="btn secondary" type="submit" value="Recover Keep/Reject from previous results" />
<div class="small">Use this after choosing the raw folder and crop set. It rebuilds KEEP/REJECT decisions from old approved-crop folders or final Excel files in the current results folder.</div>
</form>
<form method="post" action="/load_previous_job_folder" style="margin-top:14px">
<label>Load any previous job/results folder</label><br>
<input type="text" name="previous_job_folder" value="{html_escape(state.get('last_previous_job_folder',''))}" placeholder="/path/to/raw folder, result folder, raw_masked_analysis, or crop folder" />
<input class="btn" type="submit" value="Load previous job from this folder" />
<div class="small">Use this to recover a previous job from any recognized folder in that experiment tree. Recovery is owned by the experiment folder through <b>{JOB_STATE_FILENAME}</b>, not by the app folder. It restores stage/settings data, crops, KEEP/REJECT decisions, manual recrops, cleaning parameters, thresholds, fixed-cell drawings, partition drawings, channel labels, source selections, and existing PCC/Manders, count/size, fixed-cell, and partition output links.</div>
</form>
<details>
<summary>Advanced: use full paths instead</summary>
<form method="post" action="/configure">
<label>Crop selected ROI folder full path</label><br>
<input type="text" name="crop_root" value="{html_escape(state.get('crop_root',''))}" />
<label>Final output folder full path</label><br>
<input type="text" name="output_root" value="{html_escape(state.get('output_root',''))}" />
<input class="btn secondary" type="submit" value="Save full paths and scan crops" />
</form>
</details>
</div>
	<div class="panel">
	<h2>Raw Crops Used For Review</h2>
	<div class="notice"><b>Current review-analysis choice:</b> {html_escape(current_choice_label)}<br>
	Raw crop folder: {html_escape(raw_crop_root)}</div>
	<div class="small">The enhanced visual-QC crop option has been removed from this auto-cleaned app copy. Final analysis can use either the original reviewed crop or a derived auto-cleaned crop, chosen separately for PCC/Manders and count/size.</div>
	<p><b>{len(raw_crop_items)} crops</b></p>
	<div id="rawChoiceGrid" class="grid">{raw_choice_cards_html}</div>
	{raw_choice_more_html}
	</div>
<div class="panel">
<b>Loaded crops:</b> {len(items)} &nbsp; <b>Kept:</b> <span id="kept-count">{kept}</span> &nbsp; <b>Rejected:</b> <span id="rejected-count">{rejected}</span> &nbsp; <b>Last scan:</b> {html_escape(state.get('last_scan',''))}<br>
<form id="rangeFilterForm" method="post" action="/select_crop_numbers" class="number-select-form">
  <label>Analyze crop # from<br><input type="number" name="crop_start" min="1" value="1" /></label>
  <label>to<br><input type="number" name="crop_end" min="1" value="{len(items)}" /></label>
  <button class="btn secondary" type="submit">Apply range filter</button>
  <div id="rangeFilterStatus" class="small">Keeps your current KEEP/REJECT choices inside this range. Crops outside the range are excluded.</div>
</form>
<form id="analysis-labels" method="post" action="/analyze_pcc_manders">
<div class="notice">
	<b>Final analysis settings:</b> choose the shared input image and default channel roles. Each crop card can override these roles for PCC/Manders and count/size.
	<div class="form-grid" style="margin-top:10px">
		<label>Input image for PCC/Manders and count/size<br>
		<select name="analysis_input_mode" data-shared-analysis-input-mode>
		<option value="raw" {'selected' if normalize_analysis_input_mode(state.get('analysis_input_mode', 'raw')) == 'raw' else ''}>Raw cropped image</option>
		<option value="auto_cleaned" {'selected' if normalize_analysis_input_mode(state.get('analysis_input_mode', 'raw')) == 'auto_cleaned' else ''}>Auto-cleaned image</option>
		</select></label>
		<div id="analysisInputSaveStatus" class="small">This one choice is used by PCC/Manders and count/size.</div>
		<label>PCC/Manders channel pair<br>
	<select name="pcc_pair_mode">
	<option value="red_green" {'selected' if state.get('pcc_pair_mode', 'red_green') == 'red_green' else ''}>Compare red vs green</option>
	<option value="red_blue" {'selected' if state.get('pcc_pair_mode', 'red_green') == 'red_blue' else ''}>Compare red vs blue</option>
	<option value="green_blue" {'selected' if state.get('pcc_pair_mode', 'red_green') == 'green_blue' else ''}>Compare blue vs green</option>
	<option value="all_three" {'selected' if state.get('pcc_pair_mode', 'red_green') == 'all_three' else ''}>Compare all three pairs</option>
		</select><br><span class="small">PCC/Manders will only use the selected channel pair; per-crop roles only decide the pair type/order.</span></label>
	<label>Red channel role<br>
	<select name="red_channel_role">
	<option value="condensate" {'selected' if state.get('red_channel_role', 'diffuse_marker') == 'condensate' else ''}>Condensate / puncta</option>
	<option value="diffuse_marker" {'selected' if state.get('red_channel_role', 'diffuse_marker') == 'diffuse_marker' else ''}>Marker / protein signal</option>
	<option value="nucleus" {'selected' if state.get('red_channel_role', 'diffuse_marker') == 'nucleus' else ''}>Nucleus / reference</option>
	<option value="ignore" {'selected' if state.get('red_channel_role', 'diffuse_marker') == 'ignore' else ''}>Ignore</option>
	</select></label>
	<label>Green channel role<br>
	<select name="green_channel_role">
	<option value="condensate" {'selected' if state.get('green_channel_role', 'condensate') == 'condensate' else ''}>Condensate / puncta</option>
	<option value="diffuse_marker" {'selected' if state.get('green_channel_role', 'condensate') == 'diffuse_marker' else ''}>Marker / protein signal</option>
	<option value="nucleus" {'selected' if state.get('green_channel_role', 'condensate') == 'nucleus' else ''}>Nucleus / reference</option>
	<option value="ignore" {'selected' if state.get('green_channel_role', 'condensate') == 'ignore' else ''}>Ignore</option>
	</select></label>
	<label>Blue channel role<br>
	<select name="blue_channel_role">
	<option value="condensate" {'selected' if state.get('blue_channel_role', 'nucleus') == 'condensate' else ''}>Condensate / puncta</option>
	<option value="diffuse_marker" {'selected' if state.get('blue_channel_role', 'nucleus') == 'diffuse_marker' else ''}>Marker / protein signal</option>
	<option value="nucleus" {'selected' if state.get('blue_channel_role', 'nucleus') == 'nucleus' else ''}>Nucleus / reference</option>
	<option value="ignore" {'selected' if state.get('blue_channel_role', 'nucleus') == 'ignore' else ''}>Ignore</option>
	</select></label>
	</div>
	<div class="small">PCC/Manders can compare any selected channel pair, including condensate vs condensate. In channel-cleaned mode, every channel marked condensate/puncta is puncta-cleaned; marker/protein, nucleus/reference, and ignored channels stay raw. Original raw images and reviewed crops are not overwritten.</div>
	<div class="form-grid" style="margin-top:10px">
	<label>Red channel is<br><input type="text" name="pcc_red_label" placeholder="Example: N1ICD" value="{html_escape(state.get('pcc_red_label',''))}" /></label>
<label>Green channel is<br><input type="text" name="pcc_green_label" placeholder="Example: H3K27ac" value="{html_escape(state.get('pcc_green_label',''))}" /></label>
<label>Blue channel is<br><input type="text" name="pcc_blue_label" placeholder="Example: DAPI" value="{html_escape(state.get('pcc_blue_label',''))}" /></label>
</div>
<h3>Condensate count/size</h3>
<div class="form-grid" style="margin-top:10px">
<label><input type="checkbox" name="count_size_red" value="1" {'checked' if state.get('count_size_red', True) else ''} /> Count/size Red channel</label>
<label><input type="checkbox" name="count_size_green" value="1" {'checked' if state.get('count_size_green', True) else ''} /> Count/size Green channel</label>
<label><input type="checkbox" name="count_size_blue" value="1" {'checked' if state.get('count_size_blue', True) else ''} /> Count/size Blue channel</label>
<label>Fiji particle size (px²)<br><input type="text" name="count_size_particle_size" value="{html_escape(normalize_count_size_particle_size(state.get('count_size_particle_size', '10-Infinity'))[1])}" placeholder="10-Infinity" inputmode="decimal" /></label>
</div>
	<div class="small">Count/size uses the same final analysis input selected above and only measures channels marked condensate/puncta for each crop. Fiji particle size is sent directly to Analyze Particles; for example, 10-Infinity excludes particles smaller than 10 px². If channel-cleaned copy is selected, Fiji count/size runs on the cleaned kept crop channels and converts area from px² to µm² using the pixel size above.</div>
<div class="panel" style="margin-top:14px; background:#f8fbff">
  <h3>Fixed-cell MS2/condensate measurements</h3>
  <div class="small">Choose a two-color core experiment, a three-color experiment with DAPI, or a three-color experiment with a modifier such as HP1β. The primary condensate/protein can be N1ICD or any other marker you name. Segmentation may use the selected raw/cleaned count-size input, but fluorescence intensity is always measured from the original reviewed raw crop.</div>
  <div class="notice"><b>Manual annotations saved:</b> {fixed_cell_annotated_kept} of {kept} kept crops. Manual-required analysis reports an error for every kept crop that has not been annotated.</div>
  <div class="form-grid" style="margin-top:10px">
    <label>Color mode<br><select name="fixed_cell_color_mode">
      <option value="two_color" {'selected' if state.get('fixed_cell_color_mode', 'three_color_dapi') == 'two_color' else ''}>Two color: condensate/protein + MS2/MCP</option>
      <option value="three_color_dapi" {'selected' if state.get('fixed_cell_color_mode', 'three_color_dapi') == 'three_color_dapi' else ''}>Three color: condensate/protein + MS2/MCP + DAPI</option>
      <option value="three_color_modifier" {'selected' if state.get('fixed_cell_color_mode', 'three_color_dapi') == 'three_color_modifier' else ''}>Three color: condensate/protein + MS2/MCP + modifier</option>
    </select></label>
    <label>ROI and MCP/MS2 focus source<br><select name="fixed_cell_annotation_mode">
      <option value="manual_required" {'selected' if state.get('fixed_cell_annotation_mode', 'manual_required') == 'manual_required' else ''}>Require my manual ROI and focus drawings</option>
      <option value="prefer_manual" {'selected' if state.get('fixed_cell_annotation_mode', 'manual_required') == 'prefer_manual' else ''}>Use manual drawings when saved; otherwise automatic</option>
      <option value="automatic" {'selected' if state.get('fixed_cell_annotation_mode', 'manual_required') == 'automatic' else ''}>Automatic ROI and focus detection</option>
    </select></label>
    <label>Fixed-cell intensity source<br><select name="fixed_cell_intensity_source_mode">
      <option value="reviewed_raw" {'selected' if normalize_fixed_cell_intensity_source_mode(state.get('fixed_cell_intensity_source_mode')) == 'reviewed_raw' else ''}>Reviewed raw/raw-masked source (current method)</option>
      <option value="untouched_acquisition" {'selected' if normalize_fixed_cell_intensity_source_mode(state.get('fixed_cell_intensity_source_mode')) == 'untouched_acquisition' else ''}>Matching untouched acquisition TIFF</option>
    </select></label>
    <label>Fixed-cell condensate source<br><select name="fixed_cell_condensate_source">
      <option value="count_size" {'selected' if state.get('fixed_cell_condensate_source', 'count_size') == 'count_size' else ''}>Reuse intact Fiji count/size particles</option>
      <option value="manual" {'selected' if state.get('fixed_cell_condensate_source', 'count_size') == 'manual' else ''}>Use my manually traced condensates</option>
    </select></label>
    <label>Primary condensate/protein channel<br><select name="fixed_cell_n1icd_channel">
      <option value="red" {'selected' if state.get('fixed_cell_n1icd_channel', 'green') == 'red' else ''}>Red</option>
      <option value="green" {'selected' if state.get('fixed_cell_n1icd_channel', 'green') == 'green' else ''}>Green</option>
      <option value="blue" {'selected' if state.get('fixed_cell_n1icd_channel', 'green') == 'blue' else ''}>Blue</option>
    </select></label>
    <label>Primary condensate/protein label<br><input type="text" name="fixed_cell_condensate_label" value="{html_escape(state.get('fixed_cell_condensate_label', 'N1ICD'))}" placeholder="Example: N1ICD, NICD, HP1beta" /></label>
    <label>MS2/MCP channel<br><select name="fixed_cell_ms2_channel">
      <option value="red" {'selected' if state.get('fixed_cell_ms2_channel', 'red') == 'red' else ''}>Red</option>
      <option value="green" {'selected' if state.get('fixed_cell_ms2_channel', 'red') == 'green' else ''}>Green</option>
      <option value="blue" {'selected' if state.get('fixed_cell_ms2_channel', 'red') == 'blue' else ''}>Blue</option>
    </select></label>
    <label>MCP/readout protein label<br><input type="text" name="fixed_cell_ms2_label" value="{html_escape(state.get('fixed_cell_ms2_label', 'MCP/MS2'))}" placeholder="Example: MCP, MS2, protein X" /></label>
    <label>DAPI channel<br><select name="fixed_cell_dapi_channel">
      <option value="red" {'selected' if state.get('fixed_cell_dapi_channel', 'blue') == 'red' else ''}>Red</option>
      <option value="green" {'selected' if state.get('fixed_cell_dapi_channel', 'blue') == 'green' else ''}>Green</option>
      <option value="blue" {'selected' if state.get('fixed_cell_dapi_channel', 'blue') == 'blue' else ''}>Blue</option>
    </select></label>
    <label>Optional modifier channel<br><select name="fixed_cell_modifier_channel">
      <option value="red" {'selected' if state.get('fixed_cell_modifier_channel', 'blue') == 'red' else ''}>Red</option>
      <option value="green" {'selected' if state.get('fixed_cell_modifier_channel', 'blue') == 'green' else ''}>Green</option>
      <option value="blue" {'selected' if state.get('fixed_cell_modifier_channel', 'blue') == 'blue' else ''}>Blue</option>
    </select></label>
    <label>Modifier label<br><input type="text" name="fixed_cell_modifier_label" value="{html_escape(state.get('fixed_cell_modifier_label', 'HP1beta'))}" placeholder="Example: HP1beta" /></label>
    <label>MS2 focus rule<br><select name="fixed_cell_focus_rule">
      <option value="brightest" {'selected' if state.get('fixed_cell_focus_rule', 'brightest') == 'brightest' else ''}>Brightest focus only</option>
      <option value="sum" {'selected' if state.get('fixed_cell_focus_rule', 'brightest') == 'sum' else ''}>Sum all detected foci</option>
    </select></label>
    <label>Minimum MS2 focus area (px²)<br><input type="number" name="fixed_cell_ms2_min_area_px2" min="1" step="1" value="{clamp_int(state.get('fixed_cell_ms2_min_area_px2'), 3, 1, 10000)}" /></label>
    <label>Background annulus inner radius (px)<br><input type="number" name="fixed_cell_background_inner_px" min="1" step="1" value="{clamp_int(state.get('fixed_cell_background_inner_px'), 3, 1, 100)}" /></label>
    <label>Background annulus outer radius (px)<br><input type="number" name="fixed_cell_background_outer_px" min="2" step="1" value="{clamp_int(state.get('fixed_cell_background_outer_px'), 12, 2, 500)}" /></label>
    <label>Replicate ID<br><input type="text" name="fixed_cell_replicate_id" value="{html_escape(state.get('fixed_cell_replicate_id', ''))}" placeholder="Example: Replicate 1" /></label>
    <label>Condition<br><input type="text" name="fixed_cell_condition" value="{html_escape(state.get('fixed_cell_condition', ''))}" placeholder="Example: HES1 WT" /></label>
  </div>
  <div class="small">In manual-required mode, use <b>Draw ROI + MCP focus + condensates</b> on every kept crop. The selected MCP/readout channel may represent MCP/MS2 or another protein; enter its biological name above. The fixed-cell workbook reports mean intensity and integrated density for Red, Green, and Blue in the measurement ROI, selected focus, background ROI, nearest condensate, and each condensate particle. Choose whether fixed-cell condensates come from intact Fiji count/size particles assigned by centroid or from condensate polygons you trace yourself. The untouched-acquisition option changes <b>only fixed-cell fluorescence intensity measurements</b>; PCC/Manders and count/size remain on their existing selected inputs. The app requires an exact matching TIFF and validates dimensions, coordinates, alignment, and RGB channel order before measuring. If you draw multiple MCP foci, the focus rule above decides whether to use the brightest drawing or sum all drawings. After analysis, inspect every color-coded overlay and the QC Summary worksheet: cyan = ROI, green = accepted condensates, orange = MS2 background, magenta = selected focus, yellow = nearest condensate.</div>
</div>
<div class="panel" style="margin-top:14px; background:#f4fff8; border-color:#86c99b">
  <h3>Manual measurement-ROI partition coefficient</h3>
  <div class="small"><b>Independent workflow:</b> Draw one outer measurement/reference ROI and the condensates defined by one protein. A = mean channel intensity inside the condensate ROIs. B = mean of that same channel in the outer ROI after excluding those condensates. Partition coefficient = A / B. For two proteins + DAPI, trace the outer ROI using DAPI. For three proteins without DAPI, trace it on the merged view or whichever protein channel shows the desired cell/reference boundary. A new crop is not required.</div>
  <div class="notice"><b>Partition annotations saved:</b> {partition_annotated_kept} of {kept} kept cells. The first time you open this editor, existing fixed-cell nuclear/condensate drawings are offered as a non-destructive starting copy when available.</div>
  <div class="form-grid" style="margin-top:10px">
    <label>Primary result to highlight<br><select name="partition_target_channel">
      <option value="red" {'selected' if state.get('partition_target_channel', 'red') == 'red' else ''}>Red</option>
      <option value="green" {'selected' if state.get('partition_target_channel', 'red') == 'green' else ''}>Green</option>
      <option value="blue" {'selected' if state.get('partition_target_channel', 'red') == 'blue' else ''}>Blue</option>
    </select></label>
    <label>Primary highlighted protein label<br><input type="text" name="partition_target_label" value="{html_escape(state.get('partition_target_label', 'UB-HA'))}" placeholder="Example: UB-HA" /></label>
    <label>Red channel protein label<br><input type="text" name="partition_red_label" value="{html_escape(state.get('partition_red_label', ''))}" placeholder="Example: NICD" /></label>
    <label>Green channel protein label<br><input type="text" name="partition_green_label" value="{html_escape(state.get('partition_green_label', ''))}" placeholder="Example: UB-HA" /></label>
    <label>Blue channel protein label<br><input type="text" name="partition_blue_label" value="{html_escape(state.get('partition_blue_label', ''))}" placeholder="Example: RNAPol II or DAPI" /></label>
    <label>Condensate label<br><input type="text" name="partition_condensate_label" value="{html_escape(state.get('partition_condensate_label', 'NICD'))}" placeholder="Example: NICD" /></label>
    <label>Channel defining condensate ROIs<br><select name="partition_condensate_channel">
      <option value="red" {'selected' if state.get('partition_condensate_channel', 'red') == 'red' else ''}>Red</option>
      <option value="green" {'selected' if state.get('partition_condensate_channel', 'red') == 'green' else ''}>Green</option>
      <option value="blue" {'selected' if state.get('partition_condensate_channel', 'red') == 'blue' else ''}>Blue</option>
    </select></label>
    <label>Partition intensity source<br><select name="partition_intensity_source_mode">
      <option value="untouched_acquisition" {'selected' if normalize_fixed_cell_intensity_source_mode(state.get('partition_intensity_source_mode', 'untouched_acquisition')) == 'untouched_acquisition' else ''}>Matching untouched acquisition TIFF (recommended)</option>
      <option value="reviewed_raw" {'selected' if normalize_fixed_cell_intensity_source_mode(state.get('partition_intensity_source_mode', 'untouched_acquisition')) == 'reviewed_raw' else ''}>Reviewed raw/raw-masked source</option>
    </select></label>
    <label>Condensate edge exclusion (px)<br><input type="number" name="partition_boundary_exclusion_px" min="0" max="10" step="1" value="{clamp_int(state.get('partition_boundary_exclusion_px'), 1, 0, 10)}" /></label>
    <label>Replicate ID<br><input type="text" name="partition_replicate_id" value="{html_escape(state.get('partition_replicate_id', ''))}" placeholder="Example: Replicate 1" /></label>
    <label>Condition<br><input type="text" name="partition_condition" value="{html_escape(state.get('partition_condition', ''))}" placeholder="Example: Control" /></label>
  </div>
  <div class="small">All three channels are measured independently in the same masks and receive names based on the protein labels you enter. The app retains raw ratios and also reports background-corrected partition coefficients: (condensate mean − cell-free background) / (dilute-phase mean − cell-free background). The diffuse protein inside the cell remains the denominator. The edge exclusion removes boundary pixels from both phases. This section does not call or modify PCC/Manders, count/size, or fixed-cell MS2 analysis.</div>
</div>
<div class="threshold-panel">
  <h3>Manual final-analysis threshold for count/size</h3>
  <div class="threshold-controls">
    <button id="thresholdPrevCrop" class="btn secondary" type="button">Previous crop</button>
    <div id="thresholdCropDisplay" class="threshold-crop-display">Kept crop</div>
    <label>Jump to kept crop<br><select id="thresholdCropRel">{threshold_crop_options}</select></label>
    <button id="refreshThresholdCrops" class="btn secondary" type="button">Refresh count/size crop list</button>
    <button id="thresholdNextCrop" class="btn secondary" type="button">Next crop</button>
    <label>Channel<br><select id="thresholdChannel"><option value="red">Red</option><option value="green">Green</option><option value="blue">Blue</option></select></label>
    <label>Min<br><input id="thresholdMinSlider" type="range" min="0" max="255" step="1" value="0" list="thresholdTicks"></label>
    <input id="thresholdMinValue" type="number" min="0" max="255" step="1" value="0">
    <label>Max<br><input id="thresholdMaxSlider" type="range" min="0" max="255" step="1" value="255" list="thresholdTicks"></label>
    <input id="thresholdMaxValue" type="number" min="0" max="255" step="1" value="255">
    <datalist id="thresholdTicks"><option value="0"></option><option value="64"></option><option value="128"></option><option value="192"></option><option value="255"></option></datalist>
    <button id="fijiAutoThreshold" class="btn secondary" type="button">Fiji Auto</button>
    <button id="resetThreshold" class="btn secondary" type="button">Reset</button>
    <button id="saveThreshold" class="btn" type="button">Save threshold for this crop</button>
  </div>
  <div class="small" id="thresholdStatus">Use Previous/Next and choose a channel to load Fiji Default dark threshold.</div>
  <div class="threshold-histogram-wrap">
    <canvas id="thresholdHistogram" width="720" height="96"></canvas>
    <div class="threshold-scale"><b>Fiji 8-bit histogram:</b><span>0</span><span>64</span><span>128</span><span>192</span><span>255</span></div>
  </div>
  <div class="small">Move the sliders or type exact Fiji Min/Max values. The live preview and final count/size analysis send those exact numbers back to Fiji before Convert to Mask.</div>
  <div class="threshold-preview-grid">
    <div class="threshold-preview-pane"><h3>Original selected-channel crop</h3><div id="thresholdOriginalPreview" class="threshold-preview-box">Original selected channel</div></div>
    <div class="threshold-preview-pane"><h3>Fiji 8-bit threshold selection</h3><div id="threshold8bitPreview" class="threshold-preview-box">Fiji 8-bit threshold selection</div></div>
    <div class="threshold-preview-pane"><h3>Fiji watershed mask</h3><div id="thresholdWatershedPreview" class="threshold-preview-box">Fiji watershed mask</div></div>
  </div>
</div>
</div>
<button class="btn secondary" type="submit" formaction="/preview_channel_cleaning">Autoclean kept crops</button>
<button class="btn" type="submit" formaction="/analyze_pcc_manders">Run Real Fiji PCC/Manders only</button>
<button class="btn" type="submit" formaction="/analyze_count_size">Run count/size only</button>
<button class="btn" type="submit" formaction="/analyze_pcc_manders_count_size">Run PCC/Manders + count/size</button>
<div id="countSizeProgressPanel" class="small" style="display:none; margin-top:12px; max-width:760px">
  <b>Analysis progress:</b> <span id="countSizeProgressText">Starting…</span><br>
  <progress id="countSizeProgressBar" value="0" max="100" style="width:100%; height:18px"></progress>
</div>

<button class="btn" type="submit" formaction="/analyze_fixed_cell_ms2">Run fixed-cell MS2/condensate measurements</button>
<button class="btn" type="submit" formaction="/analyze_manual_partition">Run manual partition coefficient</button>
<a class="btn secondary" href="#final-previews" style="display:inline-block; text-decoration:none">Show latest analysis previews</a>
<a class="btn secondary" href="#fixed-cell-qc-previews" style="display:inline-block; text-decoration:none">Show fixed-cell QC overlays</a>
</form>
<form method="post" action="/all_keep" style="display:inline"><input class="btn secondary" type="submit" value="Set all to keep" /></form>
<div class="small">All analysis buttons use the same KEEP/REJECT decisions, selected crop set, channel-cleaned/original input choice, and channel labels. PCC/Manders sends RGB crops to Fiji/ImageJ, lets Fiji split channels, and computes the exported PCC/Manders metrics without fragile Java reflection. Count/size sends selected channels to Fiji Analyze Particles and reports px² plus converted µm² size values.</div>
</div>
{f'<div class="panel"><h2>Last Pipeline Log</h2><pre style="white-space:pre-wrap; max-height:320px; overflow:auto">{html_escape(state.get("last_stage_output", ""))}</pre></div>' if state.get('last_stage_output') else ''}
{latest_outputs_html}
{recovered_panel_html}
	<div class="panel">
	<h2>Raw vs cleaned image previews</h2>
	<div class="small">These previews show the original reviewed crop channels beside the derived cleaned channels used for analysis. Preview folder: {html_escape(cleaning_preview_root)}</div>
	<form id="globalCleaningPreviewForm" class="global-cleaning-preview-form" method="post" action="/api/preview_global_channel_cleaning">
	<h3>Global raw-vs-cleaned preview sliders</h3>
	<div class="form-grid" style="margin-top:10px">
	<label>Red raw to lowest background: <b><span data-global-value-for="red_cleaning_amount">{clamp_int(state.get('red_cleaning_amount'), 100, 0, 100)}</span>%</b><br><input class="global-cleaning-slider" type="range" name="red_cleaning_amount" min="0" max="100" step="1" value="{clamp_int(state.get('red_cleaning_amount'), 100, 0, 100)}" /></label>
	<label>Green raw to lowest background: <b><span data-global-value-for="green_cleaning_amount">{clamp_int(state.get('green_cleaning_amount'), 100, 0, 100)}</span>%</b><br><input class="global-cleaning-slider" type="range" name="green_cleaning_amount" min="0" max="100" step="1" value="{clamp_int(state.get('green_cleaning_amount'), 100, 0, 100)}" /></label>
	<label>Blue raw to lowest background: <b><span data-global-value-for="blue_cleaning_amount">{clamp_int(state.get('blue_cleaning_amount'), 100, 0, 100)}</span>%</b><br><input class="global-cleaning-slider" type="range" name="blue_cleaning_amount" min="0" max="100" step="1" value="{clamp_int(state.get('blue_cleaning_amount'), 100, 0, 100)}" /></label>
	<label>Background kept: <b><span data-global-value-for="cleaning_background_removal">{clamp_int(state.get('cleaning_background_removal'), 0, 0, 100)}</span>%</b><br><input class="global-cleaning-slider" type="range" name="cleaning_background_removal" min="0" max="100" step="1" value="{clamp_int(state.get('cleaning_background_removal'), 0, 0, 100)}" /></label>
	<label>Shared radius: <b><span data-global-value-for="cleaning_radius_px">{clamp_int(state.get('cleaning_radius_px'), 14, 4, 40)}</span> px</b><br><input class="global-cleaning-slider" type="range" name="cleaning_radius_px" min="4" max="40" step="1" value="{clamp_int(state.get('cleaning_radius_px'), 14, 4, 40)}" /></label>
	<label>Shared sensitivity: <b><span data-global-value-for="cleaning_sensitivity">{clamp_int(state.get('cleaning_sensitivity'), 50, 1, 100)}</span></b><br><input class="global-cleaning-slider" type="range" name="cleaning_sensitivity" min="1" max="100" step="1" value="{clamp_int(state.get('cleaning_sensitivity'), 50, 1, 100)}" /></label>
	<label>Shared minimum puncta size: <b><span data-global-value-for="cleaning_min_area_px">{clamp_int(state.get('cleaning_min_area_px'), 3, 1, 100)}</span> px</b><br><input class="global-cleaning-slider" type="range" name="cleaning_min_area_px" min="1" max="100" step="1" value="{clamp_int(state.get('cleaning_min_area_px'), 3, 1, 100)}" /></label>
	</div>
	<div class="small">Moving these sliders regenerates the kept raw-vs-cleaned preview cards from the global settings. Per-crop overrides are cleared for kept crops when global sliders are applied.</div>
	</form>
	<div style="margin:10px 0">
	  <button id="applyGlobalCleaningButton" class="btn secondary" type="button">Autoclean kept crops</button>
	  <button id="resetAllCleaningButton" class="btn secondary" type="button">Reset all previews to global settings</button>
	  <span id="resetAllCleaningStatus" class="small"></span>
	</div>
	<div id="cleaningPreviewGrid" class="grid">{cleaning_cards_html}</div>
	<div id="cleaningShowMoreHolder">{cleaning_more_html}</div>
	</div>
<div id="final-previews" class="panel">
<h2>Latest PCC/count-size preview images</h2>
<div class="small">Preview PNG folder: {html_escape(final_preview_root)}<br>Approved TIFF folder: {html_escape(approved_root)}</div>
<div class="grid">{approved_cards_html}</div>
</div>
<div id="fixed-cell-qc-previews" class="panel">
<h2>Fixed-cell MS2/condensate QC overlays</h2>
{f'<div class="notice"><b>{fixed_cell_pending_qc} manual annotation(s) changed.</b> The cards currently show immediate drawing-only previews. Run <b>fixed-cell MS2/condensate measurements</b> to generate the complete condensate, background, and nearest-condensate QC layers.</div>' if fixed_cell_pending_qc else ''}
<div class="small">QC overlay folder: {html_escape(fixed_cell_preview_root or 'not run yet')}<br>Immediately after drawing: cyan = measurement ROI and magenta = saved MCP/MS2 focus. After analysis: green = accepted condensates, orange = MS2 background, and yellow = nearest condensate.</div>
<div class="grid">{fixed_cell_cards_html}</div>
</div>
<h2 id="crop-review">Crop review</h2>
<div class="review-filter-bar" aria-label="Crop review filters">
  <button type="button" class="active" data-review-filter="all">Show all ({len(items)})</button>
  <button type="button" data-review-filter="keep">Show kept ({kept})</button>
  <button type="button" data-review-filter="reject">Show rejected ({rejected})</button>
</div>
<div class="grid">{cards_html}</div>
<iframe name="review_sink" style="display:none"></iframe>
<script>
async function uploadRawFolder(event) {{
  event.preventDefault();
  const input = document.getElementById('rawFolderInput');
  const status = document.getElementById('uploadStatus');
  if (!input.files || input.files.length === 0) {{
    status.textContent = 'Choose a raw image folder first.';
    return;
  }}
  const totalBytes = Array.from(input.files).reduce((sum, file) => sum + Number(file.size || 0), 0);
  const maxBytes = 256 * 1024 * 1024;
  if (totalBytes > maxBytes) {{
    status.textContent = 'This folder is ' + (totalBytes / (1024 * 1024)).toFixed(1) + ' MB. Browser copy is limited to 256 MB. Paste the original folder path below; no files need to be copied.';
    return;
  }}
  const data = new FormData();
  for (const file of input.files) {{
    data.append('files', file, file.webkitRelativePath || file.name);
  }}
  status.textContent = 'Uploading folder... keep this browser tab open.';
  try {{
    const response = await fetch('/upload_raw_folder', {{ method: 'POST', body: data }});
    const html = await response.text();
    document.open();
    document.write(html);
    document.close();
  }} catch (error) {{
    status.textContent = 'Upload failed: ' + error;
  }}
}}
const uploadForm = document.getElementById('uploadRawFolderForm');
if (uploadForm) uploadForm.addEventListener('submit', uploadRawFolder);
const rawFolderInput = document.getElementById('rawFolderInput');
if (rawFolderInput) rawFolderInput.addEventListener('change', () => {{
  const status = document.getElementById('uploadStatus');
  const count = rawFolderInput.files ? rawFolderInput.files.length : 0;
  const totalBytes = rawFolderInput.files ? Array.from(rawFolderInput.files).reduce((sum, file) => sum + Number(file.size || 0), 0) : 0;
  if (status) status.textContent = count
    ? count + ' files selected (' + (totalBytes / (1024 * 1024)).toFixed(1) + ' MB). ' + (totalBytes > 256 * 1024 * 1024 ? 'Too large for browser copy—paste the original folder path below.' : 'Click Upload and use selected folder.')
    : 'No folder selected.';
}});

const sharedAnalysisInput = document.querySelector('[data-shared-analysis-input-mode]');
if (sharedAnalysisInput) sharedAnalysisInput.addEventListener('change', async () => {{
  const status = document.getElementById('analysisInputSaveStatus');
  const label = sharedAnalysisInput.value === 'auto_cleaned' ? 'Auto-cleaned image' : 'Raw cropped image';
  if (status) status.textContent = 'Saving shared input choice...';
  const body = new URLSearchParams();
  body.set('analysis_input_mode', sharedAnalysisInput.value);
  try {{
    const response = await fetch('/api/set_analysis_input_mode', {{ method: 'POST', body }});
    const payload = await response.json().catch(() => null);
    if (!response.ok || !payload || !payload.ok) throw new Error(payload && payload.error ? payload.error : 'Save failed');
    if (status) status.textContent = 'Saved. PCC/Manders and count/size will use: ' + label + '.';
    if (thresholdEl('thresholdCropRel')) loadFijiThreshold(false);
  }} catch (error) {{
    if (status) status.textContent = 'Input choice save failed: ' + String(error.message || error);
  }}
}});

document.addEventListener('click', (event) => {{
  const button = event.target.closest('.auxiliary-show-all');
  if (!button) return;
  if (button.dataset.section === 'cleaning') {{
    refreshCleaningPreviewSection(true);
    return;
  }}
  const target = document.getElementById(button.dataset.target || '');
  if (!target) return;
  target.querySelectorAll('.auxiliary-extra').forEach((card) => {{
    card.classList.remove('auxiliary-extra');
  }});
  button.remove();
}});

let thresholdHistogram = [];
let thresholdPreviewTimer = null;
let thresholdPreviewRequest = 0;
let thresholdStateRequest = 0;
let thresholdDirty = false;
let thresholdLoadedRel = '';
let thresholdLoadedChannel = 'red';
const thresholdSessionValues = new Map();
function thresholdSessionKey(rel, channel) {{ return `${{rel || ''}}|||${{channel || 'red'}}`; }}
function thresholdEl(id) {{ return document.getElementById(id); }}
function clampThreshold(value, fallback) {{
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.max(0, Math.min(255, Math.round(number)));
}}
function setThresholdControls(minValue, maxValue) {{
  let lower = clampThreshold(minValue, 0);
  let upper = clampThreshold(maxValue, 255);
  if (upper < lower) {{ const tmp = lower; lower = upper; upper = tmp; }}
  thresholdEl('thresholdMinSlider').value = lower;
  thresholdEl('thresholdMinValue').value = lower;
  thresholdEl('thresholdMaxSlider').value = upper;
  thresholdEl('thresholdMaxValue').value = upper;
  drawThresholdHistogram(lower, upper);
  return {{ lower, upper }};
}}
function drawThresholdHistogram(lower, upper) {{
  const canvas = thresholdEl('thresholdHistogram');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const w = canvas.width;
  const h = canvas.height;
  ctx.clearRect(0, 0, w, h);
  ctx.fillStyle = '#050506';
  ctx.fillRect(0, 0, w, h);
  const hist = thresholdHistogram || [];
  const maxCount = Math.max(1, ...hist);
  for (let i = 0; i < 256; i++) {{
    const x0 = Math.floor(i * w / 256);
    const x1 = Math.max(x0 + 1, Math.floor((i + 1) * w / 256));
    const barHeight = Math.round((hist[i] || 0) / maxCount * (h - 18));
    ctx.fillStyle = i >= lower && i <= upper ? '#64a6ff' : '#343946';
    ctx.fillRect(x0, h - barHeight - 9, x1 - x0, barHeight);
  }}
  ctx.fillStyle = '#f4f7ff';
  ctx.fillRect(Math.round(lower * w / 255), 2, 2, h - 4);
  ctx.fillRect(Math.round(upper * w / 255), 2, 2, h - 4);
}}
function selectedThresholdCrop() {{
  const select = thresholdEl('thresholdCropRel');
  return select ? select.value : '';
}}
function updateThresholdCropDisplay() {{
  const select = thresholdEl('thresholdCropRel');
  const display = thresholdEl('thresholdCropDisplay');
  const prev = thresholdEl('thresholdPrevCrop');
  const next = thresholdEl('thresholdNextCrop');
  if (!select || !display) return;
  const total = select.options.length;
  if (!total) {{
    display.textContent = 'No kept crops available';
    if (prev) prev.disabled = true;
    if (next) next.disabled = true;
    return;
  }}
  const current = Math.max(0, select.selectedIndex);
  const label = select.options[current]?.text || select.value || 'Kept crop';
  display.textContent = `Crop ${{current + 1}} of ${{total}}: ${{label}}`;
  display.title = label;
  if (prev) prev.disabled = current <= 0;
  if (next) next.disabled = current >= total - 1;
}}
async function moveThresholdCrop(delta) {{
  const select = thresholdEl('thresholdCropRel');
  if (!select || !select.options.length) return;
  if (thresholdDirty && !(await saveCountSizeThreshold(true, thresholdLoadedRel, thresholdLoadedChannel))) return;
  const nextIndex = Math.max(0, Math.min(select.options.length - 1, select.selectedIndex + delta));
  if (nextIndex === select.selectedIndex) {{
    thresholdEl('thresholdStatus').textContent = delta < 0 ? 'Already at the first kept crop.' : 'Already at the last kept crop.';
    return;
  }}
  select.selectedIndex = nextIndex;
  updateThresholdCropDisplay();
  await loadFijiThreshold(false);
}}
function selectedThresholdChannel() {{
  const select = thresholdEl('thresholdChannel');
  return select ? select.value : 'red';
}}
function selectedSharedAnalysisInputMode() {{
  const select = document.querySelector('[data-shared-analysis-input-mode]');
  return select ? select.value : 'raw';
}}
async function loadFijiThreshold(forceAuto=false) {{
  const requestId = ++thresholdStateRequest;
  const rel = selectedThresholdCrop();
  const channel = selectedThresholdChannel();
  updateThresholdCropDisplay();
  if (!rel) {{
    thresholdEl('thresholdStatus').textContent = 'No kept crop selected.';
    return;
  }}
  thresholdHistogram = null;
  drawThresholdHistogram(0, 255);
  thresholdEl('thresholdStatus').textContent = 'asking Fiji for threshold...';
  try {{
    const mode = selectedSharedAnalysisInputMode();
    const response = await fetch(`/count_size_threshold_default?rel=${{encodeURIComponent(rel)}}&channel=${{encodeURIComponent(channel)}}&analysis_input_mode=${{encodeURIComponent(mode)}}&v=${{Date.now()}}`);
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || response.statusText);
    if (requestId !== thresholdStateRequest) return;
    thresholdHistogram = data.histogram || [];
    const sessionSaved = thresholdSessionValues.get(thresholdSessionKey(rel, channel));
    const hasServerSaved = data.saved_min !== null && data.saved_min !== undefined && data.saved_max !== null && data.saved_max !== undefined;
    const hasSaved = Boolean(sessionSaved) || hasServerSaved;
    const lower = hasSaved && !forceAuto ? (sessionSaved ? sessionSaved.min : data.saved_min) : data.threshold_min;
    const upper = hasSaved && !forceAuto ? (sessionSaved ? sessionSaved.max : data.saved_max) : data.threshold_max;
    const values = setThresholdControls(lower, upper);
    thresholdLoadedRel = rel;
    thresholdLoadedChannel = channel;
    thresholdDirty = false;
    thresholdEl('thresholdStatus').textContent = hasSaved && !forceAuto
      ? `saved ${{values.lower}}-${{values.upper}} | input: ${{data.analysis_input_mode || ''}}`
      : `Fiji Auto Default dark ${{values.lower}}-${{values.upper}} | input: ${{data.analysis_input_mode || ''}}`;
    renderThresholdPreview();
  }} catch (error) {{
    if (requestId !== thresholdStateRequest) return;
    thresholdHistogram = null;
    const values = setThresholdControls(0, 255);
    thresholdEl('thresholdStatus').textContent = `Fiji threshold probe failed: ${{error.message || 'unknown error'}} | ${{values.lower}}-${{values.upper}}`;
    renderThresholdPreview();
  }}
}}
async function refreshCountSizeThresholdPanel(forceAuto=false) {{
  if (!thresholdEl('thresholdCropRel')) return;
  const status = thresholdEl('thresholdStatus');
  if (status) status.textContent = 'Autoclean updated the final count/size input; reloading Fiji threshold...';
  ['thresholdOriginalPreview', 'threshold8bitPreview', 'thresholdWatershedPreview'].forEach((id) => {{
    const box = thresholdEl(id);
    if (box) box.innerHTML = '<div class="empty">Reloading from auto-cleaned input...</div>';
  }});
  await loadFijiThreshold(forceAuto);
}}
async function updateCountSizeAfterIndividualAutoclean(rel, mode='auto_cleaned') {{
  const sharedInput = document.querySelector('[data-shared-analysis-input-mode]');
  if (sharedInput) sharedInput.value = mode || 'auto_cleaned';
  const select = thresholdEl('thresholdCropRel');
  if (!select || !rel) return;
  let hasCrop = Array.from(select.options).some((option) => option.value === rel);
  if (!hasCrop) {{
    await refreshCountSizeThresholdCrops();
    hasCrop = Array.from(select.options).some((option) => option.value === rel);
  }}
  if (hasCrop) select.value = rel;
  updateThresholdCropDisplay();
  await refreshCountSizeThresholdPanel(false);
}}
function renderThresholdPreview(includeWatershed=true, fastWatershed=false) {{
  const rel = selectedThresholdCrop();
  const channel = selectedThresholdChannel();
  if (!rel) return;
  const values = setThresholdControls(thresholdEl('thresholdMinValue').value, thresholdEl('thresholdMaxValue').value);
  const requestId = ++thresholdPreviewRequest;
  const status = thresholdEl('thresholdStatus');
  if (status) status.textContent = `${{status.textContent.replace(' | updating Fiji preview...', '')}} | updating Fiji preview...`;
  const mode = selectedSharedAnalysisInputMode();
  const base = `/count_size_threshold_preview?rel=${{encodeURIComponent(rel)}}&channel=${{encodeURIComponent(channel)}}&analysis_input_mode=${{encodeURIComponent(mode)}}&threshold_min=${{values.lower}}&threshold_max=${{values.upper}}&v=${{Date.now()}}`;
  const panels = [
    ['thresholdOriginalPreview', 'original', 'Loading original channel...'],
    ['threshold8bitPreview', '8bit', 'Loading 8-bit threshold...'],
  ];
  if (includeWatershed) panels.push([
    'thresholdWatershedPreview',
    'watershed',
    fastWatershed ? 'Updating live watershed preview...' : 'Running exact Fiji watershed preview...'
  ]);
  let remaining = panels.length;
  panels.forEach(([id, view, loading]) => {{
    const box = thresholdEl(id);
    if (!box) return;
    if (!box.querySelector('img')) box.innerHTML = `<div class="empty">${{loading}}</div>`;
    const img = new Image();
    img.onload = () => {{
      if (requestId !== thresholdPreviewRequest) return;
      box.innerHTML = '';
      box.appendChild(img);
      remaining -= 1;
      if (remaining === 0 && status) status.textContent = status.textContent.replace(' | updating Fiji preview...', '');
    }};
    img.onerror = () => {{
      if (requestId !== thresholdPreviewRequest) return;
      box.innerHTML = `<div class="empty">${{view}} preview failed</div>`;
      remaining -= 1;
      if (remaining === 0 && status) status.textContent = status.textContent.replace(' | updating Fiji preview...', ' | Fiji preview failed');
    }};
    const fastQuery = view === 'watershed' && fastWatershed ? '&fast_watershed=1' : '';
    img.src = `${{base}}&view=${{encodeURIComponent(view)}}${{fastQuery}}`;
  }});
}}
function scheduleThresholdPreview() {{
  clearTimeout(thresholdPreviewTimer);
  thresholdPreviewTimer = setTimeout(() => renderThresholdPreview(true, true), 90);
}}
async function saveCountSizeThreshold(quiet=false, relOverride='', channelOverride='') {{
  const rel = relOverride || selectedThresholdCrop();
  const channel = channelOverride || selectedThresholdChannel();
  if (!rel) return false;
  const values = setThresholdControls(thresholdEl('thresholdMinValue').value, thresholdEl('thresholdMaxValue').value);
  const body = new URLSearchParams();
  body.set('rel', rel);
  body.set('channel', channel);
  body.set('threshold_min', values.lower);
  body.set('threshold_max', values.upper);
  if (!quiet) thresholdEl('thresholdStatus').textContent = `saving ${{values.lower}}-${{values.upper}}...`;
  try {{
    const response = await fetch('/save_count_size_threshold', {{ method: 'POST', body }});
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || response.statusText);
    thresholdDirty = false;
    thresholdLoadedRel = rel;
    thresholdLoadedChannel = channel;
    thresholdSessionValues.set(thresholdSessionKey(rel, channel), {{ min: data.threshold_min, max: data.threshold_max }});
    thresholdEl('thresholdStatus').textContent = `saved ${{data.threshold_min}}-${{data.threshold_max}} for ${{channel}} count/size in this experiment`;
    return true;
  }} catch (error) {{
    thresholdEl('thresholdStatus').textContent = `threshold save failed: ${{error.message}}`;
    return false;
  }}
}}
['thresholdMinSlider','thresholdMinValue','thresholdMaxSlider','thresholdMaxValue'].forEach(id => {{
  const input = thresholdEl(id);
  if (!input) return;
  input.addEventListener('input', () => {{
    thresholdDirty = true;
    if (id.includes('MinSlider')) thresholdEl('thresholdMinValue').value = input.value;
    if (id.includes('MinValue')) thresholdEl('thresholdMinSlider').value = input.value;
    if (id.includes('MaxSlider')) thresholdEl('thresholdMaxValue').value = input.value;
    if (id.includes('MaxValue')) thresholdEl('thresholdMaxSlider').value = input.value;
    scheduleThresholdPreview();
  }});
  input.addEventListener('change', () => {{
    if (id.includes('MinSlider')) thresholdEl('thresholdMinValue').value = input.value;
    if (id.includes('MinValue')) thresholdEl('thresholdMinSlider').value = input.value;
    if (id.includes('MaxSlider')) thresholdEl('thresholdMaxValue').value = input.value;
    if (id.includes('MaxValue')) thresholdEl('thresholdMaxSlider').value = input.value;
    clearTimeout(thresholdPreviewTimer);
    renderThresholdPreview(true);
  }});
}});
if (thresholdEl('thresholdCropRel')) thresholdEl('thresholdCropRel').addEventListener('change', async () => {{
  if (thresholdDirty && !(await saveCountSizeThreshold(true, thresholdLoadedRel, thresholdLoadedChannel))) return;
  updateThresholdCropDisplay();
  await loadFijiThreshold(false);
}});
if (thresholdEl('refreshThresholdCrops')) thresholdEl('refreshThresholdCrops').addEventListener('click', async () => {{
  const button = thresholdEl('refreshThresholdCrops');
  const oldText = button.textContent;
  button.disabled = true;
  button.textContent = 'Refreshing crop list...';
  try {{
    if (thresholdDirty && !(await saveCountSizeThreshold(true, thresholdLoadedRel, thresholdLoadedChannel))) throw new Error('Unsaved threshold could not be saved');
    const refreshed = await refreshCountSizeThresholdCrops(true);
    if (!refreshed) throw new Error('Count/size crop-list refresh failed');
    button.textContent = 'Count/size list updated';
    setTimeout(() => {{ button.textContent = oldText; }}, 1800);
  }} catch (error) {{
    button.textContent = 'Refresh failed';
    setTimeout(() => {{ button.textContent = oldText; }}, 2500);
  }} finally {{
    button.disabled = false;
  }}
}});
if (thresholdEl('thresholdPrevCrop')) thresholdEl('thresholdPrevCrop').addEventListener('click', () => moveThresholdCrop(-1));
if (thresholdEl('thresholdNextCrop')) thresholdEl('thresholdNextCrop').addEventListener('click', () => moveThresholdCrop(1));
if (thresholdEl('thresholdChannel')) thresholdEl('thresholdChannel').addEventListener('change', async () => {{
  if (thresholdDirty && !(await saveCountSizeThreshold(true, thresholdLoadedRel, thresholdLoadedChannel))) return;
  await loadFijiThreshold(false);
}});
if (thresholdEl('fijiAutoThreshold')) thresholdEl('fijiAutoThreshold').addEventListener('click', async () => {{
  await loadFijiThreshold(true);
  thresholdDirty = true;
  thresholdEl('thresholdStatus').textContent += ' | unsaved—click Save threshold';
}});
if (thresholdEl('resetThreshold')) thresholdEl('resetThreshold').addEventListener('click', () => {{
  setThresholdControls(0, 255);
  thresholdDirty = true;
  thresholdEl('thresholdStatus').textContent = 'reset 0-255 | unsaved—click Save threshold';
  renderThresholdPreview();
}});
if (thresholdEl('saveThreshold')) thresholdEl('saveThreshold').addEventListener('click', () => saveCountSizeThreshold(false));
document.querySelectorAll('form[action="/analyze_count_size"], form[action="/analyze_pcc_manders_count_size"]').forEach((form) => {{
  form.addEventListener('submit', async (event) => {{
    if (!thresholdDirty || form.dataset.thresholdSavedForSubmit === '1') {{
      form.dataset.thresholdSavedForSubmit = '0';
      return;
    }}
    event.preventDefault();
    const submitter = event.submitter;
    if (!(await saveCountSizeThreshold(true, thresholdLoadedRel, thresholdLoadedChannel))) return;
    form.dataset.thresholdSavedForSubmit = '1';
    if (form.requestSubmit) form.requestSubmit(submitter); else form.submit();
  }});
}});
if (thresholdEl('thresholdCropRel')) loadFijiThreshold(false);


let countSizeProgressTimer = null;
async function pollCountSizeProgress() {{
  const panel = document.getElementById('countSizeProgressPanel');
  const bar = document.getElementById('countSizeProgressBar');
  const text = document.getElementById('countSizeProgressText');
  if (!panel || !bar || !text) return;
  try {{
    const response = await fetch('/api/analysis_progress?v=' + Date.now(), {{cache:'no-store'}});
    const progress = await response.json();
    panel.style.display = 'block';
    bar.value = Number(progress.percent || 0);
    text.textContent = `${{progress.message || progress.stage || 'Working'}} (${{progress.done || 0}}/${{progress.total || 0}})`;
    if (!progress.active && (progress.stage === 'complete' || progress.stage === 'failed')) {{
      clearInterval(countSizeProgressTimer);
      countSizeProgressTimer = null;
    }}
  }} catch (error) {{
    text.textContent = 'Analysis is running; waiting for the next progress update…';
  }}
}}
function startCountSizeProgress() {{
  const panel = document.getElementById('countSizeProgressPanel');
  if (panel) panel.style.display = 'block';
  if (countSizeProgressTimer) clearInterval(countSizeProgressTimer);
  pollCountSizeProgress();
  countSizeProgressTimer = setInterval(pollCountSizeProgress, 750);
}}
document.addEventListener('submit', (event) => {{
  const submitter = event.submitter;
  const action = submitter ? (submitter.getAttribute('formaction') || '') : (event.target.getAttribute('action') || '');
  if (['/analyze_pcc_manders', '/analyze_count_size', '/analyze_pcc_manders_count_size', '/analyze_fixed_cell_ms2', '/analyze_manual_partition'].includes(action)) startCountSizeProgress();
}});

function updateCounts(oldStatus, newStatus) {{
  const kept = document.getElementById('kept-count');
  const rejected = document.getElementById('rejected-count');
  let keptValue = parseInt(kept.textContent || '0', 10);
  let rejectedValue = parseInt(rejected.textContent || '0', 10);
  if (oldStatus === newStatus) return;
  if (oldStatus === 'keep') keptValue -= 1;
  if (oldStatus === 'reject') rejectedValue -= 1;
  if (newStatus === 'keep') keptValue += 1;
  if (newStatus === 'reject') rejectedValue += 1;
  kept.textContent = Math.max(0, keptValue);
  rejected.textContent = Math.max(0, rejectedValue);
  updateReviewFilterCounts(keptValue + rejectedValue, keptValue, rejectedValue);
}}
function updateReviewFilterCounts(total, kept, rejected) {{
  const values = {{ all: total, keep: kept, reject: rejected }};
  document.querySelectorAll('[data-review-filter]').forEach((button) => {{
    const filter = button.dataset.reviewFilter || 'all';
    const label = filter === 'all' ? 'Show all' : filter === 'keep' ? 'Show kept' : 'Show rejected';
    const count = values[filter];
    if (Number.isFinite(count)) button.textContent = `${{label}} (${{count}})`;
  }});
}}
function elementByDataRel(selector, rel) {{
  return Array.from(document.querySelectorAll(selector)).find((element) => element.dataset.rel === rel) || null;
}}
function setGridEmptyMessage(grid, message) {{
  if (!grid) return;
  grid.querySelectorAll('[data-empty-message]').forEach((node) => node.remove());
  if (!grid.children.length) {{
    const empty = document.createElement('p');
    empty.dataset.emptyMessage = '1';
    empty.textContent = message;
    grid.appendChild(empty);
  }}
}}
function applyAuxiliaryLimit(grid) {{
  if (!grid) return;
  const cards = Array.from(grid.children).filter((node) => !node.dataset.emptyMessage);
  cards.forEach((card, index) => {{
    if (index >= {AUXILIARY_PREVIEW_LIMIT}) {{
      card.classList.add('auxiliary-extra');
    }} else {{
      card.classList.remove('auxiliary-extra');
      card.style.display = '';
    }}
  }});
}}
async function ensureCleaningCardForRel(rel, grid) {{
  if (!rel || !grid || elementByDataRel('.cleaning-card', rel)) return;
  await refreshCleaningCardForRel(rel, grid, false);
}}
async function refreshCleaningCardForRel(rel, grid, replaceExisting=true) {{
  if (!rel || !grid) return;
  try {{
    const response = await fetch('/api/cleaning_card?rel=' + encodeURIComponent(rel) + '&v=' + Date.now());
    const payload = await response.json();
    if (!response.ok || !payload.ok) return;
    const existing = elementByDataRel('.cleaning-card', rel);
    if (!payload.html) {{
      if (existing) existing.remove();
      applyAuxiliaryLimit(grid);
      return;
    }}
    const wrapper = document.createElement('div');
    wrapper.innerHTML = payload.html;
    const card = wrapper.firstElementChild;
    if (!card) return;
    grid.querySelectorAll('[data-empty-message]').forEach((node) => node.remove());
    if (existing && replaceExisting) {{
      existing.replaceWith(card);
    }} else if (!existing) {{
      grid.appendChild(card);
    }}
    bindLoadedCleaningPreviewCards(grid);
    applyAuxiliaryLimit(grid);
  }} catch (error) {{
    console.warn('Could not refresh raw-vs-cleaned card for kept crop.', error);
  }}
}}
function cleaningPreviewShouldShowAll() {{
  const grid = document.getElementById('cleaningPreviewGrid');
  const holder = document.getElementById('cleaningShowMoreHolder');
  if (!grid || !holder) return false;
  return !holder.querySelector('.auxiliary-show-all') && grid.querySelectorAll('.cleaning-card').length >= {AUXILIARY_PREVIEW_LIMIT};
}}
async function refreshCleaningPreviewSection(showAll=null) {{
  const grid = document.getElementById('cleaningPreviewGrid');
  const holder = document.getElementById('cleaningShowMoreHolder');
  if (!grid) return;
  if (showAll === null) showAll = cleaningPreviewShouldShowAll();
  try {{
    const response = await fetch('/api/cleaning_preview_section?show_all=' + (showAll ? '1' : '0') + '&v=' + Date.now());
    const payload = await response.json();
    if (!response.ok || !payload.ok) throw new Error(payload.error || 'Raw vs cleaned refresh failed');
    grid.innerHTML = payload.html || '<p data-empty-message="1">No kept crops are available for raw-vs-cleaned preview yet.</p>';
    if (holder) holder.innerHTML = payload.button_html || '';
    bindLoadedCleaningPreviewCards(grid);
    if (!showAll) applyAuxiliaryLimit(grid);
  }} catch (error) {{
    console.warn('Could not refresh raw-vs-cleaned preview section.', error);
  }}
}}
async function refreshRecoveredCropsSection() {{
  const section = document.getElementById('recoveredCropsSection');
  if (!section) return;
  try {{
    const response = await fetch('/api/recovered_crops_section?v=' + Date.now());
    const payload = await response.json();
    if (!response.ok || !payload.ok || !payload.html) throw new Error(payload && payload.error ? payload.error : 'Recovered refresh failed');
    section.outerHTML = payload.html;
  }} catch (error) {{
    console.warn('Could not refresh recovered crop section.', error);
  }}
}}
async function refreshCountSizeThresholdCrops(reloadSelected=false) {{
  const select = thresholdEl('thresholdCropRel');
  if (!select) return;
  const previousRel = select.value;
  try {{
    const response = await fetch('/api/count_size_threshold_crops?v=' + Date.now());
    const payload = await response.json();
    if (!response.ok || !payload.ok) throw new Error(payload.error || 'Count/size crop refresh failed');
    select.innerHTML = '';
    (payload.crops || []).forEach((crop) => {{
      const option = document.createElement('option');
      option.value = crop.rel || '';
      option.textContent = crop.label || crop.rel || 'Kept crop';
      select.appendChild(option);
    }});
    if (!select.options.length) {{
      const option = document.createElement('option');
      option.value = '';
      option.textContent = 'No kept crops yet';
      select.appendChild(option);
    }}
    const previousStillKept = Array.from(select.options).some((option) => option.value === previousRel && option.value);
    if (previousStillKept) select.value = previousRel;
    updateThresholdCropDisplay();
    if (select.value && (select.value !== previousRel || reloadSelected)) {{
      await loadFijiThreshold(false);
    }} else if (!select.value) {{
      thresholdHistogram = [];
      setThresholdControls(0, 255);
      const status = thresholdEl('thresholdStatus');
      if (status) status.textContent = 'No kept crops available for count/size thresholding.';
      ['thresholdOriginalPreview', 'threshold8bitPreview', 'thresholdWatershedPreview'].forEach((id) => {{
        const box = thresholdEl(id);
        if (box) box.innerHTML = '<div class="empty">No kept crop selected</div>';
      }});
    }}
    return true;
  }} catch (error) {{
    console.warn('Could not refresh count/size kept-crop list.', error);
    return false;
  }}
}}
async function refreshReviewDependentSectionsFromServer() {{
  await Promise.all([
    refreshRecoveredCropsSection(),
    refreshCleaningPreviewSection(null),
    refreshCountSizeThresholdCrops()
  ]);
}}
function updateReviewDependentSections() {{
  const reviewCards = Array.from(document.querySelectorAll('.review-card'));
  const keptGrid = document.getElementById('recoveredKeptGrid');
  const rejectedGrid = document.getElementById('recoveredRejectedGrid');
  const cleaningGrid = document.getElementById('cleaningPreviewGrid');
  if (keptGrid) keptGrid.querySelectorAll('[data-empty-message]').forEach((node) => node.remove());
  if (rejectedGrid) rejectedGrid.querySelectorAll('[data-empty-message]').forEach((node) => node.remove());
  if (cleaningGrid) cleaningGrid.querySelectorAll('[data-empty-message]').forEach((node) => node.remove());
  let kept = 0;
  let rejected = 0;
  reviewCards.forEach((card) => {{
    const rel = card.dataset.rel || '';
    const status = card.dataset.status === 'reject' ? 'reject' : 'keep';
    const recoveredCard = elementByDataRel('.recovered-crop-card', rel);
    const cleaningCard = elementByDataRel('.cleaning-card', rel);
    if (status === 'keep') {{
      kept += 1;
      if (keptGrid && recoveredCard) {{
        recoveredCard.dataset.recoveryStatus = 'keep';
        keptGrid.appendChild(recoveredCard);
      }}
      if (cleaningCard) {{
        cleaningCard.style.display = '';
      }} else {{
        ensureCleaningCardForRel(rel, cleaningGrid);
      }}
    }} else {{
      rejected += 1;
      if (rejectedGrid && recoveredCard) {{
        recoveredCard.dataset.recoveryStatus = 'reject';
        rejectedGrid.appendChild(recoveredCard);
      }}
      if (cleaningCard) cleaningCard.remove();
    }}
  }});
  const recoveredKeptCount = document.getElementById('recoveredKeptCount');
  const recoveredRejectedCount = document.getElementById('recoveredRejectedCount');
  const recoveredKeptSummary = document.getElementById('recoveredKeptSummary');
  const recoveredRejectedSummary = document.getElementById('recoveredRejectedSummary');
  if (recoveredKeptCount) recoveredKeptCount.textContent = kept;
  if (recoveredRejectedCount) recoveredRejectedCount.textContent = rejected;
  if (recoveredKeptSummary) recoveredKeptSummary.textContent = `Recovered kept cropped images (${{kept}})`;
  if (recoveredRejectedSummary) recoveredRejectedSummary.textContent = `Recovered rejected cropped images (${{rejected}})`;
  setGridEmptyMessage(keptGrid, 'No kept crop images are currently loaded.');
  setGridEmptyMessage(rejectedGrid, 'No rejected crop images are currently loaded.');
  setGridEmptyMessage(cleaningGrid, 'No kept crops are available for raw-vs-cleaned preview yet.');
  applyAuxiliaryLimit(keptGrid);
  applyAuxiliaryLimit(rejectedGrid);
  applyAuxiliaryLimit(cleaningGrid);
}}
function currentReviewFilter() {{
  const active = document.querySelector('[data-review-filter].active');
  return active ? active.dataset.reviewFilter : 'all';
}}
function applyReviewFilter(filter) {{
  document.querySelectorAll('[data-review-filter]').forEach((button) => {{
    button.classList.toggle('active', button.dataset.reviewFilter === filter);
  }});
  document.querySelectorAll('.review-card').forEach((card) => {{
    const show = filter === 'all' || card.dataset.status === filter;
    card.style.display = show ? '' : 'none';
  }});
}}
document.querySelectorAll('[data-review-filter]').forEach((button) => {{
  button.addEventListener('click', () => applyReviewFilter(button.dataset.reviewFilter || 'all'));
}});
function paintCardStatus(card, banner, newStatus) {{
  if (!card || !banner) return;
  const isKeep = newStatus === 'keep';
  card.dataset.status = newStatus;
  card.style.borderColor = isKeep ? '#15803d' : '#b91c1c';
  banner.className = 'status-banner ' + newStatus;
  banner.innerHTML = isKeep
    ? '<span>KEEP</span><span>Included in final analysis</span>'
    : '<span>REJECTED</span><span>Excluded from final analysis</span>';
  card.querySelectorAll('[data-choice-button]').forEach((button) => {{
    button.classList.toggle('status-clicked', button.dataset.choiceButton === newStatus);
  }});
  if (newStatus === 'keep') {{
    card.dataset.explicitKeep = 'true';
  }} else if (newStatus === 'reject') {{
    card.dataset.explicitKeep = 'false';
  }}
  applyReviewFilter(currentReviewFilter());
}}
function instantSetStatus(cardId, badgeId, newStatus) {{
  const card = document.getElementById(cardId);
  const banner = document.getElementById(badgeId);
  if (!card || !banner) return;
  const oldStatus = card.dataset.status || 'keep';
  paintCardStatus(card, banner, newStatus);
  updateCounts(oldStatus, newStatus);
}}

document.querySelectorAll('.reject-image-rest-button').forEach((button) => {{
  button.addEventListener('click', async () => {{
    const rel = button.dataset.rel || '';
    const card = elementByDataRel('.review-card', rel);
    if (!rel || !card) return;
    const oldText = button.textContent;
    button.disabled = true;
    button.textContent = 'Rejecting...';
    try {{
      const body = new URLSearchParams();
      body.set('rel', rel);
      const response = await fetch('/api/reject_other_crops_from_image', {{ method: 'POST', body }});
      const payload = await response.json().catch(() => null);
      if (!response.ok || !payload || !payload.ok) throw new Error(payload && payload.error ? payload.error : 'Batch reject failed');
      (payload.rejected_rels || []).forEach((rejectedRel) => {{
        const rejectedCard = elementByDataRel('.review-card', rejectedRel);
        if (!rejectedCard) return;
        const banner = rejectedCard.querySelector('.status-banner');
        paintCardStatus(rejectedCard, banner, 'reject');
      }});
      if (payload.current_rel) {{
        const currentCard = elementByDataRel('.review-card', payload.current_rel);
        if (currentCard) {{
          const currentBanner = currentCard.querySelector('.status-banner');
          paintCardStatus(currentCard, currentBanner, 'keep');
          currentCard.dataset.explicitKeep = 'true';
        }}
      }}
      if (Number.isFinite(payload.kept) && Number.isFinite(payload.rejected)) {{
        document.getElementById('kept-count').textContent = payload.kept;
        document.getElementById('rejected-count').textContent = payload.rejected;
        updateReviewFilterCounts(payload.total, payload.kept, payload.rejected);
      }}
      await refreshReviewDependentSectionsFromServer();
      button.textContent = 'Rejected ' + (payload.rejected_rels || []).length + ' other crop(s)';
      setTimeout(() => {{ button.textContent = oldText; }}, 1800);
    }} catch (error) {{
      button.textContent = 'Batch reject failed';
      setTimeout(() => {{ button.textContent = oldText; }}, 2200);
    }} finally {{
      button.disabled = false;
    }}
  }});
}});

document.querySelectorAll('.status-form').forEach((form) => {{
  form.addEventListener('submit', async (event) => {{
    event.preventDefault();
    const cardId = form.dataset.cardId;
    const badgeId = form.dataset.badgeId;
    const newStatus = form.dataset.status;
    instantSetStatus(cardId, badgeId, newStatus);
    try {{
      const response = await fetch('/api/set_status', {{ method: 'POST', body: new URLSearchParams(new FormData(form)) }});
      if (!response.ok) throw new Error('Save failed');
      const payload = await response.json();
      if (payload && payload.explicit_keep !== undefined) {{
        const card = document.getElementById(cardId);
        if (card) card.dataset.explicitKeep = payload.explicit_keep ? 'true' : 'false';
      }}
      if (payload && Number.isFinite(payload.kept) && Number.isFinite(payload.rejected)) {{
        document.getElementById('kept-count').textContent = payload.kept;
        document.getElementById('rejected-count').textContent = payload.rejected;
        updateReviewFilterCounts(payload.total, payload.kept, payload.rejected);
      }}
      await refreshReviewDependentSectionsFromServer();
    }} catch (error) {{
      console.warn('Status save check failed; reload will show the saved state if the server received it.', error);
    }}
}});
}});

const cleaningTimers = new WeakMap();
const cleaningRequests = new WeakMap();
const cleaningSaveTimers = new WeakMap();
const uiPersistenceScope = {json.dumps(str(state.get("raw_image_folder", "") or ""))};
const globalUiDraftKey = 'condensate-ui-global-v1:' + encodeURIComponent(uiPersistenceScope);

function combinedGlobalSettingsBody() {{
  const body = new URLSearchParams();
  [document.getElementById('analysis-labels'), document.getElementById('globalCleaningPreviewForm')].forEach((form) => {{
    if (!form) return;
    new FormData(form).forEach((value, key) => body.set(key, value));
  }});
  return body;
}}

function restoreEncodedFormDraft(encoded, forms) {{
  if (!encoded) return false;
  const params = new URLSearchParams(encoded);
  forms.filter(Boolean).forEach((form) => {{
    form.querySelectorAll('input[name], select[name], textarea[name]').forEach((control) => {{
      if (control.type === 'file' || control.type === 'submit' || control.type === 'button') return;
      if (control.type === 'checkbox' || control.type === 'radio') {{
        control.checked = params.getAll(control.name).includes(control.value);
      }} else if (params.has(control.name)) {{
        control.value = params.get(control.name);
      }}
    }});
  }});
  return true;
}}

let globalSettingsSaveTimer = null;
async function saveGlobalUiSettings() {{
  const body = combinedGlobalSettingsBody();
  try {{
    const response = await fetch('/api/save_ui_settings', {{method: 'POST', body, keepalive: true}});
    const payload = await response.json().catch(() => null);
    if (!response.ok || !payload || !payload.ok) throw new Error('Settings save failed');
    localStorage.removeItem(globalUiDraftKey);
  }} catch (error) {{
    console.warn('Settings remain in browser recovery storage and will retry after reload.', error);
  }}
}}

function stageGlobalUiSettingsSave() {{
  const body = combinedGlobalSettingsBody();
  try {{ localStorage.setItem(globalUiDraftKey, body.toString()); }} catch (_error) {{}}
  if (globalSettingsSaveTimer) clearTimeout(globalSettingsSaveTimer);
  globalSettingsSaveTimer = setTimeout(saveGlobalUiSettings, 120);
}}

function cleaningDraftKey(form) {{
  const rel = form.querySelector('input[name="rel"]')?.value || '';
  return 'condensate-ui-crop-v1:' + encodeURIComponent(uiPersistenceScope + '\\n' + rel);
}}

async function saveOneCleaningSettings(form) {{
  const key = cleaningDraftKey(form);
  try {{
    const response = await fetch('/api/save_one_channel_cleaning_settings', {{
      method: 'POST', body: new URLSearchParams(new FormData(form)), keepalive: true
    }});
    const payload = await response.json().catch(() => null);
    if (!response.ok || !payload || !payload.ok) throw new Error('Per-crop settings save failed');
    localStorage.removeItem(key);
  }} catch (error) {{
    console.warn('Per-crop settings remain in browser recovery storage and will retry after reload.', error);
  }}
}}

function stageOneCleaningSettingsSave(form) {{
  const encoded = new URLSearchParams(new FormData(form)).toString();
  try {{ localStorage.setItem(cleaningDraftKey(form), encoded); }} catch (_error) {{}}
  if (cleaningSaveTimers.has(form)) clearTimeout(cleaningSaveTimers.get(form));
  cleaningSaveTimers.set(form, setTimeout(() => saveOneCleaningSettings(form), 120));
}}

function updateCleaningValueLabel(form, input) {{
  const valueLabel = form.querySelector('[data-value-for="' + input.name + '"]');
  if (valueLabel) valueLabel.textContent = input.value;
}}

let globalCleaningTimer = null;
let globalCleaningRequestId = 0;

function updateGlobalCleaningValueLabel(input) {{
  const valueLabel = document.querySelector('[data-global-value-for="' + input.name + '"]');
  if (valueLabel) valueLabel.textContent = input.value;
}}

async function refreshGlobalCleaningPreviews() {{
  const settingsForm = document.getElementById('analysis-labels');
  const globalForm = document.getElementById('globalCleaningPreviewForm');
  const grid = document.getElementById('cleaningPreviewGrid');
  const holder = document.getElementById('cleaningShowMoreHolder');
  const status = document.getElementById('resetAllCleaningStatus');
  const applyButton = document.getElementById('applyGlobalCleaningButton');
  if (!globalForm || !grid) return;
  const requestId = ++globalCleaningRequestId;
  const body = new URLSearchParams();
  [settingsForm, globalForm].forEach((form) => {{
    if (!form) return;
    new FormData(form).forEach((value, key) => body.set(key, value));
  }});
  if (status) status.textContent = 'Autocleaning kept crops from global sliders...';
  if (applyButton) {{
    applyButton.disabled = true;
    applyButton.textContent = 'Autocleaning kept crops...';
  }}
  try {{
    const response = await fetch('/api/preview_global_channel_cleaning', {{
      method: 'POST',
      body: body
    }});
    const payload = await response.json().catch(() => null);
    if (requestId !== globalCleaningRequestId) return;
    if (!response.ok || !payload || !payload.ok) {{
      throw new Error(payload && payload.error ? payload.error : 'Global preview update failed');
    }}
    localStorage.removeItem(globalUiDraftKey);
    const sharedInput = document.querySelector('[data-shared-analysis-input-mode]');
    if (sharedInput) sharedInput.value = payload.analysis_input_mode || 'auto_cleaned';
    grid.innerHTML = payload.html || '<p data-empty-message="1">No kept crops are available for raw-vs-cleaned preview yet.</p>';
    if (holder) holder.innerHTML = payload.button_html || '';
    bindLoadedCleaningPreviewCards(grid);
    document.querySelectorAll('.review-card[data-status="keep"]').forEach((card) => setCardAutocleanNeeded(card, false));
    await refreshRecoveredCropsSection();
    await refreshCountSizeThresholdPanel(false);
    if (status) status.textContent = 'Updated ' + payload.count + ' kept preview(s) from global sliders.';
  }} catch (error) {{
    if (status) status.textContent = 'Global preview update failed: ' + String(error.message || error);
  }} finally {{
    if (requestId === globalCleaningRequestId && applyButton) {{
      applyButton.disabled = false;
      applyButton.textContent = 'Autoclean kept crops';
    }}
  }}
}}

const globalDraft = (() => {{ try {{ return localStorage.getItem(globalUiDraftKey); }} catch (_error) {{ return null; }} }})();
if (restoreEncodedFormDraft(globalDraft, [document.getElementById('analysis-labels'), document.getElementById('globalCleaningPreviewForm')])) {{
  document.querySelectorAll('.global-cleaning-slider').forEach(updateGlobalCleaningValueLabel);
  stageGlobalUiSettingsSave();
}}
[document.getElementById('analysis-labels'), document.getElementById('globalCleaningPreviewForm')].filter(Boolean).forEach((form) => {{
  form.querySelectorAll('input[name], select[name], textarea[name]').forEach((control) => {{
    if (control.type === 'file' || control.type === 'submit' || control.type === 'button') return;
    control.addEventListener('input', stageGlobalUiSettingsSave);
    control.addEventListener('change', stageGlobalUiSettingsSave);
  }});
}});

document.querySelectorAll('.global-cleaning-slider').forEach((input) => {{
  input.addEventListener('input', () => {{
    updateGlobalCleaningValueLabel(input);
    if (globalCleaningTimer) clearTimeout(globalCleaningTimer);
    const status = document.getElementById('resetAllCleaningStatus');
    if (status) status.textContent = 'Release the slider or click Autoclean kept crops to regenerate every kept raw-vs-cleaned preview.';
  }});
  input.addEventListener('change', () => {{
    updateGlobalCleaningValueLabel(input);
    if (globalCleaningTimer) clearTimeout(globalCleaningTimer);
    globalCleaningTimer = setTimeout(refreshGlobalCleaningPreviews, 150);
  }});
}});

const applyGlobalCleaningButton = document.getElementById('applyGlobalCleaningButton');
if (applyGlobalCleaningButton) {{
  applyGlobalCleaningButton.addEventListener('click', () => {{
    if (globalCleaningTimer) clearTimeout(globalCleaningTimer);
    refreshGlobalCleaningPreviews();
  }});
}}

async function refreshCleaningPreview(form) {{
  const card = form.closest('.cleaning-card');
  if (!card) return;
  const panel = card.querySelector('[data-role="cleaning-preview-panel"]');
  const status = card.querySelector('[data-role="cleaning-status"]');
  const roleStatus = card.querySelector('[data-role="role-status"]');
  const button = form.querySelector('button[type="submit"]');
  const requestId = (cleaningRequests.get(form) || 0) + 1;
  cleaningRequests.set(form, requestId);
  if (status) status.innerHTML = '<b>Autocleaning only this crop...</b>';
  if (button) {{
    button.disabled = true;
    button.textContent = 'Autocleaning this crop...';
  }}
  try {{
    const response = await fetch('/api/preview_one_channel_cleaning', {{
      method: 'POST',
      body: new URLSearchParams(new FormData(form))
    }});
    const payload = await response.json().catch(() => null);
    if (cleaningRequests.get(form) !== requestId) return;
    if (!response.ok || !payload || !payload.ok) {{
      throw new Error(payload && payload.error ? payload.error : 'Per-crop autoclean failed');
    }}
    localStorage.removeItem(cleaningDraftKey(form));
    applyCleaningPayloadToCard(card, payload);
  }} catch (error) {{
    if (status) status.innerHTML = '<b>Per-crop autoclean failed:</b> ' + String(error.message || error);
  }} finally {{
    if (cleaningRequests.get(form) === requestId && button) {{
      button.disabled = false;
      button.textContent = 'Autoclean this crop only';
    }}
  }}
}}

document.querySelectorAll('.live-cleaning-form').forEach((form) => {{
  let recoveredDraft = null;
  try {{ recoveredDraft = localStorage.getItem(cleaningDraftKey(form)); }} catch (_error) {{}}
  if (restoreEncodedFormDraft(recoveredDraft, [form])) {{
    form.querySelectorAll('input[type="range"]').forEach((input) => updateCleaningValueLabel(form, input));
    stageOneCleaningSettingsSave(form);
  }}
  form.querySelectorAll('input[type="range"]').forEach((input) => {{
    input.addEventListener('input', () => {{
      updateCleaningValueLabel(form, input);
      stageOneCleaningSettingsSave(form);
      if (cleaningTimers.has(form)) clearTimeout(cleaningTimers.get(form));
      const status = form.closest('.cleaning-card')?.querySelector('[data-role="cleaning-status"]');
      if (status) status.innerHTML = '<b>Release slider or click Autoclean this crop only.</b>';
    }});
    input.addEventListener('change', () => {{
      updateCleaningValueLabel(form, input);
      if (cleaningTimers.has(form)) clearTimeout(cleaningTimers.get(form));
      refreshCleaningPreview(form);
    }});
  }});
  form.querySelectorAll('select[name$="_role"]').forEach((select) => {{
    select.addEventListener('change', () => {{
      stageOneCleaningSettingsSave(form);
      if (cleaningTimers.has(form)) clearTimeout(cleaningTimers.get(form));
      const channel = select.name.replace(/_role$/, '');
      const rule = form.querySelector('[data-role="' + channel + '-cleaning-rule"]');
      if (rule) rule.innerHTML = '<b>' + (select.value === 'condensate' ? 'Will be puncta-cleaned' : 'Will be preserved raw') + '</b>';
      refreshCleaningPreview(form);
    }});
  }});
  form.addEventListener('submit', (event) => {{
    event.preventDefault();
    if (cleaningTimers.has(form)) clearTimeout(cleaningTimers.get(form));
    refreshCleaningPreview(form);
  }});
}});

window.addEventListener('pagehide', () => {{
  const globalBody = combinedGlobalSettingsBody().toString();
  try {{ localStorage.setItem(globalUiDraftKey, globalBody); }} catch (_error) {{}}
  navigator.sendBeacon('/api/save_ui_settings', new Blob([globalBody], {{type: 'application/x-www-form-urlencoded;charset=UTF-8'}}));
  document.querySelectorAll('.live-cleaning-form').forEach((form) => {{
    const key = cleaningDraftKey(form);
    let draft = null;
    try {{ draft = localStorage.getItem(key); }} catch (_error) {{}}
    if (draft) navigator.sendBeacon('/api/save_one_channel_cleaning_settings', new Blob([draft], {{type: 'application/x-www-form-urlencoded;charset=UTF-8'}}));
  }});
}});

function applyCleaningPayloadToCard(card, payload) {{
  const panel = card.querySelector('[data-role="cleaning-preview-panel"]');
  const status = card.querySelector('[data-role="cleaning-status"]');
  const roleStatus = card.querySelector('[data-role="role-status"]');
  const form = card.querySelector('.live-cleaning-form');
  if (panel && payload.preview_url) panel.innerHTML = '<img src="' + payload.preview_url + '" loading="lazy" />';
  if (status) status.innerHTML = '<b>' + (payload.source_label || 'GLOBAL SETTINGS') + '</b>';
  if (roleStatus) roleStatus.innerHTML = '<b>' + (payload.role_source_label || 'GLOBAL ROLES') + '</b>';
  const reviewCard = cardForRel(payload.rel || card.dataset.rel || '');
  setCardAutocleanNeeded(reviewCard, false);
  refreshRecoveredCropsSection();
  const updatedRel = payload.rel || card.dataset.rel || '';
  updateCountSizeAfterIndividualAutoclean(updatedRel, payload.analysis_input_mode || 'auto_cleaned').catch((error) => {{
    console.warn('Could not refresh count/size after individual AutoClean.', error);
  }});
  if (!form || !payload.params) return;
  ['red', 'green', 'blue'].forEach((channel) => {{
    ['amount', 'background_removal', 'radius_px', 'sensitivity', 'min_area_px'].forEach((name) => {{
      const fieldName = channel + '_' + name;
      const input = form.querySelector('input[name="' + fieldName + '"]');
      if (input && payload.params[fieldName] !== undefined) {{
        input.value = payload.params[fieldName];
        updateCleaningValueLabel(form, input);
      }}
    }});
  }});
  const channels = new Set(payload.channels || []);
  ['red', 'green', 'blue'].forEach((channel) => {{
    const select = form.querySelector('select[name="' + channel + '_role"]');
    if (select && payload.roles && payload.roles[channel] !== undefined) select.value = payload.roles[channel];
    const rule = form.querySelector('[data-role="' + channel + '-cleaning-rule"]');
    if (rule) rule.innerHTML = '<b>' + (channels.has(channel) ? 'Will be puncta-cleaned' : 'Will be preserved raw') + '</b>';
  }});
}}

document.querySelectorAll('.live-reset-cleaning-form').forEach((form) => {{
  form.addEventListener('submit', async (event) => {{
    event.preventDefault();
    const card = form.closest('.cleaning-card');
    const status = card ? card.querySelector('[data-role="cleaning-status"]') : null;
    const button = form.querySelector('button[type="submit"]');
    if (status) status.innerHTML = '<b>Resetting...</b>';
    if (button) {{
      button.disabled = true;
      button.textContent = 'Resetting...';
    }}
    try {{
      const response = await fetch('/api/reset_one_channel_cleaning', {{
        method: 'POST',
        body: new URLSearchParams(new FormData(form))
      }});
      const payload = await response.json().catch(() => null);
      if (!response.ok || !payload || !payload.ok) {{
        throw new Error(payload && payload.error ? payload.error : 'Reset failed');
      }}
      if (card) applyCleaningPayloadToCard(card, payload);
    }} catch (error) {{
      if (status) status.innerHTML = '<b>Reset failed:</b> ' + String(error.message || error);
    }} finally {{
      if (button) {{
        button.disabled = false;
        button.textContent = 'Reset this crop to global settings';
      }}
    }}
  }});
}});

function bindLoadedCleaningPreviewCards(root) {{
  root.querySelectorAll('.live-cleaning-form').forEach((form) => {{
    if (form.dataset.loadedBatchBound === '1') return;
    form.dataset.loadedBatchBound = '1';
    form.querySelectorAll('input[type="range"]').forEach((input) => {{
      input.addEventListener('input', () => {{
        updateCleaningValueLabel(form, input);
        if (cleaningTimers.has(form)) clearTimeout(cleaningTimers.get(form));
        const status = form.closest('.cleaning-card')?.querySelector('[data-role="cleaning-status"]');
        if (status) status.innerHTML = '<b>Release slider or click Autoclean this crop only.</b>';
      }});
      input.addEventListener('change', () => {{
        updateCleaningValueLabel(form, input);
        if (cleaningTimers.has(form)) clearTimeout(cleaningTimers.get(form));
        refreshCleaningPreview(form);
      }});
    }});
    form.querySelectorAll('select[name$="_role"]').forEach((select) => {{
      select.addEventListener('change', () => {{
        if (cleaningTimers.has(form)) clearTimeout(cleaningTimers.get(form));
        const channel = select.name.replace(/_role$/, '');
        const rule = form.querySelector('[data-role="' + channel + '-cleaning-rule"]');
        if (rule) rule.innerHTML = '<b>' + (select.value === 'condensate' ? 'Will be puncta-cleaned' : 'Will be preserved raw') + '</b>';
        refreshCleaningPreview(form);
      }});
    }});
    form.addEventListener('submit', (event) => {{
      event.preventDefault();
      if (cleaningTimers.has(form)) clearTimeout(cleaningTimers.get(form));
      refreshCleaningPreview(form);
    }});
  }});
  root.querySelectorAll('.live-reset-cleaning-form').forEach((form) => {{
    if (form.dataset.loadedBatchBound === '1') return;
    form.dataset.loadedBatchBound = '1';
    form.addEventListener('submit', async (event) => {{
      event.preventDefault();
      const card = form.closest('.cleaning-card');
      const status = card ? card.querySelector('[data-role="cleaning-status"]') : null;
      const button = form.querySelector('button[type="submit"]');
      if (status) status.innerHTML = '<b>Resetting...</b>';
      if (button) {{
        button.disabled = true;
        button.textContent = 'Resetting...';
      }}
      try {{
        const response = await fetch('/api/reset_one_channel_cleaning', {{
          method: 'POST',
          body: new URLSearchParams(new FormData(form))
        }});
        const payload = await response.json().catch(() => null);
        if (!response.ok || !payload || !payload.ok) {{
          throw new Error(payload && payload.error ? payload.error : 'Reset failed');
        }}
        if (card) applyCleaningPayloadToCard(card, payload);
      }} catch (error) {{
        if (status) status.innerHTML = '<b>Reset failed:</b> ' + String(error.message || error);
      }} finally {{
        if (button) {{
          button.disabled = false;
          button.textContent = 'Reset this crop to global settings';
        }}
      }}
    }});
  }});
}}

const resetAllCleaningButton = document.getElementById('resetAllCleaningButton');
if (resetAllCleaningButton) resetAllCleaningButton.addEventListener('click', async () => {{
  const status = document.getElementById('resetAllCleaningStatus');
  const forms = Array.from(document.querySelectorAll('.live-reset-cleaning-form'));
  if (!forms.length) {{
    if (status) status.textContent = 'No preview cards to reset.';
    return;
  }}
  resetAllCleaningButton.disabled = true;
  resetAllCleaningButton.textContent = 'Resetting all...';
  let completed = 0;
  let failed = 0;
  for (const form of forms) {{
    const card = form.closest('.cleaning-card');
    const cardStatus = card ? card.querySelector('[data-role="cleaning-status"]') : null;
    if (cardStatus) cardStatus.innerHTML = '<b>Resetting...</b>';
    try {{
      const response = await fetch('/api/reset_one_channel_cleaning', {{
        method: 'POST',
        body: new URLSearchParams(new FormData(form))
      }});
      const payload = await response.json().catch(() => null);
      if (!response.ok || !payload || !payload.ok) {{
        throw new Error(payload && payload.error ? payload.error : 'Reset failed');
      }}
      if (card) applyCleaningPayloadToCard(card, payload);
      completed += 1;
    }} catch (error) {{
      failed += 1;
      if (cardStatus) cardStatus.innerHTML = '<b>Reset failed:</b> ' + String(error.message || error);
    }}
    if (status) status.textContent = `Reset ${{completed}} of ${{forms.length}} preview cards${{failed ? `, ${{failed}} failed` : ''}}.`;
  }}
  resetAllCleaningButton.disabled = false;
  resetAllCleaningButton.textContent = 'Reset all previews to global settings';
}});

const rangeFilterForm = document.getElementById('rangeFilterForm');
if (rangeFilterForm) rangeFilterForm.addEventListener('submit', async (event) => {{
  event.preventDefault();
  const status = document.getElementById('rangeFilterStatus');
  const submitButton = rangeFilterForm.querySelector('button[type="submit"]');
  if (status) status.textContent = 'Applying range filter...';
  if (submitButton) submitButton.disabled = true;
  try {{
    const response = await fetch('/api/select_crop_numbers', {{
      method: 'POST',
      body: new URLSearchParams(new FormData(rangeFilterForm))
    }});
    const payload = await response.json().catch(() => null);
    if (!response.ok || !payload || !payload.ok) {{
      throw new Error(payload && payload.error ? payload.error : 'Range filter failed.');
    }}
    if (Array.isArray(payload.crop_statuses)) {{
      payload.crop_statuses.forEach((entry) => {{
        const card = elementByDataRel('.review-card', entry.rel || '');
        if (!card) return;
        const banner = card.querySelector('.status-banner');
        paintCardStatus(card, banner, entry.status === 'reject' ? 'reject' : 'keep');
      }});
    }} else {{
      const selected = new Set((payload.selected_numbers || []).map((number) => String(number)));
      document.querySelectorAll('.review-card').forEach((card) => {{
        if (!selected.has(String(card.dataset.number || ''))) {{
          const banner = card.querySelector('.status-banner');
          paintCardStatus(card, banner, 'reject');
        }}
      }});
    }}
	    if (Number.isFinite(payload.kept)) document.getElementById('kept-count').textContent = payload.kept;
	    if (Number.isFinite(payload.rejected)) document.getElementById('rejected-count').textContent = payload.rejected;
	    if (Number.isFinite(payload.kept) && Number.isFinite(payload.rejected)) updateReviewFilterCounts(payload.total, payload.kept, payload.rejected);
	    applyReviewFilter(currentReviewFilter());
        if (payload.recovered_html) {{
          const recoveredSection = document.getElementById('recoveredCropsSection');
          if (recoveredSection) recoveredSection.outerHTML = payload.recovered_html;
        }}
        if (payload.cleaning_html !== undefined) {{
          const cleaningGrid = document.getElementById('cleaningPreviewGrid');
          const cleaningHolder = document.getElementById('cleaningShowMoreHolder');
          if (cleaningGrid) {{
            cleaningGrid.innerHTML = payload.cleaning_html || '<p data-empty-message="1">No kept crops are available for raw-vs-cleaned preview yet.</p>';
            bindLoadedCleaningPreviewCards(cleaningGrid);
          }}
          if (cleaningHolder) cleaningHolder.innerHTML = payload.cleaning_button_html || '';
        }} else {{
	      await refreshReviewDependentSectionsFromServer();
        }}
	    if (status) status.textContent = 'Range filter applied. Inside the range, your existing KEEP/REJECT choices were preserved. Outside the range was rejected.';
  }} catch (error) {{
    if (status) status.textContent = String(error.message || error);
  }} finally {{
    if (submitButton) submitButton.disabled = false;
  }}
}});

document.querySelectorAll('.manual-recrop-button').forEach((button) => {{
  button.addEventListener('click', () => openManualCrop(button.dataset.rel || ''));
}});

document.querySelectorAll('.fixed-cell-annotation-button').forEach((button) => {{
  button.addEventListener('click', () => openFixedCellAnnotation(button.dataset.rel || '', button));
}});

document.querySelectorAll('.partition-annotation-button').forEach((button) => {{
  button.addEventListener('click', () => openPartitionAnnotation(button.dataset.rel || '', button));
}});

document.querySelectorAll('.undo-manual-button').forEach((button) => {{
  button.addEventListener('click', () => undoManualRecrop(button.dataset.rel || '', button));
}});

document.querySelectorAll('.autoclean-one-crop-button').forEach((button) => {{
  button.addEventListener('click', async () => {{
    const rel = button.dataset.rel || '';
    const reviewCard = cardForRel(rel);
    if (!rel || !reviewCard) return;
    if ((reviewCard.dataset.status || 'keep') !== 'keep') {{
      button.textContent = 'Keep this ROI before autocleaning';
      setTimeout(() => {{ button.textContent = 'Autoclean only this crop'; }}, 2200);
      return;
    }}
    const oldText = button.textContent;
    button.disabled = true;
    button.textContent = 'Autocleaning only this crop...';
    try {{
      const body = new URLSearchParams();
      body.set('rel', rel);
      const response = await fetch('/api/autoclean_one_crop', {{ method: 'POST', body }});
      const payload = await response.json().catch(() => null);
      if (!response.ok || !payload || !payload.ok) {{
        throw new Error(payload && payload.error ? payload.error : 'Per-crop autoclean failed');
      }}
      setCardAutocleanNeeded(reviewCard, false);
      const cleaningGrid = document.getElementById('cleaningPreviewGrid');
      if (cleaningGrid) await refreshCleaningCardForRel(rel, cleaningGrid, true);
      await updateCountSizeAfterIndividualAutoclean(rel, payload.analysis_input_mode || 'auto_cleaned');
      button.textContent = 'This crop is autocleaned';
      setTimeout(() => {{ button.textContent = oldText; }}, 1800);
    }} catch (error) {{
      button.textContent = 'Autoclean failed: ' + String(error.message || error);
      setTimeout(() => {{ button.textContent = oldText; }}, 3000);
    }} finally {{
      button.disabled = false;
    }}
  }});
}});

let manualRel = '';
let manualImage = null;
let manualStart = null;
let manualRect = null;
let manualSourceMode = 'crop';

function closeManualCrop() {{
  document.getElementById('manualCropModal').classList.remove('open');
}}

function drawManualCanvas() {{
  const canvas = document.getElementById('manualCropCanvas');
  const ctx = canvas.getContext('2d');
  if (!manualImage) return;
  canvas.width = manualImage.naturalWidth;
  canvas.height = manualImage.naturalHeight;
  ctx.drawImage(manualImage, 0, 0);
  if (manualRect) {{
    ctx.strokeStyle = '#ffdd00';
    ctx.lineWidth = Math.max(2, Math.round(canvas.width / 500));
    ctx.strokeRect(manualRect.x, manualRect.y, manualRect.w, manualRect.h);
    ctx.fillStyle = 'rgba(255,221,0,0.18)';
    ctx.fillRect(manualRect.x, manualRect.y, manualRect.w, manualRect.h);
  }}
}}

function canvasPoint(event) {{
  const canvas = document.getElementById('manualCropCanvas');
  const box = canvas.getBoundingClientRect();
  return {{
    x: Math.round((event.clientX - box.left) * canvas.width / box.width),
    y: Math.round((event.clientY - box.top) * canvas.height / box.height)
  }};
}}

function syncManualFields() {{
  if (!manualRect) return;
  document.getElementById('manualX').value = Math.round(manualRect.x);
  document.getElementById('manualY').value = Math.round(manualRect.y);
  document.getElementById('manualW').value = Math.round(manualRect.w);
  document.getElementById('manualH').value = Math.round(manualRect.h);
}}

function clearManualCropFields() {{
  ['manualX', 'manualY', 'manualW', 'manualH'].forEach((id) => {{
    const input = document.getElementById(id);
    if (input) input.value = '';
  }});
}}

function clearManualCanvas() {{
  const canvas = document.getElementById('manualCropCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  canvas.width = 1;
  canvas.height = 1;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
}}

function loadManualCropImage() {{
  manualSourceMode = document.getElementById('manualSource').value || 'crop';
  manualRect = null;
  manualStart = null;
  manualImage = null;
  clearManualCropFields();
  clearManualCanvas();
  document.getElementById('manualCropLabel').textContent = manualSourceMode === 'original'
    ? 'Loading original source image for: ' + manualRel
    : 'Loading already-cropped ROI image for: ' + manualRel;
  manualImage = new Image();
  manualImage.onload = () => {{
    document.getElementById('manualCropLabel').textContent = manualSourceMode === 'original'
      ? 'Drag a rectangle inside the original source image: ' + manualRel
      : 'Drag a rectangle inside this already-cropped ROI image: ' + manualRel;
    drawManualCanvas();
  }};
  manualImage.onerror = () => {{
    document.getElementById('manualCropLabel').textContent = manualSourceMode === 'original'
      ? 'Could not load the original source image for this ROI.'
      : 'Could not load the already-cropped ROI image.';
  }};
  manualImage.src = '/manual_crop_png?rel=' + encodeURIComponent(manualRel) + '&source=' + encodeURIComponent(manualSourceMode) + '&v=' + Date.now();
}}

function openManualCrop(rel) {{
  manualRel = rel;
  document.getElementById('manualSource').value = 'crop';
  document.getElementById('manualCropModal').classList.add('open');
  loadManualCropImage();
}}

const manualSourceSelect = document.getElementById('manualSource');
if (manualSourceSelect) manualSourceSelect.addEventListener('change', loadManualCropImage);

const manualCanvas = document.getElementById('manualCropCanvas');
manualCanvas.addEventListener('mousedown', (event) => {{
  manualStart = canvasPoint(event);
  manualRect = {{ x: manualStart.x, y: manualStart.y, w: 1, h: 1 }};
  drawManualCanvas();
}});
manualCanvas.addEventListener('mousemove', (event) => {{
  if (!manualStart) return;
  const point = canvasPoint(event);
  manualRect = {{
    x: Math.max(0, Math.min(manualStart.x, point.x)),
    y: Math.max(0, Math.min(manualStart.y, point.y)),
    w: Math.abs(point.x - manualStart.x),
    h: Math.abs(point.y - manualStart.y)
  }};
  syncManualFields();
  drawManualCanvas();
}});
manualCanvas.addEventListener('mouseup', () => {{
  manualStart = null;
  syncManualFields();
}});

async function submitManualCrop() {{
  const saveButton = document.getElementById('saveManualCropButton');
  const width = Number(document.getElementById('manualW').value || 0);
  const height = Number(document.getElementById('manualH').value || 0);
  if (width < 5 || height < 5) {{
    document.getElementById('manualCropLabel').textContent = 'Crop is too small. Drag a larger rectangle inside the cropped ROI image before saving.';
    return;
  }}
  saveButton.disabled = true;
  saveButton.textContent = 'Saving...';
  document.getElementById('manualCropLabel').textContent = 'Saving corrected crop...';
  const body = new URLSearchParams();
  body.set('rel', manualRel);
  body.set('x', document.getElementById('manualX').value);
  body.set('y', document.getElementById('manualY').value);
  body.set('width', String(width));
  body.set('height', String(height));
  body.set('source', manualSourceMode);
  try {{
    const response = await fetch('/manual_recrop', {{ method: 'POST', body }});
    if (!response.ok) {{
      const payload = await response.json().catch(() => null);
      document.getElementById('manualCropLabel').textContent = payload && payload.error ? payload.error : 'Manual crop save failed.';
      saveButton.disabled = false;
      saveButton.textContent = 'Save corrected crop';
      return;
    }}
    const payload = await response.json().catch(() => null);
    updateCardAfterManualCrop(payload && payload.rel ? payload.rel : manualRel, payload && payload.thumb_url ? payload.thumb_url : ('/thumb?rel=' + encodeURIComponent(manualRel) + '&v=' + Date.now()));
    closeManualCrop();
    saveButton.disabled = false;
    saveButton.textContent = 'Save corrected crop';
  }} catch (error) {{
    saveButton.disabled = false;
    saveButton.textContent = 'Save corrected crop';
    document.getElementById('manualCropLabel').textContent = 'Save failed: ' + error;
  }}
}}

const saveManualCropButton = document.getElementById('saveManualCropButton');
if (saveManualCropButton) saveManualCropButton.addEventListener('click', submitManualCrop);

let fixedAnnotationRel = '';
let fixedAnnotationButton = null;
let fixedAnnotationImages = {{}};
let fixedAnnotationDisplayCanvases = {{}};
let fixedAnnotationRoi = [];
let fixedAnnotationFoci = [];
let fixedAnnotationCondensates = [];
let fixedAnnotationBackgrounds = [];
let fixedAnnotationTool = 'roi';
let fixedAnnotationDrawing = false;
let fixedAnnotationCurrent = [];
let fixedAnnotationGestureMoved = false;
let fixedAnnotationSourceMode = 'crop';
let fixedAnnotationPurpose = 'fixed';
let fixedAnnotationCropBox = null;

function fixedAnnotationViewportFor(image) {{
  const fullWidth = image ? image.naturalWidth : 1;
  const fullHeight = image ? image.naturalHeight : 1;
  const drawingView = document.getElementById('fixedCellDrawingView').value;
  if (fixedAnnotationSourceMode === 'crop' || drawingView === 'full' || !fixedAnnotationCropBox) {{
    return {{ x0: 0, y0: 0, width: fullWidth, height: fullHeight }};
  }}
  const box = fixedAnnotationCropBox;
  const padX = Math.max(24, box.width * 0.45);
  const padY = Math.max(24, box.height * 0.45);
  const x0 = Math.max(0, Math.floor(box.x0 - padX));
  const y0 = Math.max(0, Math.floor(box.y0 - padY));
  const x1 = Math.min(fullWidth, Math.ceil(box.x1 + padX));
  const y1 = Math.min(fullHeight, Math.ceil(box.y1 + padY));
  return {{ x0, y0, width: Math.max(1, x1 - x0), height: Math.max(1, y1 - y0) }};
}}

function buildFixedAnnotationDisplayCanvas(image, channel) {{
  const output = document.createElement('canvas');
  output.width = image.naturalWidth;
  output.height = image.naturalHeight;
  const ctx = output.getContext('2d', {{ willReadFrequently: true }});
  ctx.drawImage(image, 0, 0);
  if (channel === 'merged') return output;

  const displayMode = document.getElementById('fixedCellChannelDisplay').value;
  const contrastMode = document.getElementById('fixedCellDisplayContrast').value;
  const pixels = ctx.getImageData(0, 0, output.width, output.height);
  const data = pixels.data;
  let low = 0;
  let high = 255;
  if (contrastMode === 'auto') {{
    const histogram = new Uint32Array(256);
    for (let i = 0; i < data.length; i += 4) {{
      const gray = Math.max(data[i], data[i + 1], data[i + 2]);
      histogram[gray] += 1;
    }}
    const total = output.width * output.height;
    const lowTarget = Math.floor(total * 0.005);
    const highTarget = Math.floor(total * 0.995);
    let cumulative = 0;
    for (let value = 0; value < 256; value += 1) {{
      cumulative += histogram[value];
      if (cumulative >= lowTarget) {{ low = value; break; }}
    }}
    cumulative = 0;
    for (let value = 0; value < 256; value += 1) {{
      cumulative += histogram[value];
      if (cumulative >= highTarget) {{ high = value; break; }}
    }}
    if (high <= low + 2) {{ low = 0; high = 255; }}
  }}
  const scale = 255 / Math.max(1, high - low);
  for (let i = 0; i < data.length; i += 4) {{
    const sourceGray = Math.max(data[i], data[i + 1], data[i + 2]);
    const gray = Math.max(0, Math.min(255, Math.round((sourceGray - low) * scale)));
    if (displayMode === 'pseudocolor') {{
      data[i] = channel === 'red' ? gray : 0;
      data[i + 1] = channel === 'green' ? gray : 0;
      data[i + 2] = channel === 'blue' ? gray : 0;
    }} else {{
      data[i] = gray;
      data[i + 1] = gray;
      data[i + 2] = gray;
    }}
  }}
  ctx.putImageData(pixels, 0, 0);
  return output;
}}

function rebuildFixedAnnotationDisplayCanvases() {{
  fixedAnnotationDisplayCanvases = {{}};
  ['merged', 'red', 'green', 'blue'].forEach((channel) => {{
    const image = fixedAnnotationImages[channel];
    if (image && image.complete && image.naturalWidth) {{
      fixedAnnotationDisplayCanvases[channel] = buildFixedAnnotationDisplayCanvas(image, channel);
    }}
  }});
}}

function configureAnnotationPurpose(purpose) {{
  fixedAnnotationPurpose = purpose === 'partition' ? 'partition' : 'fixed';
  const partition = fixedAnnotationPurpose === 'partition';
  document.getElementById('fixedCellAnnotationTitle').textContent = partition
    ? 'Draw measurement ROI, condensates, and cell-free backgrounds'
    : 'Draw measurement ROI, MCP/MS2 focus, and condensates';
  document.getElementById('fixedCellAnnotationHelp').textContent = partition
    ? 'Draw the thin cyan measurement/reference ROI, every thin green condensate ROI, and one or more orange cell-free background ROIs. Background ROIs may be outside the yellow reviewed-cell box. If the selected cell truly has no condensate, check No condensate in this cell.'
    : 'Choose the already-cropped ROI or the original raw image before drawing. Draw on any split image; coordinates are synchronized. Cyan is the measurement ROI, magenta is MCP/MS2 focus, green is condensates, and orange is optional background.';
  ['addMcpFocusButton', 'undoMcpFocusButton'].forEach((id) => {{ document.getElementById(id).style.display = partition ? 'none' : ''; }});
  document.getElementById('drawMs2BackgroundButton').style.display = '';
  document.getElementById('removeMs2BackgroundButton').style.display = '';
  document.getElementById('drawMs2BackgroundButton').textContent = partition ? 'Add cell-free background ROI' : 'Draw/replace MS2 background ROI (optional)';
  document.getElementById('removeMs2BackgroundButton').textContent = partition ? 'Undo last background ROI' : 'Remove manual background — use automatic';
  document.getElementById('partitionNoCondensateWrap').style.display = partition ? 'inline-flex' : 'none';
  document.getElementById('saveFixedCellAnnotationButton').textContent = partition
    ? 'Save ROI + condensates + backgrounds'
    : 'Save ROI + focus + condensates + background';
}}

function closeFixedCellAnnotation() {{
  document.getElementById('fixedCellAnnotationModal').classList.remove('open');
}}

function fixedAnnotationCanvasPoint(event, canvas) {{
  const box = canvas.getBoundingClientRect();
  const channel = canvas.dataset.channel || 'merged';
  const image = fixedAnnotationImages[channel];
  const viewport = fixedAnnotationViewportFor(image);
  let x = viewport.x0 + Math.max(0, Math.min(canvas.width - 1, (event.clientX - box.left) * canvas.width / box.width));
  let y = viewport.y0 + Math.max(0, Math.min(canvas.height - 1, (event.clientY - box.top) * canvas.height / box.height));
  if (fixedAnnotationSourceMode !== 'crop' && fixedAnnotationCropBox && !(fixedAnnotationPurpose === 'partition' && fixedAnnotationTool === 'background')) {{
    x = Math.max(fixedAnnotationCropBox.x0, Math.min(fixedAnnotationCropBox.x1 - 1, x));
    y = Math.max(fixedAnnotationCropBox.y0, Math.min(fixedAnnotationCropBox.y1 - 1, y));
  }}
  return [x, y];
}}

function drawAnnotationPolygon(ctx, points, stroke, fill, closeShape = true) {{
  if (!points || points.length < 1) return;
  ctx.beginPath();
  ctx.moveTo(points[0][0], points[0][1]);
  for (let i = 1; i < points.length; i += 1) ctx.lineTo(points[i][0], points[i][1]);
  if (closeShape && points.length >= 3) ctx.closePath();
  ctx.strokeStyle = stroke;
  ctx.lineWidth = 1;
  ctx.lineJoin = 'round';
  ctx.lineCap = 'round';
  ctx.stroke();
}}

function drawFixedCellAnnotationCanvas() {{
  document.querySelectorAll('.fixed-cell-annotation-canvas').forEach((canvas) => {{
    const channel = canvas.dataset.channel || 'merged';
    const image = fixedAnnotationImages[channel];
    if (!image) return;
    const ctx = canvas.getContext('2d');
    const displayCanvas = fixedAnnotationDisplayCanvases[channel];
    const viewport = fixedAnnotationViewportFor(image);
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    ctx.drawImage(
      displayCanvas || image,
      viewport.x0, viewport.y0, viewport.width, viewport.height,
      0, 0, viewport.width, viewport.height
    );
    ctx.save();
    ctx.translate(-viewport.x0, -viewport.y0);
    if (fixedAnnotationSourceMode !== 'crop' && fixedAnnotationCropBox) {{
      const b = fixedAnnotationCropBox;
      ctx.save();
      ctx.strokeStyle = '#ffe600';
      ctx.lineWidth = Math.max(2, ctx.canvas.width / 900);
      ctx.setLineDash([Math.max(6, ctx.canvas.width / 220), Math.max(4, ctx.canvas.width / 320)]);
      ctx.strokeRect(b.x0, b.y0, b.width, b.height);
      ctx.setLineDash([]);
      ctx.restore();
    }}
    drawAnnotationPolygon(ctx, fixedAnnotationRoi, '#00dcff', 'rgba(0,220,255,0.035)');
    fixedAnnotationFoci.forEach((points) => drawAnnotationPolygon(ctx, points, '#ff00ff', 'rgba(255,0,255,0.055)'));
    fixedAnnotationCondensates.forEach((points) => drawAnnotationPolygon(ctx, points, '#00ff50', 'rgba(0,255,80,0.045)'));
    fixedAnnotationBackgrounds.forEach((points) => drawAnnotationPolygon(ctx, points, '#ff8c00', 'rgba(255,140,0,0.055)'));
    if (fixedAnnotationCurrent.length) {{
      const isRoi = fixedAnnotationTool === 'roi';
      const isBackground = fixedAnnotationTool === 'background';
      const isCondensate = fixedAnnotationTool === 'condensate';
      drawAnnotationPolygon(
        ctx,
        fixedAnnotationCurrent,
        isRoi ? '#00dcff' : (isBackground ? '#ff8c00' : (isCondensate ? '#00ff50' : '#ff00ff')),
        isRoi ? 'rgba(0,220,255,0.035)' : (isBackground ? 'rgba(255,140,0,0.055)' : (isCondensate ? 'rgba(0,255,80,0.045)' : 'rgba(255,0,255,0.055)')),
        false
      );
    }}
    ctx.restore();
  }});
}}

function setFixedAnnotationTool(tool) {{
  fixedAnnotationTool = tool;
  document.getElementById('drawMeasurementRoiButton').classList.toggle('annotation-tool-active', tool === 'roi');
  document.getElementById('addMcpFocusButton').classList.toggle('annotation-tool-active', tool === 'focus');
  document.getElementById('addCondensateButton').classList.toggle('annotation-tool-active', tool === 'condensate');
  document.getElementById('drawMs2BackgroundButton').classList.toggle('annotation-tool-active', tool === 'background');
  document.getElementById('fixedCellAnnotationLabel').textContent = tool === 'roi'
    ? 'ROI tool active: draw the cyan measurement ROI.'
    : (tool === 'background'
      ? (fixedAnnotationPurpose === 'partition'
        ? 'Cell-free background tool active: draw an orange ROI outside the cell. Repeat to add more backgrounds.'
        : 'Background tool active: draw one orange ROI inside the measurement ROI, away from bright MS2 foci.')
      : (tool === 'condensate'
        ? 'Condensate tool active: trace one green condensate ROI. Repeat to add more.'
        : 'MCP/MS2 focus tool active: draw a magenta focus. Repeat to add more.'));
}}

function commitFixedAnnotationCurrent() {{
  const label = document.getElementById('fixedCellAnnotationLabel');
  if (fixedAnnotationCurrent.length < 3) {{
    label.textContent = 'Add at least three boundary points before finishing this shape.';
    return false;
  }}
  const completed = fixedAnnotationCurrent.slice();
  fixedAnnotationCurrent = [];
  if (fixedAnnotationTool === 'roi') {{
    fixedAnnotationRoi = completed;
    setFixedAnnotationTool(fixedAnnotationPurpose === 'partition' ? 'condensate' : 'focus');
  }} else if (fixedAnnotationTool === 'background') {{
    if (fixedAnnotationPurpose === 'partition') fixedAnnotationBackgrounds.push(completed);
    else fixedAnnotationBackgrounds = [completed];
    label.textContent = fixedAnnotationPurpose === 'partition'
      ? 'Cell-free background ROI added. Add another background or save all drawings.'
      : 'Manual MS2 background ROI saved in the editor. Save all drawings when ready.';
  }} else if (fixedAnnotationTool === 'condensate') {{
    fixedAnnotationCondensates.push(completed);
    label.textContent = 'Condensate ROI saved in the editor. Add another condensate or save all drawings.';
  }} else {{
    fixedAnnotationFoci.push(completed);
    label.textContent = 'MCP/MS2 focus saved in the editor. Add another focus or save all drawings.';
  }}
  drawFixedCellAnnotationCanvas();
  return true;
}}

function loadFixedAnnotationImages(rel, sourceMode) {{
  fixedAnnotationImages = {{}};
  fixedAnnotationDisplayCanvases = {{}};
  const channels = ['merged', 'red', 'green', 'blue'];
  let loadedChannels = 0;
  channels.forEach((channel) => {{
    const image = new Image();
    fixedAnnotationImages[channel] = image;
    image.onload = () => {{
      loadedChannels += 1;
      fixedAnnotationDisplayCanvases[channel] = buildFixedAnnotationDisplayCanvas(image, channel);
      drawFixedCellAnnotationCanvas();
      if (loadedChannels === channels.length) setFixedAnnotationTool(
        fixedAnnotationRoi.length >= 3 ? (fixedAnnotationPurpose === 'partition' ? 'condensate' : 'focus') : 'roi'
      );
    }};
    image.onerror = () => {{ document.getElementById('fixedCellAnnotationLabel').textContent = 'Could not load the ' + channel + ' view from the selected image source.'; }};
    image.src = '/manual_crop_png?rel=' + encodeURIComponent(rel) + '&source=' + encodeURIComponent(sourceMode) + '&channel=' + encodeURIComponent(channel) + '&v=' + Date.now();
  }});
}}

async function openFixedCellAnnotation(rel, button) {{
  configureAnnotationPurpose('fixed');
  document.getElementById('fixedCellSaveStatus').style.display = 'none';
  fixedAnnotationRel = rel;
  fixedAnnotationButton = button || null;
  fixedAnnotationImages = {{}};
  fixedAnnotationRoi = [];
  fixedAnnotationFoci = [];
  fixedAnnotationCondensates = [];
  fixedAnnotationBackgrounds = [];
  document.getElementById('partitionNoCondensate').checked = false;
  fixedAnnotationCurrent = [];
  document.getElementById('fixedCellAnnotationModal').classList.add('open');
  document.getElementById('fixedCellAnnotationLabel').textContent = 'Loading crop and saved drawings...';
  try {{
    const response = await fetch('/api/fixed_cell_annotation?rel=' + encodeURIComponent(rel) + '&v=' + Date.now());
    const payload = await response.json();
    if (!response.ok || !payload.ok) throw new Error(payload.error || 'Could not load annotation');
    const annotation = payload.annotation || {{}};
    fixedAnnotationSourceMode = ['crop', 'original', 'untouched'].includes(annotation.source_mode) ? annotation.source_mode : (payload.default_source_mode || 'crop');
    fixedAnnotationCropBox = payload.reviewed_crop_box || annotation.reviewed_crop_box || null;
    document.getElementById('fixedCellAnnotationSource').value = fixedAnnotationSourceMode;
    fixedAnnotationRoi = Array.isArray(annotation.roi_points) ? annotation.roi_points : [];
    fixedAnnotationFoci = Array.isArray(annotation.focus_polygons) ? annotation.focus_polygons : [];
    fixedAnnotationCondensates = Array.isArray(annotation.condensate_polygons) ? annotation.condensate_polygons : [];
    fixedAnnotationBackgrounds = Array.isArray(annotation.background_polygon) && annotation.background_polygon.length ? [annotation.background_polygon] : [];
    loadFixedAnnotationImages(rel, fixedAnnotationSourceMode);
  }} catch (error) {{
    document.getElementById('fixedCellAnnotationLabel').textContent = 'Load failed: ' + String(error.message || error);
  }}
}}

async function openPartitionAnnotation(rel, button) {{
  configureAnnotationPurpose('partition');
  document.getElementById('fixedCellSaveStatus').style.display = 'none';
  fixedAnnotationRel = rel;
  fixedAnnotationButton = button || null;
  fixedAnnotationImages = {{}};
  fixedAnnotationRoi = [];
  fixedAnnotationFoci = [];
  fixedAnnotationCondensates = [];
  fixedAnnotationBackgrounds = [];
  document.getElementById('partitionNoCondensate').checked = false;
  fixedAnnotationCurrent = [];
  document.getElementById('fixedCellAnnotationModal').classList.add('open');
  document.getElementById('fixedCellAnnotationLabel').textContent = 'Loading original image and saved partition drawings...';
  try {{
    const response = await fetch('/api/partition_annotation?rel=' + encodeURIComponent(rel) + '&v=' + Date.now());
    const payload = await response.json();
    if (!response.ok || !payload.ok) throw new Error(payload.error || 'Could not load partition annotation');
    const annotation = payload.annotation || {{}};
    fixedAnnotationSourceMode = ['crop', 'original', 'untouched'].includes(annotation.source_mode) ? annotation.source_mode : (payload.default_source_mode || 'untouched');
    fixedAnnotationCropBox = payload.reviewed_crop_box || annotation.reviewed_crop_box || null;
    document.getElementById('fixedCellAnnotationSource').value = fixedAnnotationSourceMode;
    fixedAnnotationRoi = Array.isArray(annotation.roi_points) ? annotation.roi_points : [];
    fixedAnnotationCondensates = Array.isArray(annotation.condensate_polygons) ? annotation.condensate_polygons : [];
    fixedAnnotationBackgrounds = Array.isArray(annotation.background_polygons) ? annotation.background_polygons : [];
    document.getElementById('partitionNoCondensate').checked = Boolean(annotation.no_condensate);
    if (payload.seeded_from_fixed_cell) {{
      document.getElementById('fixedCellAnnotationLabel').textContent = 'Existing fixed-cell nucleus/condensates copied into this editor only. Review and save to create an independent partition annotation.';
    }}
    loadFixedAnnotationImages(rel, fixedAnnotationSourceMode);
  }} catch (error) {{
    document.getElementById('fixedCellAnnotationLabel').textContent = 'Load failed: ' + String(error.message || error);
  }}
}}

document.getElementById('fixedCellAnnotationSource').addEventListener('change', (event) => {{
  fixedAnnotationSourceMode = ['crop', 'original', 'untouched'].includes(event.target.value) ? event.target.value : 'crop';
  fixedAnnotationRoi = [];
  fixedAnnotationFoci = [];
  fixedAnnotationCondensates = [];
  fixedAnnotationBackgrounds = [];
  document.getElementById('partitionNoCondensate').checked = false;
  fixedAnnotationCurrent = [];
  document.getElementById('fixedCellAnnotationLabel').textContent = fixedAnnotationPurpose === 'partition'
    ? 'Image source changed. Draw a new measurement/reference ROI and condensate ROIs on this source.'
    : (fixedAnnotationSourceMode === 'untouched'
      ? 'Untouched acquisition loaded. The yellow box is the exact PCC/Count-Size reviewed cell crop; draw only inside it.'
      : 'Image source changed. Draw a new ROI and MCP/MS2 focus on this source.');
  loadFixedAnnotationImages(fixedAnnotationRel, fixedAnnotationSourceMode);
}});

['fixedCellChannelDisplay', 'fixedCellDisplayContrast'].forEach((id) => {{
  document.getElementById(id).addEventListener('change', () => {{
    rebuildFixedAnnotationDisplayCanvases();
    drawFixedCellAnnotationCanvas();
    document.getElementById('fixedCellAnnotationLabel').textContent =
      'Display changed for easier drawing only. Saved coordinates and untouched grayscale measurement pixels are unchanged.';
  }});
}});

document.getElementById('fixedCellDrawingView').addEventListener('change', () => {{
  drawFixedCellAnnotationCanvas();
  document.getElementById('fixedCellAnnotationLabel').textContent =
    document.getElementById('fixedCellDrawingView').value === 'reviewed'
      ? 'Reviewed-cell zoom is active on all four synchronized panels. Coordinates still map to the untouched full acquisition.'
      : 'Full-acquisition view is active. Switch to Enlarge reviewed cell whenever the target is too small to trace accurately.';
}});

function finishFixedAnnotationDrawing(event) {{
  if (!fixedAnnotationDrawing) return;
  event.preventDefault();
  fixedAnnotationDrawing = false;
  if (fixedAnnotationGestureMoved && fixedAnnotationCurrent.length >= 3) {{
    const canvas = event.currentTarget;
    const first = fixedAnnotationCurrent[0];
    const last = fixedAnnotationCurrent[fixedAnnotationCurrent.length - 1];
    const closeTolerance = Math.max(8, Math.min(canvas.width, canvas.height) * 0.02);
    const closingDistance = Math.hypot(last[0] - first[0], last[1] - first[1]);
    if (closingDistance <= closeTolerance) {{
      fixedAnnotationCurrent[fixedAnnotationCurrent.length - 1] = [first[0], first[1]];
      commitFixedAnnotationCurrent();
    }} else {{
      document.getElementById('fixedCellAnnotationLabel').textContent =
        'The outline is still open. Continue tracing from the endpoint and return to the starting point; the app will snap them together without drawing a line across the ROI.';
    }}
  }} else {{
    document.getElementById('fixedCellAnnotationLabel').textContent =
      'Boundary point ' + fixedAnnotationCurrent.length + ' added. Add more points, then click Finish current shape.';
  }}
  drawFixedCellAnnotationCanvas();
}}
document.querySelectorAll('.fixed-cell-annotation-canvas').forEach((canvas) => {{
  canvas.addEventListener('pointerdown', (event) => {{
    const channel = canvas.dataset.channel || 'merged';
    if (!fixedAnnotationImages[channel]) return;
    event.preventDefault();
    canvas.setPointerCapture(event.pointerId);
    fixedAnnotationDrawing = true;
    fixedAnnotationGestureMoved = false;
    const point = fixedAnnotationCanvasPoint(event, canvas);
    const last = fixedAnnotationCurrent[fixedAnnotationCurrent.length - 1];
    if (!last || Math.hypot(point[0] - last[0], point[1] - last[1]) >= 1) fixedAnnotationCurrent.push(point);
    drawFixedCellAnnotationCanvas();
  }});
  canvas.addEventListener('pointermove', (event) => {{
    if (!fixedAnnotationDrawing) return;
    event.preventDefault();
    const point = fixedAnnotationCanvasPoint(event, canvas);
    const last = fixedAnnotationCurrent[fixedAnnotationCurrent.length - 1];
    if (!last || Math.hypot(point[0] - last[0], point[1] - last[1]) >= 2) {{
      fixedAnnotationCurrent.push(point);
      fixedAnnotationGestureMoved = true;
    }}
    drawFixedCellAnnotationCanvas();
  }});
  canvas.addEventListener('pointerup', finishFixedAnnotationDrawing);
  canvas.addEventListener('pointercancel', finishFixedAnnotationDrawing);
  canvas.addEventListener('dblclick', (event) => {{
    event.preventDefault();
    commitFixedAnnotationCurrent();
  }});
}});

document.getElementById('drawMeasurementRoiButton').addEventListener('click', () => {{
  fixedAnnotationCurrent = [];
  setFixedAnnotationTool('roi');
  drawFixedCellAnnotationCanvas();
}});
document.getElementById('addMcpFocusButton').addEventListener('click', () => {{
  if (fixedAnnotationCurrent.length >= 3) commitFixedAnnotationCurrent();
  fixedAnnotationCurrent = [];
  setFixedAnnotationTool('focus');
  drawFixedCellAnnotationCanvas();
}});
document.getElementById('addCondensateButton').addEventListener('click', () => {{
  if (fixedAnnotationCurrent.length >= 3) commitFixedAnnotationCurrent();
  fixedAnnotationCurrent = [];
  setFixedAnnotationTool('condensate');
  drawFixedCellAnnotationCanvas();
}});
document.getElementById('undoCondensateButton').addEventListener('click', () => {{
  fixedAnnotationCondensates.pop();
  drawFixedCellAnnotationCanvas();
}});
document.getElementById('drawMs2BackgroundButton').addEventListener('click', () => {{
  if (fixedAnnotationCurrent.length >= 3) commitFixedAnnotationCurrent();
  fixedAnnotationCurrent = [];
  setFixedAnnotationTool('background');
  drawFixedCellAnnotationCanvas();
}});
document.getElementById('removeMs2BackgroundButton').addEventListener('click', () => {{
  if (fixedAnnotationPurpose === 'partition') fixedAnnotationBackgrounds.pop();
  else fixedAnnotationBackgrounds = [];
  if (fixedAnnotationTool === 'background') fixedAnnotationCurrent = [];
  setFixedAnnotationTool(fixedAnnotationFoci.length ? 'focus' : 'roi');
  document.getElementById('fixedCellAnnotationLabel').textContent = fixedAnnotationPurpose === 'partition'
    ? 'Last cell-free background ROI removed.'
    : 'Manual background removed in the editor. Save the drawings to use automatic background in the next analysis.';
  drawFixedCellAnnotationCanvas();
}});
document.getElementById('finishAnnotationShapeButton').addEventListener('click', commitFixedAnnotationCurrent);
document.getElementById('undoMcpFocusButton').addEventListener('click', () => {{
  fixedAnnotationFoci.pop();
  drawFixedCellAnnotationCanvas();
}});
document.getElementById('clearFixedCellAnnotationButton').addEventListener('click', () => {{
  fixedAnnotationRoi = [];
  fixedAnnotationFoci = [];
  fixedAnnotationCondensates = [];
  fixedAnnotationBackgrounds = [];
  document.getElementById('partitionNoCondensate').checked = false;
  fixedAnnotationCurrent = [];
  setFixedAnnotationTool('roi');
  drawFixedCellAnnotationCanvas();
}});

function showFixedAnnotationSaveStatus(message, isError) {{
  const label = document.getElementById('fixedCellAnnotationLabel');
  const status = document.getElementById('fixedCellSaveStatus');
  label.textContent = message;
  status.textContent = message;
  status.style.display = 'block';
  status.style.background = isError ? '#fff1f2' : '#eefbf3';
  status.style.borderColor = isError ? '#dc2626' : '#16803a';
  status.style.color = isError ? '#991b1b' : '#14532d';
}}

document.getElementById('saveFixedCellAnnotationButton').addEventListener('click', async () => {{
  const button = document.getElementById('saveFixedCellAnnotationButton');
  if (fixedAnnotationCurrent.length >= 3) commitFixedAnnotationCurrent();
  if (fixedAnnotationRoi.length < 3) {{ showFixedAnnotationSaveStatus('Not saved: draw and finish the cyan measurement ROI first.', true); return; }}
  if (fixedAnnotationPurpose === 'fixed' && !fixedAnnotationFoci.length) {{ showFixedAnnotationSaveStatus('Not saved: draw and finish at least one magenta MCP/MS2 focus.', true); return; }}
  const noCondensate = document.getElementById('partitionNoCondensate').checked;
  if (fixedAnnotationPurpose === 'partition' && !fixedAnnotationCondensates.length && !noCondensate) {{ showFixedAnnotationSaveStatus('Not saved: draw and finish at least one green condensate ROI, or check No condensate in this cell.', true); return; }}
  if (fixedAnnotationPurpose === 'partition' && fixedAnnotationCondensates.length && noCondensate) {{ showFixedAnnotationSaveStatus('Not saved: remove the green condensate drawings or uncheck No condensate in this cell.', true); return; }}
  button.disabled = true;
  showFixedAnnotationSaveStatus('Saving exact image coordinates...', false);
  const body = new URLSearchParams();
  body.set('rel', fixedAnnotationRel);
  body.set('roi_points', JSON.stringify(fixedAnnotationRoi));
  body.set('focus_polygons', JSON.stringify(fixedAnnotationFoci));
  body.set('condensate_polygons', JSON.stringify(fixedAnnotationCondensates));
  body.set('background_polygon', JSON.stringify(fixedAnnotationBackgrounds.length ? fixedAnnotationBackgrounds[0] : []));
  body.set('background_polygons', JSON.stringify(fixedAnnotationBackgrounds));
  body.set('no_condensate', document.getElementById('partitionNoCondensate').checked ? '1' : '0');
  body.set('source_mode', fixedAnnotationSourceMode);
  try {{
    const endpoint = fixedAnnotationPurpose === 'partition' ? '/save_partition_annotation' : '/save_fixed_cell_annotation';
    const response = await fetch(endpoint, {{ method: 'POST', body }});
    const payload = await response.json();
    if (!response.ok || !payload.ok) throw new Error(payload.error || 'Save failed');
    if (fixedAnnotationButton) fixedAnnotationButton.textContent = fixedAnnotationPurpose === 'partition'
      ? 'Edit measurement ROI + condensates'
      : 'Edit saved ROI + MCP focus + condensates';
    showFixedAnnotationSaveStatus(
      fixedAnnotationPurpose === 'partition'
        ? 'Saved. The measurement ROI, condensates, and backgrounds are stored. You may continue editing or close this window.'
        : 'Saved. The measurement ROI and MCP/MS2 drawings are stored. You may continue editing or close this window.',
      false
    );
  }} catch (error) {{
    showFixedAnnotationSaveStatus('Save failed: ' + String(error.message || error), true);
  }} finally {{
    button.disabled = false;
  }}
}});

function cardForRel(rel) {{
  const cards = document.querySelectorAll('.review-card');
  for (const card of cards) {{
    if (card.dataset.rel === rel) return card;
  }}
  return null;
}}

function setCardAutocleanNeeded(card, needed) {{
  if (!card) return;
  let notice = card.querySelector('[data-role="autoclean-needed"]');
  if (needed) {{
    if (!notice) {{
      notice = document.createElement('div');
      notice.className = 'autoclean-needed';
      notice.dataset.role = 'autoclean-needed';
      notice.textContent = 'Autoclean needed for this crop. Click Autoclean only this crop.';
      const banner = card.querySelector('.status-banner');
      if (banner) banner.insertAdjacentElement('afterend', notice);
      else card.appendChild(notice);
    }}
    notice.style.display = '';
  }} else if (notice) {{
    notice.remove();
  }}
}}

function updateCardAfterManualCrop(rel, thumbUrl) {{
  const card = cardForRel(rel);
  if (!card) return;
  const label = card.querySelector('[data-role="chosen-label"]');
  const panel = card.querySelector('[data-role="chosen-panel"]');
  const banner = card.querySelector('.status-banner');
  if (label) label.textContent = 'Selected: Manual corrected crop';
  if (panel) {{
    panel.innerHTML = '<img src="' + thumbUrl + '" loading="lazy" alt="Selected crop preview" />';
  }}
  if (!card.querySelector('.undo-manual-button')) {{
    const recropButton = card.querySelector('.manual-recrop-button');
    const undo = document.createElement('button');
    undo.className = 'btn secondary undo-manual-button';
    undo.type = 'button';
    undo.dataset.rel = rel;
    undo.textContent = 'Undo manual recrop';
    undo.addEventListener('click', () => undoManualRecrop(rel, undo));
    recropButton.insertAdjacentElement('afterend', undo);
  }}
  if (banner) instantSetStatus(card.id, banner.id, 'keep');
  setCardAutocleanNeeded(card, true);
  updateReviewDependentSections();
  refreshRecoveredCropsSection();
  refreshCleaningPreviewSection(false);
}}

async function undoManualRecrop(rel, button) {{
  button.disabled = true;
  button.textContent = 'Undoing...';
  const body = new URLSearchParams();
  body.set('rel', rel);
  try {{
    const response = await fetch('/undo_manual_recrop', {{ method: 'POST', body }});
    const payload = await response.json().catch(() => null);
    if (!response.ok) throw new Error(payload && payload.error ? payload.error : 'Undo failed');
    const card = cardForRel(rel);
    if (card) {{
      const label = card.querySelector('[data-role="chosen-label"]');
      const panel = card.querySelector('[data-role="chosen-panel"]');
      const banner = card.querySelector('.status-banner');
      if (label) label.textContent = 'Selected: Raw cropped ROI';
      if (panel) {{
        panel.innerHTML = '<img src="' + (payload.thumb_url || ('/thumb?rel=' + encodeURIComponent(rel) + '&v=' + Date.now())) + '" loading="lazy" alt="Selected crop preview" />';
      }}
      if (banner) instantSetStatus(card.id, banner.id, 'keep');
      setCardAutocleanNeeded(card, true);
      updateReviewDependentSections();
      refreshRecoveredCropsSection();
      refreshCleaningPreviewSection(false);
    }}
    button.remove();
  }} catch (error) {{
    button.disabled = false;
    button.textContent = 'Undo manual recrop';
    alert(error);
  }}
}}

function instantToggle(cardId, badgeId, buttonId) {{
  const card = document.getElementById(cardId);
  const badge = document.getElementById(badgeId);
  const button = document.getElementById(buttonId);
  const form = button.closest('form');
  const statusInput = form.querySelector('.status-input');
  const oldStatus = card.dataset.status || 'keep';
  const newStatus = statusInput.value;
  const isKeep = newStatus === 'keep';
  const color = isKeep ? '#188038' : '#b3261e';

  card.dataset.status = newStatus;
  card.style.borderColor = color;
  card.style.boxShadow = isKeep ? '0 0 0 0 transparent' : '0 0 0 5px rgba(179,38,30,0.22)';
  badge.style.background = color;
  badge.textContent = isKeep ? 'KEEP' : 'REJECTED';
  button.textContent = isKeep ? 'Reject this ROI' : 'Keep this ROI';
  statusInput.value = isKeep ? 'reject' : 'keep';
  updateCounts(oldStatus, newStatus);
}}

if ({str(bool(items)).lower()} && window.location.pathname === '/use_local_raw_folder_path') {{
  window.requestAnimationFrame(() => {{
    const cropReview = document.getElementById('crop-review');
    if (cropReview) cropReview.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
  }});
}}
</script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def send_html(self, content, status=200):
        data = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload, status=200):
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_png(self, data):
        self.send_response(200)
        self.send_header("Content-Type", "image/png")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        state = load_state()
        parsed = urlparse(self.path)
        if parsed.path == "/health":
            items = state.get("items", {})
            self.send_json({
                "ok": True,
                "app": "auto-cleaned-fiji-pcc-manders-count-size",
                "build": APP_BUILD_ID,
                "loaded_crops": len(items),
                "kept": sum(1 for x in items.values() if x.get("status", "keep") == "keep"),
                "rejected": sum(1 for x in items.values() if x.get("status", "keep") == "reject"),
                "crop_root": state.get("crop_root", ""),
                "state_file": STATE_FILE,
            })
            return
        if parsed.path == "/api/cleaning_card":
            try:
                qs = parse_qs(parsed.query)
                rel = unquote(qs.get("rel", [""])[0])
                items = ensure_crop_numbers(state)
                item = next((candidate for candidate in items if candidate.get("rel_path") == rel), None)
                if not item:
                    self.send_json({"ok": False, "error": "Crop not found."}, status=404)
                    return
                html, _ = render_cleaning_cards(state, [item])
                self.send_json({
                    "ok": True,
                    "rel": rel,
                    "html": html if item.get("status", "keep") == "keep" else "",
                })
            except Exception:
                self.send_json({"ok": False, "error": traceback.format_exc()}, status=500)
            return
        if parsed.path == "/api/cleaning_preview_section":
            try:
                qs = parse_qs(parsed.query)
                show_all = qs.get("show_all", ["0"])[0] == "1"
                items = ensure_crop_numbers(state)
                html, _, button_html, kept_count = render_cleaning_preview_section_parts(state, items, show_all=show_all)
                self.send_json({
                    "ok": True,
                    "html": html,
                    "button_html": button_html,
                    "kept": kept_count,
                    "show_all": show_all,
                })
            except Exception:
                self.send_json({"ok": False, "error": traceback.format_exc()}, status=500)
            return
        if parsed.path == "/api/analysis_progress":
            self.send_json(get_analysis_progress())
            return
        if parsed.path == "/api/count_size_threshold_crops":
            try:
                items = ensure_crop_numbers(state)
                self.send_json({
                    "ok": True,
                    "crops": count_size_threshold_crop_rows(state, items),
                })
            except Exception:
                self.send_json({"ok": False, "error": traceback.format_exc()}, status=500)
            return
        if parsed.path == "/api/recovered_crops_section":
            try:
                items = ensure_crop_numbers(state)
                self.send_json({
                    "ok": True,
                    "html": render_recovered_crop_panel(state, items),
                })
            except Exception:
                self.send_json({"ok": False, "error": traceback.format_exc()}, status=500)
            return
        if parsed.path == "/thumb":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            item = state.get("items", {}).get(rel)
            if not item:
                self.send_error(404)
                return
            try:
                path = item["path"]
                data = png_bytes_for_tiff(path)
                self.send_png(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/recovered_crop_thumb":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            kind = qs.get("kind", ["selected"])[0]
            item = state.get("items", {}).get(rel)
            if not item:
                self.send_error(404)
                return
            try:
                selected_path, _ = selected_crop_path_for_item(state, rel, item)
                path = selected_path
                if kind == "auto_cleaned":
                    cached = current_auto_cleaned_input_for_item_settings(state, rel, item, selected_path)
                    if not cached:
                        self.send_error(404)
                        return
                    path = cached[0]
                data = png_bytes_for_tiff(path)
                self.send_png(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/index_crop_thumb":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            kind = qs.get("kind", ["raw"])[0]
            item = state.get("items", {}).get(rel)
            if not item:
                self.send_error(404)
                return
            context = crop_context_for_rel(state, rel)
            path = context.get("enhanced_crop") if kind == "enhanced" else context.get("raw_crop")
            if kind != "enhanced" and not path:
                path, _ = selected_crop_path_for_item(state, rel, item)
            try:
                data = png_bytes_for_index_path(path)
                self.send_png(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/index_context_thumb":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            kind = qs.get("kind", ["raw_source"])[0]
            item = state.get("items", {}).get(rel)
            if not item:
                self.send_error(404)
                return
            context = crop_context_for_rel(state, rel)
            path = context.get("enhanced_source") if kind == "enhanced_source" else context.get("raw_source")
            try:
                bbox = context.get("bbox", {})
                data = png_bytes_for_context_path(path, bbox)
                self.send_png(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/raw_m_context_thumb":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            item = state.get("items", {}).get(rel)
            if not item:
                self.send_error(404)
                return
            context = crop_context_for_rel(state, rel)
            try:
                path = find_raw_m_image(state, rel)
                bbox = context.get("bbox", {})
                data = png_bytes_for_context_path(path, bbox)
                self.send_png(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/raw_m_thumb":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            item = state.get("items", {}).get(rel)
            if not item:
                self.send_error(404)
                return
            try:
                path = find_raw_m_image(state, rel)
                data = png_bytes_for_tiff(path)
                self.send_png(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/manual_crop_png":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            source_mode = qs.get("source", ["crop"])[0]
            channel_name = qs.get("channel", ["merged"])[0]
            item = state.get("items", {}).get(rel)
            if not item:
                self.send_error(404)
                return
            try:
                data = png_bytes_for_manual_crop_image(state, rel, source_mode=source_mode, channel_name=channel_name)
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/api/fixed_cell_annotation":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            item = state.get("items", {}).get(rel)
            if not isinstance(item, dict):
                self.send_json({"ok": False, "error": "Crop is not loaded."}, status=404)
                return
            try:
                crop_box = reviewed_crop_bounds_on_full_image(state, rel, item)
                untouched_path = untouched_acquisition_path_for_item(state, rel, item)
                untouched_yxc, _ = read_yxc(untouched_path)
                crop_box = reviewed_crop_bounds_on_full_image(state, rel, item, untouched_yxc.shape)
            except Exception as error:
                self.send_json({"ok": False, "error": f"Could not map this reviewed crop to an untouched acquisition: {error}"}, status=400)
                return
            annotation = item.get("fixed_cell_manual_annotation") or {}
            default_source = (
                "untouched"
                if normalize_fixed_cell_intensity_source_mode(state.get("fixed_cell_intensity_source_mode")) == "untouched_acquisition"
                else "crop"
            )
            self.send_json({
                "ok": True, "rel": rel, "annotation": annotation,
                "default_source_mode": default_source,
                "reviewed_crop_box": crop_box,
                "untouched_path": untouched_path,
            })
            return
        if parsed.path == "/api/partition_annotation":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            item = state.get("items", {}).get(rel)
            if not isinstance(item, dict):
                self.send_json({"ok": False, "error": "Crop is not loaded."}, status=404)
                return
            annotation = item.get("partition_manual_annotation")
            seeded = False
            if not isinstance(annotation, dict):
                existing = item.get("fixed_cell_manual_annotation")
                if isinstance(existing, dict) and existing.get("roi_points") and existing.get("condensate_polygons"):
                    annotation = {
                        "source_mode": existing.get("source_mode", "original"),
                        "roi_points": existing.get("roi_points", []),
                        "condensate_polygons": existing.get("condensate_polygons", []),
                        "background_polygons": [],
                        "no_condensate": False,
                    }
                    seeded = True
                else:
                    annotation = {"source_mode": "untouched", "roi_points": [], "condensate_polygons": [], "background_polygons": [], "no_condensate": False}
            try:
                untouched_path = untouched_acquisition_path_for_item(state, rel, item)
                untouched_yxc, _ = read_yxc(untouched_path)
                crop_box = reviewed_crop_bounds_on_full_image(state, rel, item, untouched_yxc.shape)
            except Exception as error:
                self.send_json({"ok": False, "error": f"Could not map this reviewed crop to an untouched acquisition: {error}"}, status=400)
                return
            self.send_json({
                "ok": True,
                "rel": rel,
                "annotation": annotation,
                "seeded_from_fixed_cell": seeded,
                "default_source_mode": "untouched",
                "reviewed_crop_box": crop_box,
                "untouched_path": untouched_path,
            })
            return
        if parsed.path == "/choice_thumb":
            qs = parse_qs(parsed.query)
            kind = qs.get("kind", [""])[0]
            rel = unquote(qs.get("rel", [""])[0])
            raw_crop_root, enhanced_crop_root = raw_and_enhanced_crop_roots(state)
            root = enhanced_crop_root if kind == "enhanced" else raw_crop_root
            path = os.path.abspath(os.path.join(root, rel))
            root_abs = os.path.abspath(root)
            if not path.startswith(root_abs) or not os.path.exists(path):
                self.send_error(404)
                return
            try:
                data = png_bytes_for_tiff(path)
                self.send_png(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/final_preview_png":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            final_items, final_preview_root = final_preview_png_items(state)
            path = os.path.abspath(os.path.join(final_preview_root, rel))
            root_abs = os.path.abspath(final_preview_root)
            if not path.startswith(root_abs) or not os.path.exists(path):
                self.send_error(404)
                return
            try:
                with open(path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/fixed_cell_preview_png":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            _, preview_root = fixed_cell_preview_png_items(state)
            path = os.path.abspath(os.path.join(preview_root, rel))
            root_abs = os.path.abspath(preview_root)
            if not preview_root or os.path.commonpath([path, root_abs]) != root_abs or not os.path.isfile(path):
                self.send_error(404)
                return
            try:
                with open(path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/channel_cleaning_preview_png":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            cleaning_items, cleaning_preview_root = channel_cleaning_preview_items(state)
            path = os.path.abspath(os.path.join(cleaning_preview_root, rel))
            root_abs = os.path.abspath(cleaning_preview_root)
            if not path.startswith(root_abs) or not os.path.exists(path):
                self.send_error(404)
                return
            try:
                with open(path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/approved_thumb":
            qs = parse_qs(parsed.query)
            rel = unquote(qs.get("rel", [""])[0])
            approved_items, approved_root = approved_crop_items(state)
            path = os.path.abspath(os.path.join(approved_root, rel))
            approved_root_abs = os.path.abspath(approved_root)
            if not path.startswith(approved_root_abs) or not os.path.exists(path):
                self.send_error(404)
                return
            try:
                data = png_bytes_for_tiff(path)
                self.send_png(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        if parsed.path == "/count_size_threshold_default":
            try:
                qs = parse_qs(parsed.query)
                rel = unquote(qs.get("rel", [""])[0])
                channel = qs.get("channel", ["red"])[0]
                if "analysis_input_mode" in qs:
                    mode = normalize_analysis_input_mode(qs.get("analysis_input_mode", ["raw"])[0])
                    state["analysis_input_mode"] = mode
                    state["pcc_analysis_input_mode"] = mode
                    state["count_size_analysis_input_mode"] = mode
                self.send_json(count_size_threshold_default_json(state, rel, channel))
            except Exception as error:
                self.send_json({"error": str(error), "traceback": traceback.format_exc()}, status=500)
            return
        if parsed.path == "/count_size_threshold_preview":
            try:
                qs = parse_qs(parsed.query)
                rel = unquote(qs.get("rel", [""])[0])
                channel = qs.get("channel", ["red"])[0]
                threshold_min = qs.get("threshold_min", ["0"])[0]
                threshold_max = qs.get("threshold_max", ["255"])[0]
                view = qs.get("view", ["watershed"])[0]
                fast_watershed = qs.get("fast_watershed", ["0"])[0] == "1"
                if "analysis_input_mode" in qs:
                    mode = normalize_analysis_input_mode(qs.get("analysis_input_mode", ["raw"])[0])
                    state["analysis_input_mode"] = mode
                    state["pcc_analysis_input_mode"] = mode
                    state["count_size_analysis_input_mode"] = mode
                data = count_size_threshold_preview_png(
                    state, rel, channel, threshold_min, threshold_max, view,
                    fast_watershed=fast_watershed,
                )
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except Exception:
                self.send_error(500, traceback.format_exc())
            return
        self.send_html(page(state))

    def read_post(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length).decode("utf-8")
        return {k: v[0] for k, v in parse_qs(raw).items()}

    def do_POST(self):
        state = load_state()
        try:
            if self.path == "/choose_local_raw_folder":
                # Older cached pages posted here to launch an AppKit/Tk dialog
                # from a background Python process. That can crash on macOS.
                # Keep this route safe and direct the user to browser upload.
                msg = (
                    "The old desktop folder picker has been disabled. "
                    "Choose the folder with the browser control below, or paste its full Finder path."
                )
                self.send_html(page(state, msg))
                return
            if self.path == "/use_local_raw_folder_path":
                data = self.read_post()
                raw_image_root = data.get("raw_image_folder_path", "")
                msg = use_local_raw_image_folder(state, raw_image_root)
                self.send_html(page(state, msg))
                return
            if self.path == "/upload_raw_folder":
                length = int(self.headers.get("Content-Length", 0))
                if length <= 0:
                    self.send_html(page(state, "The browser did not provide an upload size. Use the original-location path box instead."), status=400)
                    return
                if length > MAX_BROWSER_UPLOAD_BYTES:
                    self.close_connection = True
                    size_mb = length / (1024 * 1024)
                    msg = (
                        f"This folder is approximately {size_mb:.1f} MB and is too large for browser copy. "
                        "Nothing was uploaded and the previous active experiment was not changed. "
                        "Paste the folder's original Finder/Windows path and click the original-location button instead."
                    )
                    self.send_html(page(state, msg), status=413)
                    return
                try:
                    body = self.rfile.read(length)
                except OSError:
                    self.close_connection = True
                    msg = (
                        "The browser could not transfer this folder safely. Nothing was uploaded and the previous active experiment was not changed. "
                        "Use the original-location path box for this experiment."
                    )
                    self.send_html(page(state, msg), status=400)
                    return
                upload_dir, file_count = save_uploaded_raw_folder(self.headers, body)
                raw_image_root = choose_uploaded_raw_image_root(upload_dir)
                tiff_count = count_tiff_files(raw_image_root)
                state["raw_image_folder"] = raw_image_root
                refresh_derived_paths(state)
                state["items"] = {}
                state["last_scan"] = ""
                state["last_stage_output"] = ""
                state["last_stage_returncode"] = None
                state["final_preview_root"] = ""
                state["last_final_excel"] = ""
                save_state(state)
                if os.path.isdir(state.get("crop_root", "")):
                    scan_crops(state)
                    summary = reload_existing_results_for_current_folder(state)
                    if summary.get("status") == "recovered":
                        msg = f"Uploaded {file_count} files and found {tiff_count} TIFF images. Reloaded previous work: kept {summary.get('kept')}, rejected {summary.get('rejected')}, manual recrops {summary.get('manual_recrops')}."
                    else:
                        msg = f"Uploaded {file_count} files and found {tiff_count} TIFF images. Existing crops were loaded, but no previous decisions were found."
                else:
                    msg = f"Uploaded {file_count} files and found {tiff_count} TIFF images. All results will be written inside this uploaded raw image folder: {raw_image_root}"
                self.send_html(page(state, msg))
                return
            if self.path == "/pipeline_config":
                data = self.read_post()
                state["preprocessor_folder_name"] = data.get("preprocessor_folder_name", "ENHANCED_JACOP_PREPROCESSOR_RESULTS").strip()
                state["crop_folder_name"] = data.get("crop_folder_name", "03_cropped_rois_for_final_choice").strip()
                state["final_output_folder_name"] = data.get("final_output_folder_name", "REAL FIJI PCC MANDERS RESULTS").strip()
                state["final_input_choice"] = "raw_masked"
                state["analysis_mode"] = valid_analysis_mode(data.get("analysis_mode"), state.get("analysis_mode", "red_green"))
                state["stage3_mode"] = "dapi_red_green" if state["analysis_mode"] == "red_green" else "no_dapi_three_color"
                state["pixel_um"] = float(data.get("pixel_um", state.get("pixel_um", 0.1290271)))
                state["pcc_red_label"] = data.get("pcc_red_label", state.get("pcc_red_label", "")).strip()
                state["pcc_green_label"] = data.get("pcc_green_label", state.get("pcc_green_label", "")).strip()
                state["pcc_blue_label"] = data.get("pcc_blue_label", state.get("pcc_blue_label", "")).strip()
                state["red_channel_role"] = valid_channel_role(data.get("red_channel_role", state.get("red_channel_role", "diffuse_marker")), "diffuse_marker")
                state["green_channel_role"] = valid_channel_role(data.get("green_channel_role", state.get("green_channel_role", "condensate")), "condensate")
                state["blue_channel_role"] = valid_channel_role(data.get("blue_channel_role", state.get("blue_channel_role", "nucleus")), "nucleus")
                state["pcc_pair_mode"] = valid_pcc_pair_mode(data.get("pcc_pair_mode", state.get("pcc_pair_mode", "red_green")), "red_green")
                state["analysis_input_mode"] = normalize_analysis_input_mode(data.get("analysis_input_mode", state.get("analysis_input_mode")))
                state["pcc_analysis_input_mode"] = normalize_analysis_input_mode(data.get("pcc_analysis_input_mode", state.get("pcc_analysis_input_mode", state["analysis_input_mode"])))
                state["count_size_analysis_input_mode"] = normalize_analysis_input_mode(data.get("count_size_analysis_input_mode", state.get("count_size_analysis_input_mode", state["analysis_input_mode"])))
                refresh_derived_paths(state)
                save_state(state)
                self.send_html(page(state, "Pipeline settings saved."))
                return
            if self.path == "/run_stage":
                data = self.read_post()
                stage = data.get("stage", "")
                try:
                    output, _ = run_original_script_stage(state, stage)
                except MissingAnalysisScriptError as error:
                    state["last_workflow_action"] = f"run_stage_{stage}"
                    state["last_workflow_note"] = str(error)
                    save_state(state)
                    self.send_html(page(state, str(error)))
                    return
                msg = "Stage complete: " + stage
                if stage == "crop":
                    scan_crops(state)
                    summary = reload_existing_results_for_current_folder(state)
                    count = len(state.get("items", {}))
                    if summary.get("status") == "recovered":
                        msg += f". Loaded {count} crops and reloaded previous work: kept {summary.get('kept')}, rejected {summary.get('rejected')}, manual recrops {summary.get('manual_recrops')}. Review/manual recrop first, then click Autoclean kept crops."
                    else:
                        msg += f". Loaded {count} crops for review. Autoclean was not run; review/manual recrop first, then click Autoclean kept crops."
                self.send_html(page(state, msg))
                return
            if self.path == "/choose_crop_set":
                data = self.read_post()
                choice = "raw_masked"
                choose_crop_set_for_review(state, choice)
                summary = reload_existing_results_for_current_folder(state)
                count = len(state.get("items", {}))
                label = "raw masked"
                if summary.get("status") == "recovered":
                    msg = f"Using {label} crops. Loaded {count} crops and reloaded previous work: kept {summary.get('kept')}, rejected {summary.get('rejected')}, manual recrops {summary.get('manual_recrops')}."
                else:
                    msg = f"Using {label} crops. Loaded {count} crops for review/final analysis."
                self.send_html(page(state, msg))
                return
            if self.path == "/configure_names":
                data = self.read_post()
                state["crop_folder_name"] = data.get("crop_folder_name", state.get("crop_folder_name", "03_cropped_rois_for_final_choice")).strip()
                state["final_output_folder_name"] = data.get("final_output_folder_name", state.get("final_output_folder_name", "REAL FIJI PCC MANDERS RESULTS")).strip()
                state["final_input_choice"] = "raw_masked"
                state["analysis_mode"] = valid_analysis_mode(data.get("analysis_mode"), state.get("analysis_mode", "red_green"))
                state["pcc_red_label"] = data.get("pcc_red_label", state.get("pcc_red_label", "")).strip()
                state["pcc_green_label"] = data.get("pcc_green_label", state.get("pcc_green_label", "")).strip()
                state["pcc_blue_label"] = data.get("pcc_blue_label", state.get("pcc_blue_label", "")).strip()
                state["red_channel_role"] = valid_channel_role(data.get("red_channel_role", state.get("red_channel_role", "diffuse_marker")), "diffuse_marker")
                state["green_channel_role"] = valid_channel_role(data.get("green_channel_role", state.get("green_channel_role", "condensate")), "condensate")
                state["blue_channel_role"] = valid_channel_role(data.get("blue_channel_role", state.get("blue_channel_role", "nucleus")), "nucleus")
                state["pcc_pair_mode"] = valid_pcc_pair_mode(data.get("pcc_pair_mode", state.get("pcc_pair_mode", "red_green")), "red_green")
                state["analysis_input_mode"] = normalize_analysis_input_mode(data.get("analysis_input_mode", state.get("analysis_input_mode")))
                state["pcc_analysis_input_mode"] = normalize_analysis_input_mode(data.get("pcc_analysis_input_mode", state.get("pcc_analysis_input_mode", state["analysis_input_mode"])))
                state["count_size_analysis_input_mode"] = normalize_analysis_input_mode(data.get("count_size_analysis_input_mode", state.get("count_size_analysis_input_mode", state["analysis_input_mode"])))
                # Keep the selected stage mode roughly in sync with the review-analysis mode.
                state["stage3_mode"] = "dapi_red_green" if state["analysis_mode"] == "red_green" else "no_dapi_three_color"
                refresh_derived_paths(state)
                save_state(state)
                scan_crops(state)
                summary = reload_existing_results_for_current_folder(state)
                count = len(state.get("items", {}))
                if summary.get("status") == "recovered":
                    msg = f"Using folder names. Loaded {count} crops and reloaded previous work: kept {summary.get('kept')}, rejected {summary.get('rejected')}, manual recrops {summary.get('manual_recrops')}."
                else:
                    msg = f"Using folder names. Scanned {count} TIFF crop images."
                self.send_html(page(state, msg))
                return
            if self.path == "/configure":
                data = self.read_post()
                state["crop_root"] = data.get("crop_root", state["crop_root"]).strip()
                state["output_root"] = data.get("output_root", state["output_root"]).strip()
                state["analysis_mode"] = valid_analysis_mode(data.get("analysis_mode"), state.get("analysis_mode", "red_green"))
                apply_label_post_fields(state, data)
                scan_crops(state)
                summary = reload_existing_results_for_current_folder(state)
                count = len(state.get("items", {}))
                if summary.get("status") == "recovered":
                    msg = f"Scanned {count} crops and reloaded previous work: kept {summary.get('kept')}, rejected {summary.get('rejected')}, manual recrops {summary.get('manual_recrops')}."
                else:
                    msg = f"Scanned {count} TIFF crop images."
                self.send_html(page(state, msg))
                return
            if self.path == "/recover_review_state":
                try:
                    count, kept, rejected, manual, sources = recover_review_state_from_existing_outputs(state)
                except (ValueError, FileNotFoundError) as error:
                    self.send_html(page(state, str(error)))
                    return
                source_count = len(sources)
                self.send_html(page(state, f"Recovered review state from {source_count} previous result source(s). Loaded crops: {count}. Kept: {kept}. Rejected: {rejected}. Manual recrops restored: {manual}."))
                return
            if self.path == "/load_previous_job_folder":
                data = self.read_post()
                folder = data.get("previous_job_folder", "").strip()
                try:
                    summary = load_previous_job_folder(state, folder)
                except Exception as error:
                    self.send_html(page(state, str(error)))
                    return
                stage_summary = summary.get("stage_summary", {})
                recovered_outputs = []
                if stage_summary.get("pcc_manders_excel_exists"):
                    recovered_outputs.append("PCC/Manders")
                if stage_summary.get("count_size_excel_exists"):
                    recovered_outputs.append("count/size")
                if stage_summary.get("fixed_cell_ms2_excel_exists"):
                    recovered_outputs.append("fixed-cell MS2")
                if stage_summary.get("partition_excel_exists"):
                    recovered_outputs.append("partition coefficient")
                if summary.get("status") == "recovered":
                    msg = (
                        f"Loaded previous job folder. Crops: {summary.get('loaded_crops', 0)}. "
                        f"Kept: {summary.get('kept', 0)}. Rejected: {summary.get('rejected', 0)}. "
                        f"Manual recrops restored: {summary.get('manual_recrops', 0)}. "
                        f"Saved count/size thresholds: {stage_summary.get('saved_count_size_thresholds', 0)}. "
                        f"Recovered outputs: {', '.join(recovered_outputs) if recovered_outputs else 'output links not found'}. "
                        f"Sources recovered: {len(summary.get('sources', []))}. "
                        "No crop, cleaning, Fiji, PCC/Manders, or count/size step was run; the app is waiting for the next button you click."
                    )
                elif summary.get("status") == "preprocessor_ready":
                    msg = (
                        f"Loaded previous job folder. Preprocessor output is ready: "
                        f"{summary.get('preprocessor_ready_tiffs', 0)} raw-masked TIFF(s) found. "
                        "No reviewed ROI crops exist yet. No workflow step was run; click the Crop button when you want to create the review crop folder."
                    )
                else:
                    msg = (
                        f"Loaded previous job folder and scanned {summary.get('loaded_crops', 0)} crop(s), "
                        f"but previous decisions were not found. {summary.get('error', '')} "
                        "No crop, cleaning, Fiji, PCC/Manders, or count/size step was run."
                    )
                self.send_html(page(state, msg))
                return
            if self.path == "/choose_previous_job_folder":
                try:
                    folder = choose_local_previous_job_folder_with_dialog()
                    summary = load_previous_job_folder(state, folder)
                except Exception as error:
                    self.send_html(page(state, str(error)))
                    return
                stage_summary = summary.get("stage_summary", {})
                recovered_outputs = []
                if stage_summary.get("pcc_manders_excel_exists"):
                    recovered_outputs.append("PCC/Manders")
                if stage_summary.get("count_size_excel_exists"):
                    recovered_outputs.append("count/size")
                if stage_summary.get("fixed_cell_ms2_excel_exists"):
                    recovered_outputs.append("fixed-cell MS2")
                if stage_summary.get("partition_excel_exists"):
                    recovered_outputs.append("partition coefficient")
                if summary.get("status") == "recovered":
                    msg = (
                        f"Loaded previous job folder from {folder}. "
                        f"Raw TIFFs: {stage_summary.get('raw_tiff_count', 0)}. "
                        f"Crops: {summary.get('loaded_crops', 0)}. "
                        f"Kept: {summary.get('kept', 0)}. Rejected: {summary.get('rejected', 0)}. "
                        f"Manual recrops restored: {summary.get('manual_recrops', 0)}. "
                        f"Saved count/size thresholds: {stage_summary.get('saved_count_size_thresholds', 0)}. "
                        f"Recovered outputs: {', '.join(recovered_outputs) if recovered_outputs else 'output links not found'}. "
                        "No crop, cleaning, Fiji, PCC/Manders, or count/size step was run; the app is waiting for the next button you click."
                    )
                elif summary.get("status") == "preprocessor_ready":
                    msg = (
                        f"Loaded previous job folder from {folder}. Preprocessor output is ready: "
                        f"{summary.get('preprocessor_ready_tiffs', 0)} raw-masked TIFF(s) found. "
                        "No reviewed ROI crops exist yet. No workflow step was run; click the Crop button when you want to create the review crop folder."
                    )
                else:
                    msg = (
                        f"Loaded previous job folder from {folder} and scanned {summary.get('loaded_crops', 0)} crop(s), "
                        f"but previous decisions were not found. {summary.get('error', '')} "
                        "No crop, cleaning, Fiji, PCC/Manders, or count/size step was run."
                    )
                self.send_html(page(state, msg))
                return
            if self.path == "/set_crop_choice":
                data = self.read_post()
                rel = data.get("rel", "")
                crop_choice = "raw"
                if rel in state.get("items", {}):
                    state["items"][rel]["crop_choice"] = crop_choice
                    save_state(state)
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return
            if self.path == "/api/reject_other_crops_from_image":
                data = self.read_post()
                rel = data.get("rel", "")
                items = state.get("items", {})
                if rel not in items:
                    self.send_json({"ok": False, "error": "Crop not found."}, status=404)
                    return
                source_key = crop_source_key(rel)
                items[rel]["status"] = "keep"
                items[rel]["explicit_keep"] = True
                rejected_rels = []
                for other_rel, other_item in items.items():
                    if crop_source_key(other_rel) != source_key:
                        continue
                    if other_rel == rel or other_item.get("explicit_keep") or other_item.get("manual_crop_path"):
                        continue
                    if other_item.get("status", "keep") != "reject":
                        rejected_rels.append(other_rel)
                    other_item["status"] = "reject"
                    other_item["explicit_keep"] = False
                save_state(state)
                kept_now = sum(1 for x in items.values() if x.get("status", "keep") == "keep")
                rejected_now = sum(1 for x in items.values() if x.get("status", "keep") == "reject")
                self.send_json({
                    "ok": True,
                    "current_rel": rel,
                    "source_key": source_key,
                    "rejected_rels": rejected_rels,
                    "total": len(items),
                    "kept": kept_now,
                    "rejected": rejected_now,
                })
                return
            if self.path == "/api/set_status":
                data = self.read_post()
                rel = data.get("rel", "")
                status = data.get("status", "keep")
                if status not in ("keep", "reject"):
                    status = "keep"
                item = state.get("items", {}).get(rel)
                if rel in state.get("items", {}):
                    item = state["items"][rel]
                    item["status"] = status
                    item["explicit_keep"] = status == "keep"
                    save_state(state)
                kept_now = sum(1 for x in state.get("items", {}).values() if x.get("status", "keep") == "keep")
                rejected_now = sum(1 for x in state.get("items", {}).values() if x.get("status", "keep") == "reject")
                self.send_json({
                    "ok": True,
                    "rel": rel,
                    "status": status,
                    "explicit_keep": bool(item.get("explicit_keep")) if item else False,
                    "total": len(state.get("items", {})),
                    "kept": kept_now,
                    "rejected": rejected_now,
                })
                return
            if self.path == "/set_status":
                data = self.read_post()
                rel = data.get("rel", "")
                status = data.get("status", "keep")
                if status not in ("keep", "reject"):
                    status = "keep"
                if rel in state.get("items", {}):
                    state["items"][rel]["status"] = status
                    state["items"][rel]["explicit_keep"] = status == "keep"
                    save_state(state)
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return
            if self.path == "/set_crop_channel_roles":
                data = self.read_post()
                rel = data.get("rel", "")
                item = state.get("items", {}).get(rel)
                if item is None:
                    self.send_html(page(state, f"Crop not found for role update: {rel}"), status=404)
                    return
                roles = set_item_channel_roles(item, data)
                save_state(state)
                role_summary = ", ".join(f"{channel.title()}: {role_label(role)}" for channel, role in roles.items())
                self.send_html(page(state, f"Saved per-crop channel roles for crop: {rel}. {role_summary}"))
                return
            if self.path == "/select_crop_numbers":
                data = self.read_post()
                crop_start = data.get("crop_start", "").strip()
                crop_end = data.get("crop_end", "").strip()
                if not crop_start and not crop_end and not data.get("crop_numbers", ""):
                    crop_start, crop_end = "1", str(len(state.get("items", {})))
                crop_numbers = f"{crop_start}-{crop_end}" if crop_start and crop_end else data.get("crop_numbers", "")
                kept_now, rejected_now, _ = apply_crop_number_selection(state, crop_numbers)
                self.send_html(page(state, f"Crop number selection applied. Kept: {kept_now}. Rejected: {rejected_now}."))
                return
            if self.path == "/api/select_crop_numbers":
                data = self.read_post()
                crop_start = data.get("crop_start", "").strip()
                crop_end = data.get("crop_end", "").strip()
                if not crop_start and not crop_end and not data.get("crop_numbers", ""):
                    crop_start, crop_end = "1", str(len(state.get("items", {})))
                crop_numbers = f"{crop_start}-{crop_end}" if crop_start and crop_end else data.get("crop_numbers", "")
                try:
                    kept_now, rejected_now, selected_numbers = apply_crop_number_selection(state, crop_numbers)
                    items = ensure_crop_numbers(state)
                    crop_statuses = [
                        {"rel": item.get("rel_path", ""), "status": item.get("status", "keep")}
                        for item in items
                    ]
                    cleaning_html, _, cleaning_button_html, _ = render_cleaning_preview_section_parts(state, items, show_all=False)
                    payload = json.dumps({
                        "ok": True,
                        "total": len(items),
                        "kept": kept_now,
                        "rejected": rejected_now,
                        "selected_numbers": selected_numbers,
                        "crop_statuses": crop_statuses,
                        "recovered_html": render_recovered_crop_panel(state, items),
                        "cleaning_html": cleaning_html,
                        "cleaning_button_html": cleaning_button_html,
                    }).encode("utf-8")
                    self.send_response(200)
                except ValueError as error:
                    payload = json.dumps({"ok": False, "error": str(error)}).encode("utf-8")
                    self.send_response(400)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if self.path == "/manual_recrop":
                data = self.read_post()
                rel = data.get("rel", "")
                try:
                    manual_rel, manual_path = manual_recrop_item(
                        state,
                        rel,
                        data.get("x", 0),
                        data.get("y", 0),
                        data.get("width", 0),
                        data.get("height", 0),
                        data.get("source", "crop"),
                    )
                    self.send_json({"ok": True, "rel": manual_rel, "path": manual_path, "thumb_url": f"/thumb?rel={quote(manual_rel)}&v={int(os.path.getmtime(manual_path))}"})
                except ValueError as error:
                    self.send_json({"ok": False, "error": str(error)}, status=400)
                except (PermissionError, OSError) as error:
                    self.send_json({
                        "ok": False,
                        "error": (
                            "The app cannot write to this experiment folder. Close the app and reopen "
                            "Launch Condensate Analysis.app directly from Finder. If macOS asks, allow access "
                            "to files and removable volumes. No existing image was changed. "
                            f"System message: {error}"
                        ),
                    }, status=403)
                return
            if self.path == "/save_fixed_cell_annotation":
                data = self.read_post()
                rel = data.get("rel", "")
                try:
                    roi_points = json.loads(data.get("roi_points", "[]"))
                    focus_polygons = json.loads(data.get("focus_polygons", "[]"))
                    condensate_polygons = json.loads(data.get("condensate_polygons", "[]"))
                    background_polygon = json.loads(data.get("background_polygon", "[]"))
                    annotation = save_fixed_cell_manual_annotation(
                        state,
                        rel,
                        roi_points,
                        focus_polygons,
                        source_mode=data.get("source_mode", "crop"),
                        background_polygon=background_polygon,
                        condensate_polygons=condensate_polygons,
                    )
                    self.send_json({"ok": True, "rel": rel, "annotation": annotation})
                except (ValueError, json.JSONDecodeError) as error:
                    self.send_json({"ok": False, "error": str(error)}, status=400)
                except (PermissionError, OSError) as error:
                    self.send_json({"ok": False, "error": f"Could not save the manual annotation: {error}"}, status=403)
                return
            if self.path == "/save_partition_annotation":
                data = self.read_post()
                rel = data.get("rel", "")
                try:
                    roi_points = json.loads(data.get("roi_points", "[]"))
                    condensate_polygons = json.loads(data.get("condensate_polygons", "[]"))
                    background_polygons = json.loads(data.get("background_polygons", "[]"))
                    annotation = save_partition_manual_annotation(
                        state,
                        rel,
                        roi_points,
                        condensate_polygons,
                        source_mode=data.get("source_mode", "original"),
                        background_polygons=background_polygons,
                        no_condensate=str(data.get("no_condensate", "0")).strip().lower() in ("1", "true", "yes", "on"),
                    )
                    self.send_json({"ok": True, "rel": rel, "annotation": annotation})
                except (ValueError, json.JSONDecodeError) as error:
                    self.send_json({"ok": False, "error": str(error)}, status=400)
                except (PermissionError, OSError) as error:
                    self.send_json({"ok": False, "error": f"Could not save the partition annotation: {error}"}, status=403)
                return
            if self.path == "/undo_manual_recrop":
                data = self.read_post()
                rel = data.get("rel", "")
                try:
                    undo_rel, original_path = undo_manual_recrop(state, rel)
                    self.send_json({"ok": True, "rel": undo_rel, "path": original_path, "thumb_url": f"/thumb?rel={quote(undo_rel)}&v={int(os.path.getmtime(original_path))}"})
                except ValueError as error:
                    self.send_json({"ok": False, "error": str(error)}, status=400)
                except (PermissionError, OSError) as error:
                    self.send_json({
                        "ok": False,
                        "error": (
                            "The app cannot update this experiment folder. Reopen Launch Condensate Analysis.app "
                            "directly from Finder and allow file/removable-volume access if macOS asks. "
                            f"System message: {error}"
                        ),
                    }, status=403)
                return
            if self.path == "/all_keep":
                for item in state.get("items", {}).values():
                    item["status"] = "keep"
                    item["explicit_keep"] = True
                save_state(state)
                self.send_html(page(state, "All crops set to keep."))
                return
            if self.path == "/preview_channel_cleaning":
                data = self.read_post()
                apply_label_post_fields(state, data)
                count = preview_channel_cleaning_on_kept_crops(state, force_global=True)
                params = cleaning_params_from_state(state)
                msg = (
                    f"Autoclean complete for {count} kept crop(s). "
                    f"Radius: {params['radius_px']} px. Sensitivity: {params['sensitivity']}. "
                    f"Minimum puncta size: {params['min_area_px']} px. Raw-vs-cleaned previews and cleaned analysis inputs are ready for the final analysis buttons."
                )
                self.send_html(page(state, msg))
                return
            if self.path == "/api/preview_global_channel_cleaning":
                data = self.read_post()
                try:
                    apply_label_post_fields(state, data)
                    count = preview_channel_cleaning_on_kept_crops(state, force_global=True)
                    items = ensure_crop_numbers(state)
                    html, _, button_html, kept_count = render_cleaning_preview_section_parts(state, items, show_all=True)
                    self.send_json({
                        "ok": True,
                        "count": count,
                        "kept": kept_count,
                        "analysis_input_mode": normalize_analysis_input_mode(state.get("analysis_input_mode", "auto_cleaned")),
                        "count_size_analysis_input_mode": normalize_analysis_input_mode(state.get("count_size_analysis_input_mode", "auto_cleaned")),
                        "html": html,
                        "button_html": button_html,
                    })
                except Exception:
                    self.send_json({"ok": False, "error": traceback.format_exc()}, status=500)
                return
            if self.path == "/api/save_ui_settings":
                data = self.read_post()
                try:
                    apply_label_post_fields(state, data)
                    self.send_json({"ok": True})
                except Exception:
                    self.send_json({"ok": False, "error": traceback.format_exc()}, status=500)
                return
            if self.path == "/api/save_one_channel_cleaning_settings":
                data = self.read_post()
                rel = data.get("rel", "")
                try:
                    params = posted_channel_cleaning_params(data, state)
                    cleaned_params, cleaned_channels = persist_channel_cleaning_settings_for_one_crop(
                        state, rel, params
                    )
                    self.send_json({
                        "ok": True,
                        "rel": rel,
                        "params": flatten_cleaning_params_for_ui(cleaned_params, state),
                        "channels": sorted(cleaned_channels),
                    })
                except Exception as error:
                    self.send_json({"ok": False, "error": str(error)}, status=400)
                return
            if self.path == "/api/preview_one_channel_cleaning":
                data = self.read_post()
                rel = data.get("rel", "")
                params = posted_channel_cleaning_params(data, state)
                try:
                    preview_path, cleaned_params = preview_channel_cleaning_for_one_crop(state, rel, params)
                    preview_root = os.path.join(state.get("output_root", ""), "channel_cleaned_analysis_inputs", "raw_vs_cleaned_channel_previews")
                    preview_rel = os.path.relpath(preview_path, preview_root)
                    self.send_json({
                        "ok": True,
                        "rel": rel,
                        "preview_rel": preview_rel,
                        "preview_url": "/channel_cleaning_preview_png?rel=" + quote(preview_rel) + "&v=" + str(os.stat(preview_path).st_mtime_ns),
                        "params": flatten_cleaning_params_for_ui(cleaned_params, state),
                        "channels": sorted(channel_set_from_mapping(params, cleaning_channels_for_rel(state, rel))),
                        "roles": channel_roles_for_item(state, state.get("items", {}).get(rel, {})),
                        "source_label": "PER-CROP ADJUSTED",
                        "role_source_label": "PER-CROP ROLES",
                        "analysis_input_mode": normalize_analysis_input_mode(state.get("analysis_input_mode", "auto_cleaned")),
                        "count_size_analysis_input_mode": normalize_analysis_input_mode(state.get("count_size_analysis_input_mode", "auto_cleaned")),
                    })
                except Exception as error:
                    self.send_json({"ok": False, "error": str(error)}, status=400)
                return
            if self.path == "/api/autoclean_one_crop":
                data = self.read_post()
                rel = data.get("rel", "")
                try:
                    cleaned_path, preview_path = autoclean_one_crop_with_saved_settings(state, rel)
                    self.send_json({
                        "ok": True,
                        "rel": rel,
                        "cleaned_path": cleaned_path,
                        "preview_path": preview_path,
                        "generated_at": state.get("items", {}).get(rel, {}).get("auto_cleaned_generated_at", ""),
                        "analysis_input_mode": normalize_analysis_input_mode(state.get("analysis_input_mode", "auto_cleaned")),
                        "count_size_analysis_input_mode": normalize_analysis_input_mode(state.get("count_size_analysis_input_mode", "auto_cleaned")),
                    })
                except Exception as error:
                    self.send_json({"ok": False, "error": str(error)}, status=400)
                return
            if self.path == "/preview_one_channel_cleaning":
                data = self.read_post()
                rel = data.get("rel", "")
                params = posted_channel_cleaning_params(data, state)
                preview_path, cleaned_params = preview_channel_cleaning_for_one_crop(state, rel, params)
                msg = (
                    f"Only this crop was autocleaned: {rel}. "
                    f"Radius: {cleaned_params['radius_px']} px. Sensitivity: {cleaned_params['sensitivity']}. "
                    f"Minimum puncta size: {cleaned_params['min_area_px']} px. Preview: {preview_path}"
                )
                self.send_html(page(state, msg))
                return
            if self.path == "/api/reset_one_channel_cleaning":
                data = self.read_post()
                rel = data.get("rel", "")
                try:
                    preview_path, cleaned_params, cleaned_channels = reset_channel_cleaning_for_one_crop(state, rel)
                    preview_root = os.path.join(state.get("output_root", ""), "channel_cleaned_analysis_inputs", "raw_vs_cleaned_channel_previews")
                    preview_rel = os.path.relpath(preview_path, preview_root)
                    self.send_json({
                        "ok": True,
                        "rel": rel,
                        "preview_rel": preview_rel,
                        "preview_url": "/channel_cleaning_preview_png?rel=" + quote(preview_rel) + "&v=" + str(os.stat(preview_path).st_mtime_ns),
                        "params": flatten_cleaning_params_for_ui(cleaned_params, state),
                        "channels": sorted(cleaned_channels),
                        "roles": channel_roles_for_item(state, state.get("items", {}).get(rel, {})),
                        "source_label": "GLOBAL SETTINGS",
                        "role_source_label": "GLOBAL ROLES",
                    })
                except Exception as error:
                    self.send_json({"ok": False, "error": str(error)}, status=400)
                return
            if self.path == "/reset_one_channel_cleaning":
                data = self.read_post()
                rel = data.get("rel", "")
                reset_channel_cleaning_for_one_crop(state, rel)
                self.send_html(page(state, f"Per-crop cleaning settings reset to global settings for {rel}."))
                return
            if self.path == "/save_count_size_threshold":
                data = self.read_post()
                try:
                    lower, upper, threshold_file = save_count_size_threshold(
                        state,
                        data.get("rel", ""),
                        data.get("channel", "red"),
                        data.get("threshold_min", "0"),
                        data.get("threshold_max", "255"),
                    )
                    self.send_json({
                        "ok": True,
                        "threshold_min": lower,
                        "threshold_max": upper,
                        "threshold_file": threshold_file,
                    })
                except Exception as error:
                    self.send_json({"ok": False, "error": str(error), "traceback": traceback.format_exc()}, status=400)
                return
            if self.path == "/api/set_analysis_input_mode":
                data = self.read_post()
                mode = normalize_analysis_input_mode(data.get("analysis_input_mode", state.get("analysis_input_mode", "raw")))
                state["analysis_input_mode"] = mode
                state["pcc_analysis_input_mode"] = mode
                state["count_size_analysis_input_mode"] = mode
                save_state(state)
                self.send_json({
                    "ok": True,
                    "analysis_input_mode": mode,
                    "pcc_analysis_input_mode": mode,
                    "count_size_analysis_input_mode": mode,
                })
                return
            if self.path == "/analyze_pcc_manders":
                data = self.read_post()
                apply_label_post_fields(state, data)
                set_analysis_progress(True, "preparing", 0, 0, "Refreshing PCC/Manders inputs.")
                refreshed, refresh_errors = ensure_auto_cleaned_inputs_current_for_analysis(state, ["pcc"])
                output_excel, approved_root, final_preview_root, kept, raw_rows = run_pcc_manders_analysis_on_kept_crops(state)
                set_analysis_progress(False, "complete", kept, kept, f"PCC/Manders complete: {kept} crops, {raw_rows} result rows.")
                refresh_msg = auto_clean_refresh_message(refreshed, refresh_errors)
                msg = f"{refresh_msg} PCC/Manders analysis complete. Kept crops: {kept}. Raw result rows: {raw_rows}. Excel: {output_excel}. Shared approved crops: {approved_root}. PCC/Manders previews: {final_preview_root}"
                self.send_html(page(state, msg))
                return
            if self.path == "/analyze_count_size":
                data = self.read_post()
                apply_label_post_fields(state, data)
                set_analysis_progress(True, "preparing", 0, 0, "Refreshing count/size inputs before the Fiji batch.")
                refreshed, refresh_errors = ensure_auto_cleaned_inputs_current_for_analysis(state, ["count_size"])
                output_excel, approved_root, output_root, kept, rows = run_count_size_analysis_on_kept_crops(state)
                refresh_msg = auto_clean_refresh_message(refreshed, refresh_errors)
                msg = f"{refresh_msg} Condensate count/size analysis complete. Kept crops: {kept}. Result rows: {rows}. Excel: {output_excel}. Approved crops: {approved_root}. Output folder: {output_root}"
                self.send_html(page(state, msg))
                return
            if self.path == "/analyze_pcc_manders_count_size":
                data = self.read_post()
                apply_label_post_fields(state, data)
                set_analysis_progress(True, "pcc_manders", 0, 0, "Running PCC/Manders first; count/size will follow in one Fiji batch.")
                refreshed, refresh_errors = ensure_auto_cleaned_inputs_current_for_analysis(state, ["pcc", "count_size"])
                pcc_excel, pcc_approved_root, pcc_preview_root, pcc_kept, pcc_rows = run_pcc_manders_analysis_on_kept_crops(state)
                count_excel, count_approved_root, count_output_root, count_kept, count_rows = run_count_size_analysis_on_kept_crops(state)
                refresh_msg = auto_clean_refresh_message(refreshed, refresh_errors)
                msg = (
                    f"{refresh_msg} PCC/Manders + count/size complete. PCC/Manders kept crops: {pcc_kept}, rows: {pcc_rows}, Excel: {pcc_excel}. "
                    f"Count/size kept crops: {count_kept}, rows: {count_rows}, Excel: {count_excel}. "
                    f"PCC/Manders previews: {pcc_preview_root}. Count/size output folder: {count_output_root}."
                )
                self.send_html(page(state, msg))
                return
            if self.path == "/analyze_fixed_cell_ms2":
                data = self.read_post()
                apply_label_post_fields(state, data)
                set_analysis_progress(True, "fixed_cell_ms2", 0, 0, "Running fixed-cell MS2/condensate measurements.")
                refreshed, refresh_errors = ensure_auto_cleaned_inputs_current_for_analysis(state, ["count_size"])
                output_excel, fixed_root, processed_rows, error_rows = run_fixed_cell_ms2_n1icd_analysis(state)
                set_analysis_progress(False, "complete", processed_rows, processed_rows, f"Fixed-cell analysis complete: {processed_rows} cells.")
                refresh_msg = auto_clean_refresh_message(refreshed, refresh_errors)
                msg = (
                    f"{refresh_msg} Fixed-cell MS2/condensate analysis complete. "
                    f"Processed cells: {processed_rows}. Errors: {error_rows}. Excel: {output_excel}. Output folder: {fixed_root}."
                )
                self.send_html(page(state, msg))
                return
            if self.path == "/analyze_manual_partition":
                data = self.read_post()
                apply_label_post_fields(state, data)
                set_analysis_progress(True, "manual_partition", 0, 0, "Running manual partition-coefficient measurements.")
                output_excel, partition_root, processed_rows, error_rows = run_manual_nuclear_partition_analysis(state)
                set_analysis_progress(False, "complete", processed_rows, processed_rows, f"Partition analysis complete: {processed_rows} cells.")
                msg = (
                    "Manual measurement-ROI partition-coefficient analysis complete. "
                    f"Processed cells: {processed_rows}. Errors: {error_rows}. "
                    f"Excel: {output_excel}. Output folder: {partition_root}. "
                    "PCC/Manders, count/size, and fixed-cell MS2 workflows were not run or changed."
                )
                self.send_html(page(state, msg))
                return
            self.send_error(404)
        except Exception:
            error_text = traceback.format_exc()
            if get_analysis_progress().get("active"):
                set_analysis_progress(False, "failed", 0, 0, "Analysis failed; the error is shown on the page.")
            self.send_html("<pre>" + html_escape(error_text) + "</pre>", status=500)


class ReusableThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True


def main():
    port = 8765
    if len(sys.argv) > 1:
        port = int(sys.argv[1])
    try:
        server = ReusableThreadingHTTPServer(("127.0.0.1", port), Handler)
    except OSError:
        if port == 0:
            raise
        server = ReusableThreadingHTTPServer(("127.0.0.1", 0), Handler)
    actual_port = server.server_address[1]
    url = f"http://127.0.0.1:{actual_port}"
    try:
        with open(SERVER_PORT_FILE, "w", encoding="utf-8") as port_file:
            port_file.write(str(actual_port))
    except OSError:
        pass
    print(f"Open this in your browser: {url}", flush=True)
    print("Press Ctrl+C to stop the app.")
    if os.environ.get("CONDENSATE_OPEN_BROWSER", "").strip() == "1":
        threading.Timer(0.8, lambda: open_local_browser(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nApp stopped.", flush=True)
    finally:
        server.server_close()
        try:
            with open(SERVER_PORT_FILE, "r", encoding="utf-8") as port_file:
                recorded_port = port_file.read().strip()
            if recorded_port == str(actual_port):
                os.remove(SERVER_PORT_FILE)
        except OSError:
            pass


if __name__ == "__main__":
    main()
