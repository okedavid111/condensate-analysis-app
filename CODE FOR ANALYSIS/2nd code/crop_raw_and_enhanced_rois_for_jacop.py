#!/usr/bin/env python3
"""Recovered Fiji PCC/Manders stage 2 ROI cropper."""

import os
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import numpy as np
import pandas as pd
import tifffile as tiff
from skimage import measure, morphology
from skimage.filters import threshold_otsu


EXPERIMENT_ROOT = ""
PREPROCESSOR_FOLDER = ""
QC_EXCEL = ""
RAW_MASKED_IMAGE_FOLDER = ""
ENHANCED_IMAGE_FOLDER = ""
MASK_FOLDER = ""
OUTPUT_FOLDER = ""
RAW_CROP_FOLDER = ""
ENHANCED_CROP_FOLDER = ""
OUTPUT_EXCEL = ""

CROP_PADDING_PX = 12
MIN_CROP_SIZE_PX = 24
MAX_CROPS_PER_IMAGE = 4
MAX_WORKERS = 4


def make_safe_name(name):
    base = os.path.splitext(os.path.basename(str(name)))[0]
    return "".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in base)


def raw_masked_image_name(image_name):
    return f"{make_safe_name(image_name)}_raw_masked_JACoP_ready.tiff"


def enhanced_image_name(image_name):
    return f"{make_safe_name(image_name)}_enhanced_masked_visual_QC.tiff"


def find_file(folder, name):
    direct = os.path.join(folder, name)
    if os.path.isfile(direct):
        return direct
    target = name.lower()
    for dirpath, _, filenames in os.walk(folder):
        for filename in filenames:
            if filename.lower() == target:
                return os.path.join(dirpath, filename)
    return ""


def find_raw_and_enhanced_images(image_name):
    raw_path = find_file(RAW_MASKED_IMAGE_FOLDER, raw_masked_image_name(image_name))
    return raw_path, ""


def read_qc_rows():
    sheets = pd.read_excel(QC_EXCEL, sheet_name=None)
    for preferred in ("Selected ROIs", "All ROIs", "QC"):
        if preferred in sheets:
            df = sheets[preferred]
            break
    else:
        df = next(iter(sheets.values()))
    if "Keep_For_Cropping" in df.columns:
        keep = df["Keep_For_Cropping"].astype(str).str.lower().isin(("true", "1", "yes", "keep"))
        df = df[keep]
    elif "selected_by_marker_rule" in df.columns:
        keep = df["selected_by_marker_rule"].astype(str).str.lower().isin(("true", "1", "yes", "keep"))
        df = df[keep]
    df = df.fillna("")
    image_col = "Image_Name" if "Image_Name" in df.columns else "Image"
    if image_col in df.columns:
        sort_cols = []
        ascending = []
        if "Selection_Rank" in df.columns:
            sort_cols.append("Selection_Rank")
            ascending.append(True)
        if "Selection_Score" in df.columns:
            sort_cols.append("Selection_Score")
            ascending.append(False)
        if "ROI_Area_Pixels" in df.columns:
            sort_cols.append("ROI_Area_Pixels")
            ascending.append(False)
        if sort_cols:
            df = df.sort_values([image_col] + sort_cols, ascending=[True] + ascending)
        df = df.groupby(image_col, sort=False, dropna=False).head(MAX_CROPS_PER_IMAGE)
    return df


def archive_previous_outputs():
    archive_root = os.path.join(OUTPUT_FOLDER, "_previous_crop_runs")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    moved = []
    for path in (RAW_CROP_FOLDER, OUTPUT_EXCEL):
        if not os.path.exists(path):
            continue
        os.makedirs(archive_root, exist_ok=True)
        target = os.path.join(archive_root, timestamp, os.path.basename(path))
        os.makedirs(os.path.dirname(target), exist_ok=True)
        shutil.move(path, target)
        moved.append(target)
    return moved


def component_labels(image):
    arr = np.asarray(image)
    if arr.ndim == 3:
        marker_channel = np.max(arr[..., :2], axis=2)
    else:
        marker_channel = arr
    vals = marker_channel[marker_channel > 0]
    if vals.size < 8:
        mask = marker_channel > 0
        threshold = np.inf
    else:
        try:
            threshold = threshold_otsu(vals)
        except Exception:
            threshold = np.percentile(vals, 75)
        mask = marker_channel > threshold
    mask = morphology.remove_small_objects(mask, 8)
    return measure.label(mask), threshold


def bbox_from_component(image, preferred_label):
    labels, threshold = component_labels(image)
    regions = measure.regionprops(labels)
    if not regions:
        h, w = image.shape[:2]
        return (0, 0, h, w, 0, 0, 0, 0, threshold)
    chosen = None
    for region in regions:
        if int(region.label) == int(preferred_label):
            chosen = region
            break
    if chosen is None:
        idx = max(0, min(int(preferred_label) - 1, len(regions) - 1))
        chosen = regions[idx]
    minr, minc, maxr, maxc = chosen.bbox
    arr = np.asarray(image)
    roi = labels == chosen.label
    red_signal = int(np.count_nonzero(arr[..., 0][roi])) if arr.ndim == 3 else int(np.count_nonzero(arr[roi]))
    green_signal = int(np.count_nonzero(arr[..., 1][roi])) if arr.ndim == 3 and arr.shape[-1] > 1 else 0
    blue_signal = int(np.count_nonzero(arr[..., 2][roi])) if arr.ndim == 3 and arr.shape[-1] > 2 else 0
    return minr, minc, maxr, maxc, int(chosen.area), red_signal, green_signal, blue_signal, threshold


def clamp_bbox(row, shape):
    h, w = shape[:2]
    minr = int(float(row.get("BBox_Min_Row", 0) or 0)) - CROP_PADDING_PX
    minc = int(float(row.get("BBox_Min_Col", 0) or 0)) - CROP_PADDING_PX
    maxr = int(float(row.get("BBox_Max_Row", h) or h)) + CROP_PADDING_PX
    maxc = int(float(row.get("BBox_Max_Col", w) or w)) + CROP_PADDING_PX
    if maxr - minr < MIN_CROP_SIZE_PX:
        center = (maxr + minr) // 2
        minr = center - MIN_CROP_SIZE_PX // 2
        maxr = minr + MIN_CROP_SIZE_PX
    if maxc - minc < MIN_CROP_SIZE_PX:
        center = (maxc + minc) // 2
        minc = center - MIN_CROP_SIZE_PX // 2
        maxc = minc + MIN_CROP_SIZE_PX
    minr = max(0, min(minr, h - 1))
    minc = max(0, min(minc, w - 1))
    maxr = max(minr + 1, min(maxr, h))
    maxc = max(minc + 1, min(maxc, w))
    return minr, minc, maxr, maxc


def crop_one(row_tuple):
    idx, row = row_tuple
    image_name = str(row.get("Image_Name") or row.get("Image") or "")
    if not image_name:
        image_name = os.path.basename(str(row.get("Source_Path", f"image_{idx}")))
    raw_path = str(row.get("Raw_Masked_Path") or "") if os.path.isfile(str(row.get("Raw_Masked_Path") or "")) else ""
    enhanced_path = ""
    if not raw_path:
        raw_path, enhanced_path = find_raw_and_enhanced_images(image_name)
    if not raw_path:
        raise FileNotFoundError(f"Raw masked image not found for {image_name}")

    raw = tiff.imread(raw_path)
    if "BBox_Min_Row" in row and str(row.get("BBox_Max_Row", "")).strip() != "":
        minr, minc, maxr, maxc = clamp_bbox(row, raw.shape)
        roi_area = int(float(row.get("ROI_Area_Pixels", row.get("Area_px", 0)) or 0))
        red_signal = green_signal = blue_signal = 0
        threshold = np.inf
    else:
        preferred_label = int(float(row.get("ROI_Label", row.get("Component_ID", idx + 1)) or idx + 1))
        minr, minc, maxr, maxc, roi_area, red_signal, green_signal, blue_signal, threshold = bbox_from_component(raw, preferred_label)
        minr, minc, maxr, maxc = clamp_bbox(
            {"BBox_Min_Row": minr, "BBox_Min_Col": minc, "BBox_Max_Row": maxr, "BBox_Max_Col": maxc},
            raw.shape,
        )
    raw_crop = raw[minr:maxr, minc:maxc, ...]

    safe = make_safe_name(image_name)
    image_raw_folder = os.path.join(RAW_CROP_FOLDER, safe)
    os.makedirs(image_raw_folder, exist_ok=True)

    component_id = int(float(row.get("Marker_Component_Label", row.get("ROI_Label", row.get("Component_ID", idx + 1))) or idx + 1))
    # The parent folder already contains the source-image name. Repeating that
    # name in every TIFF can exceed Windows MAX_PATH on normal experiment paths.
    crop_name = f"MARKER_COMPONENT_{component_id:04d}_CROP_{idx + 1:04d}.tiff"
    raw_out = os.path.join(image_raw_folder, crop_name)
    tiff.imwrite(raw_out, raw_crop, photometric="rgb" if raw_crop.ndim == 3 and raw_crop.shape[-1] in (3, 4) else None)

    out = {
        "Image": image_name,
        "ROI_Label": row.get("ROI_Label", ""),
        "Marker_Component_Label": component_id,
        "Status": "Cropped",
        "Crop_Mode": "marker_components",
        "Raw_Masked_Source_Image": raw_path,
        "Enhanced_Source_Image": "",
        "Raw_Masked_Crop_File": raw_out,
        "Enhanced_Crop_File": "",
        "Raw_Layout_Read": "YXC",
        "Enhanced_Layout_Read": "",
        "Crop_Min_Row": minr,
        "Crop_Min_Col": minc,
        "Crop_Max_Row": maxr,
        "Crop_Max_Col": maxc,
        "Crop_Height": maxr - minr,
        "Crop_Width": maxc - minc,
        "ROI_Area_Pixels": roi_area,
        "Red_Signal_Pixels": red_signal,
        "Green_Signal_Pixels": green_signal,
        "Blue_Signal_Pixels": blue_signal,
        "Red_Threshold": np.inf,
        "Green_Threshold": threshold,
        "Blue_Threshold": threshold,
        "Recommended_For_Quantitative_JACoP": "Raw_Masked_Crop_File",
        "Optional_Enhanced_Analysis_File": "",
    }
    out["Crop_Y0"] = out.pop("Crop_Min_Row")
    out["Crop_X0"] = out.pop("Crop_Min_Col")
    out["Crop_Y1"] = out.pop("Crop_Max_Row")
    out["Crop_X1"] = out.pop("Crop_Max_Col")
    return out


def main():
    global PREPROCESSOR_FOLDER, QC_EXCEL, RAW_MASKED_IMAGE_FOLDER, ENHANCED_IMAGE_FOLDER, MASK_FOLDER, OUTPUT_FOLDER, RAW_CROP_FOLDER, ENHANCED_CROP_FOLDER, OUTPUT_EXCEL
    PREPROCESSOR_FOLDER = PREPROCESSOR_FOLDER or os.path.join(EXPERIMENT_ROOT, "ENHANCED_JACOP_PREPROCESSOR_RESULTS")
    QC_EXCEL = QC_EXCEL or os.path.join(PREPROCESSOR_FOLDER, "Enhanced_Preprocessor_QC.xlsx")
    RAW_MASKED_IMAGE_FOLDER = RAW_MASKED_IMAGE_FOLDER or os.path.join(PREPROCESSOR_FOLDER, "01_raw_masked_for_jacop")
    ENHANCED_IMAGE_FOLDER = ENHANCED_IMAGE_FOLDER or os.path.join(PREPROCESSOR_FOLDER, "02_enhanced_masked_for_visual_qc")
    MASK_FOLDER = MASK_FOLDER or os.path.join(PREPROCESSOR_FOLDER, "masks")
    OUTPUT_FOLDER = OUTPUT_FOLDER or os.path.join(PREPROCESSOR_FOLDER, "03_cropped_rois_for_final_choice")
    RAW_CROP_FOLDER = RAW_CROP_FOLDER or os.path.join(OUTPUT_FOLDER, "01_raw_masked_roi_crops_for_jacop")
    ENHANCED_CROP_FOLDER = ENHANCED_CROP_FOLDER or os.path.join(OUTPUT_FOLDER, "02_enhanced_roi_crops_for_optional_analysis")
    OUTPUT_EXCEL = OUTPUT_EXCEL or os.path.join(OUTPUT_FOLDER, "Raw_Cropped_ROI_Index.xlsx")
    moved = archive_previous_outputs()
    if moved:
        print("Archived previous crop outputs:")
        for path in moved:
            print(f"  - {path}")
    for folder in (OUTPUT_FOLDER, RAW_CROP_FOLDER):
        os.makedirs(folder, exist_ok=True)

    print(f"Using QC Excel: {QC_EXCEL}")
    df = read_qc_rows()
    print(f"Selected ROI rows found: {len(df)}")
    workers = min(MAX_WORKERS, max(1, len(df)))
    print(f"Using {workers} parallel crop worker(s).")
    rows = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(crop_one, item) for item in df.iterrows()]
        for future in as_completed(futures):
            rows.append(future.result())

    crop_df = pd.DataFrame(rows)
    columns = [
        "Image",
        "ROI_Label",
        "Marker_Component_Label",
        "Status",
        "Crop_Mode",
        "Raw_Masked_Source_Image",
        "Enhanced_Source_Image",
        "Raw_Masked_Crop_File",
        "Enhanced_Crop_File",
        "Raw_Layout_Read",
        "Enhanced_Layout_Read",
        "Crop_Y0",
        "Crop_Y1",
        "Crop_X0",
        "Crop_X1",
        "Crop_Height",
        "Crop_Width",
        "ROI_Area_Pixels",
        "Red_Signal_Pixels",
        "Green_Signal_Pixels",
        "Blue_Signal_Pixels",
        "Red_Threshold",
        "Green_Threshold",
        "Blue_Threshold",
        "Recommended_For_Quantitative_JACoP",
        "Optional_Enhanced_Analysis_File",
    ]
    crop_df = crop_df.reindex(columns=columns)
    summary = crop_df.groupby("Image", dropna=False).size().reset_index(name="Crops_Created") if not crop_df.empty else pd.DataFrame(columns=["Image", "Crops_Created"])
    summary["Crop_Failures"] = 0
    with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
        crop_df.to_excel(writer, sheet_name="Raw Crop Index", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print("DONE")
    print(f"Raw masked ROI crops saved here: {RAW_CROP_FOLDER}")
    print(f"Index Excel saved here: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
