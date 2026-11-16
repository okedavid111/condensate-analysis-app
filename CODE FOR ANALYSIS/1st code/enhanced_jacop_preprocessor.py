#!/usr/bin/env python3
"""
Recovered Fiji PCC/Manders stage 1 preprocessor.

This replacement preserves the app-facing contract of the original script:
read raw TIFF images, build cell/marker masks, save raw masked JACoP-safe RGB
TIFFs, previews, masks, and an Enhanced_Preprocessor_QC.xlsx workbook.
"""

import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

import numpy as np
import pandas as pd
import tifffile as tiff
from PIL import Image, ImageDraw
from scipy import ndimage as ndi
from skimage import exposure, measure, morphology
from skimage.filters import threshold_otsu, threshold_triangle


INPUT_FOLDER = ""
OUTPUT_FOLDER_NAME = "ENHANCED_JACOP_PREPROCESSOR_RESULTS"
OUTPUT_FOLDER = ""
OUTPUT_EXCEL = ""
RAW_MASKED_JACOP_FOLDER = ""
ENHANCED_JACOP_FOLDER = ""
MASK_FOLDER = ""
PREVIEW_FOLDER = ""

BACKGROUND_VALUE = 0
MIN_NUCLEUS_AREA = 80
MARKER_MIN_OBJECT_AREA = 8
MARKER_EXPANSION_PIXELS = 2
MAX_SELECTED_ROIS_PER_IMAGE = 4
MIN_CELL_ROI_AREA = 120
MAX_WORKERS = 4
MERGED_IMAGE_NAME_MARKERS = ("merged", "merge")
NON_SOURCE_NAME_MARKERS = (
    "raw_masked",
    "jacop",
    "mask",
    "crop",
    "enhanced",
    "preview",
    "threshold",
    "count_size",
    "split",
    "channel",
)
DERIVED_DIR_NAME_MARKERS = (
    "enhanced_jacop_preprocessor_results",
    "final_reviewed",
    "real_fiji",
    "pcc",
    "manders",
    "analysis_results",
    "analysis",
    "channel_and_merged",
    "comparison",
    "processed_condensates",
    "fiji_measurement",
    "count",
    "size",
)


def make_safe_name(name):
    base = os.path.splitext(os.path.basename(str(name)))[0]
    return "".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in base)


def is_merged_source_image(path):
    name = os.path.basename(path).lower()
    stem, _ = os.path.splitext(name)
    if any(marker in stem for marker in NON_SOURCE_NAME_MARKERS):
        return False
    if any(marker in stem for marker in MERGED_IMAGE_NAME_MARKERS):
        return True
    return bool(re.search(r"(^|[_\s-])m(?=[_\s-])", stem, flags=re.IGNORECASE))


def is_derived_dir(name):
    lowered = name.lower()
    if lowered in ("", "."):
        return False
    if lowered.startswith("."):
        return True
    if lowered == OUTPUT_FOLDER_NAME.lower():
        return True
    return any(marker in lowered for marker in DERIVED_DIR_NAME_MARKERS)


def list_tiffs(folder):
    out = []
    excluded = {
        os.path.abspath(OUTPUT_FOLDER),
        os.path.abspath(RAW_MASKED_JACOP_FOLDER),
        os.path.abspath(ENHANCED_JACOP_FOLDER),
        os.path.abspath(MASK_FOLDER),
        os.path.abspath(PREVIEW_FOLDER),
    }
    for dirpath, dirnames, filenames in os.walk(folder):
        abs_dir = os.path.abspath(dirpath)
        if any(abs_dir == item or abs_dir.startswith(item + os.sep) for item in excluded if item):
            dirnames[:] = []
            continue
        if any(is_derived_dir(part) for part in os.path.relpath(abs_dir, os.path.abspath(folder)).split(os.sep)):
            dirnames[:] = []
            continue
        dirnames[:] = [name for name in dirnames if not is_derived_dir(name)]
        for name in sorted(filenames):
            if name.startswith("."):
                continue
            path = os.path.join(dirpath, name)
            if name.lower().endswith((".tif", ".tiff")) and is_merged_source_image(path):
                out.append(path)
    return out


def as_rgb(arr):
    arr = np.asarray(arr)
    if arr.ndim == 2:
        arr = np.stack([arr, arr, arr], axis=-1)
    elif arr.ndim == 3:
        if arr.shape[-1] >= 3:
            arr = arr[..., :3]
        elif arr.shape[0] >= 3 and arr.shape[-1] != 3:
            arr = np.moveaxis(arr[:3], 0, -1)
        else:
            arr = np.repeat(arr[..., :1], 3, axis=-1)
    elif arr.ndim == 4:
        arr = np.squeeze(arr)
        return as_rgb(arr)
    else:
        raise ValueError(f"Unsupported TIFF shape: {arr.shape}")
    if arr.dtype == np.uint8:
        return arr
    arr = arr.astype(np.float32)
    finite = arr[np.isfinite(arr)]
    hi = float(np.percentile(finite, 99.8)) if finite.size else 1.0
    if hi <= 0:
        hi = float(np.max(arr)) or 1.0
    arr = np.clip(arr / hi, 0, 1)
    return (arr * 255).astype(np.uint8)


def threshold_mask(channel, fallback_percentile=75):
    channel = np.asarray(channel)
    positive = channel[channel > 0]
    if positive.size < 16:
        return channel > 0
    try:
        thr = threshold_otsu(positive)
    except Exception:
        thr = np.percentile(positive, fallback_percentile)
    if thr <= 0:
        thr = np.percentile(positive, fallback_percentile)
    return channel > thr


def build_output_mask(rgb):
    blue = rgb[..., 2]
    green = rgb[..., 1]
    red = rgb[..., 0]

    nucleus = threshold_mask(blue, 65)
    nucleus = morphology.remove_small_objects(nucleus, MIN_NUCLEUS_AREA)
    nucleus = morphology.binary_closing(nucleus, morphology.disk(3))
    nucleus = ndi.binary_fill_holes(nucleus)

    marker = threshold_mask(np.maximum(red, green), 80)
    marker = morphology.remove_small_objects(marker, MARKER_MIN_OBJECT_AREA)
    marker = morphology.binary_dilation(marker, morphology.disk(MARKER_EXPANSION_PIXELS))

    if np.count_nonzero(nucleus) == 0:
        base = threshold_mask(np.max(rgb, axis=2), 70)
        base = morphology.remove_small_objects(base, MIN_NUCLEUS_AREA)
    else:
        base = nucleus
    output_mask = morphology.binary_dilation(base | marker, morphology.disk(2))
    output_mask = ndi.binary_fill_holes(output_mask)
    return output_mask.astype(bool), marker.astype(bool), nucleus.astype(bool)


def cell_candidate_mask(rgb, output_mask, marker_mask, nucleus_mask):
    color_mask = np.zeros(rgb.shape[:2], dtype=bool)
    for channel_index, percentile in ((0, 82), (1, 82), (2, 70)):
        channel_mask = threshold_mask(rgb[..., channel_index], percentile)
        color_mask |= channel_mask
    candidate = (output_mask | marker_mask | nucleus_mask) & color_mask
    candidate = morphology.binary_closing(candidate, morphology.disk(3))
    candidate = morphology.remove_small_objects(candidate, MIN_CELL_ROI_AREA)
    candidate = ndi.binary_fill_holes(candidate)
    return candidate.astype(bool)


def signal_stats(rgb, roi):
    area = int(np.count_nonzero(roi))
    stats = {"area": area, "score": 0.0}
    if area <= 0:
        for color in ("Red", "Green", "Blue"):
            stats[f"{color}_Mean_Over_Background"] = 0.0
            stats[f"{color}_Positive_Area_Fraction"] = 0.0
            stats[f"{color}_Positive"] = False
            stats[f"{color}_Signal_Pixels"] = 0
        return stats
    for color, channel_index in (("Red", 0), ("Green", 1), ("Blue", 2)):
        values = rgb[..., channel_index][roi]
        positive = values > 0
        positive_count = int(np.count_nonzero(positive))
        mean_value = float(np.mean(values)) if values.size else 0.0
        frac = float(positive_count / area)
        stats[f"{color}_Mean_Over_Background"] = mean_value
        stats[f"{color}_Positive_Area_Fraction"] = frac
        stats[f"{color}_Positive"] = bool(positive_count > 0)
        stats[f"{color}_Signal_Pixels"] = positive_count
        stats["score"] += mean_value * max(frac, 0.05)
    stats["score"] += min(area, 2000) * 0.03
    return stats


def save_preview(path, rgb, output_mask, marker_mask):
    preview = Image.fromarray(rgb).convert("RGB")
    overlay = Image.new("RGBA", preview.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    for mask, color in ((output_mask, (0, 255, 255, 95)), (marker_mask, (255, 255, 0, 105))):
        labels = measure.label(mask)
        for region in measure.regionprops(labels):
            minr, minc, maxr, maxc = region.bbox
            draw.rectangle([minc, minr, maxc, maxr], outline=color[:3] + (220,), width=2)
    Image.alpha_composite(preview.convert("RGBA"), overlay).convert("RGB").save(path)


def save_preview_set(base_path, rgb, raw_masked, output_mask, marker_mask):
    original = os.path.join(PREVIEW_FOLDER, f"{base_path}_01_original_display.png")
    enhanced = os.path.join(PREVIEW_FOLDER, f"{base_path}_02_enhanced_display.png")
    overlay = os.path.join(PREVIEW_FOLDER, f"{base_path}_03_roi_overlay.png")
    selected = os.path.join(PREVIEW_FOLDER, f"{base_path}_04_selected_preview.png")
    Image.fromarray(rgb).convert("RGB").save(original)
    Image.fromarray(raw_masked).convert("RGB").save(enhanced)
    save_preview(overlay, rgb, output_mask, marker_mask)
    save_preview(selected, raw_masked, output_mask, marker_mask)
    return [original, enhanced, overlay, selected]


def process_image(path):
    image_name = os.path.basename(path)
    base = make_safe_name(image_name)
    rgb = as_rgb(tiff.imread(path))
    output_mask, marker_mask, nucleus_mask = build_output_mask(rgb)

    raw_masked = np.array(rgb, copy=True)
    raw_masked[~output_mask, :] = BACKGROUND_VALUE

    raw_out = os.path.join(RAW_MASKED_JACOP_FOLDER, f"{base}_raw_masked_JACoP_ready.tiff")
    mask_out = os.path.join(MASK_FOLDER, f"{base}_mask.tiff")
    marker_out = os.path.join(MASK_FOLDER, f"{base}_marker_mask.tiff")

    tiff.imwrite(raw_out, raw_masked, photometric="rgb")
    tiff.imwrite(mask_out, output_mask.astype(np.uint8) * 255)
    tiff.imwrite(marker_out, marker_mask.astype(np.uint8) * 255)
    preview_paths = save_preview_set(base, rgb, raw_masked, output_mask, marker_mask)

    rows = []
    cell_mask = cell_candidate_mask(rgb, output_mask, marker_mask, nucleus_mask)
    labels = measure.label(cell_mask)
    regions = measure.regionprops(labels, intensity_image=np.max(rgb, axis=2))
    ranked_regions = []
    for region in regions:
        roi = labels == region.label
        stats = signal_stats(rgb, roi)
        if stats["area"] >= MIN_CELL_ROI_AREA and (
            stats["Red_Positive"] or stats["Green_Positive"] or stats["Blue_Positive"]
        ):
            ranked_regions.append((stats["score"], region, stats))
    ranked_regions.sort(key=lambda item: item[0], reverse=True)
    selected_labels = {int(region.label) for _, region, _ in ranked_regions[:MAX_SELECTED_ROIS_PER_IMAGE]}
    for idx, (_, region, stats) in enumerate(ranked_regions, start=1):
        minr, minc, maxr, maxc = region.bbox
        selected = int(region.label) in selected_labels
        rows.append({
            "Image_Name": image_name,
            "Base_Name": base,
            "Source_Path": path,
            "Raw_Masked_Path": raw_out,
            "Mask_Path": mask_out,
            "Marker_Mask_Path": marker_out,
            "Preview_Path": "; ".join(preview_paths),
            "Component_ID": idx,
            "Label": int(region.label),
            "Area_px": int(stats["area"]),
            "BBox_Min_Row": int(minr),
            "BBox_Min_Col": int(minc),
            "BBox_Max_Row": int(maxr),
            "BBox_Max_Col": int(maxc),
            "Centroid_Row": float(region.centroid[0]),
            "Centroid_Col": float(region.centroid[1]),
            "Mean_Intensity": float(region.mean_intensity),
            "selected_by_marker_rule": selected,
            "Keep_For_Cropping": selected,
            "Image": image_name,
            "ROI_Label": int(region.label),
            "ROI_Area_Pixels": int(stats["area"]),
            "Selected_For_JACoP": selected,
            "Selection_Rank": idx,
            "Selection_Score": float(stats["score"]),
            "Red_Mean_Over_Background": stats["Red_Mean_Over_Background"],
            "Red_Positive_Area_Fraction": stats["Red_Positive_Area_Fraction"],
            "Red_Positive": stats["Red_Positive"],
            "Green_Mean_Over_Background": stats["Green_Mean_Over_Background"],
            "Green_Positive_Area_Fraction": stats["Green_Positive_Area_Fraction"],
            "Green_Positive": stats["Green_Positive"],
            "Blue_Mean_Over_Background": stats["Blue_Mean_Over_Background"],
            "Blue_Positive_Area_Fraction": stats["Blue_Positive_Area_Fraction"],
            "Blue_Positive": stats["Blue_Positive"],
        })
    if not rows:
        rows.append({
            "Image_Name": image_name,
            "Base_Name": base,
            "Source_Path": path,
            "Raw_Masked_Path": raw_out,
            "Mask_Path": mask_out,
            "Marker_Mask_Path": marker_out,
            "Preview_Path": "; ".join(preview_paths),
            "Component_ID": 0,
            "Label": 0,
            "Area_px": int(np.count_nonzero(output_mask)),
            "BBox_Min_Row": 0,
            "BBox_Min_Col": 0,
            "BBox_Max_Row": int(rgb.shape[0]),
            "BBox_Max_Col": int(rgb.shape[1]),
            "Centroid_Row": float(rgb.shape[0] / 2),
            "Centroid_Col": float(rgb.shape[1] / 2),
            "Mean_Intensity": 0.0,
            "selected_by_marker_rule": True,
            "Keep_For_Cropping": True,
            "Image": image_name,
            "ROI_Label": 0,
            "ROI_Area_Pixels": int(np.count_nonzero(output_mask)),
            "Selected_For_JACoP": True,
            "Selection_Rank": 1,
            "Selection_Score": 0,
            "Red_Mean_Over_Background": 0,
            "Red_Positive_Area_Fraction": 0,
            "Red_Positive": False,
            "Green_Mean_Over_Background": 0,
            "Green_Positive_Area_Fraction": 0,
            "Green_Positive": False,
            "Blue_Mean_Over_Background": 0,
            "Blue_Positive_Area_Fraction": 0,
            "Blue_Positive": False,
        })
    return rows


def main():
    global OUTPUT_FOLDER, OUTPUT_EXCEL, RAW_MASKED_JACOP_FOLDER, ENHANCED_JACOP_FOLDER, MASK_FOLDER, PREVIEW_FOLDER
    if not OUTPUT_FOLDER:
        OUTPUT_FOLDER = os.path.join(INPUT_FOLDER, OUTPUT_FOLDER_NAME)
    OUTPUT_EXCEL = OUTPUT_EXCEL or os.path.join(OUTPUT_FOLDER, "Enhanced_Preprocessor_QC.xlsx")
    RAW_MASKED_JACOP_FOLDER = RAW_MASKED_JACOP_FOLDER or os.path.join(OUTPUT_FOLDER, "01_raw_masked_for_jacop")
    ENHANCED_JACOP_FOLDER = ENHANCED_JACOP_FOLDER or os.path.join(OUTPUT_FOLDER, "02_enhanced_masked_for_visual_qc")
    MASK_FOLDER = MASK_FOLDER or os.path.join(OUTPUT_FOLDER, "masks")
    PREVIEW_FOLDER = PREVIEW_FOLDER or os.path.join(OUTPUT_FOLDER, "previews")
    for folder in (OUTPUT_FOLDER, RAW_MASKED_JACOP_FOLDER, ENHANCED_JACOP_FOLDER, MASK_FOLDER, PREVIEW_FOLDER):
        os.makedirs(folder, exist_ok=True)

    paths = list_tiffs(INPUT_FOLDER)
    print(f"Found {len(paths)} merged TIFF images to process.")
    if not paths:
        raise RuntimeError(
            "No merged TIFF images were found. The preprocessor is intentionally "
            "restricted to source merged images only, so files such as masks, crops, "
            "JACoP inputs, enhanced previews, channel splits, and count/size outputs "
            "are skipped. Check that the raw image folder contains TIFF files with "
            "'merged' in the filename or the microscope channel marker '_M_'."
        )
    print(f"Using {min(MAX_WORKERS, max(1, len(paths)))} parallel image worker(s).")
    all_rows = []
    workers = min(MAX_WORKERS, max(1, len(paths)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_map = {pool.submit(process_image, path): path for path in paths}
        for future in as_completed(future_map):
            path = future_map[future]
            print(f"Processing: {os.path.basename(path)}")
            all_rows.extend(future.result())

    df = pd.DataFrame(all_rows)
    summary_rows = []
    if not df.empty:
        for image_name, group in df.groupby("Image_Name", sort=True):
            first = group.iloc[0]
            selected_count = int(group["Selected_For_JACoP"].astype(bool).sum())
            summary_rows.append({
                "Image": image_name,
                "Layout": "YXC",
                "Total_ROIs": int(len(group)),
                "Selected_ROIs": selected_count,
                "Recommended_JACoP_File": first.get("Raw_Masked_Path", ""),
                "Raw_Masked_JACoP_File": first.get("Raw_Masked_Path", ""),
                "Enhanced_Masked_File": "",
                "Preview_File(s)": first.get("Preview_Path", ""),
            })
    image_summary = pd.DataFrame(summary_rows)
    roi_columns = [
        "Image_Name",
        "Base_Name",
        "Source_Path",
        "Raw_Masked_Path",
        "Mask_Path",
        "Marker_Mask_Path",
        "Preview_Path",
        "Component_ID",
        "ROI_Label",
        "ROI_Area_Pixels",
        "BBox_Min_Row",
        "BBox_Min_Col",
        "BBox_Max_Row",
        "BBox_Max_Col",
        "Centroid_Row",
        "Centroid_Col",
        "Mean_Intensity",
        "Keep_For_Cropping",
        "Selected_For_JACoP",
        "Selection_Rank",
        "Selection_Score",
        "Red_Mean_Over_Background",
        "Red_Positive_Area_Fraction",
        "Red_Positive",
        "Green_Mean_Over_Background",
        "Green_Positive_Area_Fraction",
        "Green_Positive",
        "Blue_Mean_Over_Background",
        "Blue_Positive_Area_Fraction",
        "Blue_Positive",
        "Image",
    ]
    roi_df = df.reindex(columns=roi_columns) if not df.empty else pd.DataFrame(columns=roi_columns)
    with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
        image_summary.to_excel(writer, sheet_name="Image Summary", index=False)
        roi_df.to_excel(writer, sheet_name="All ROIs", index=False)
        roi_df[roi_df["Selected_For_JACoP"] == True].to_excel(writer, sheet_name="Selected ROIs", index=False)

    print("DONE")
    print(f"Output saved here: {OUTPUT_FOLDER}")
    print(f"Raw masked JACoP-safe images: {RAW_MASKED_JACOP_FOLDER}")
    print(f"Enhanced visual-QC images: {ENHANCED_JACOP_FOLDER}")
    print(f"Previews: {PREVIEW_FOLDER}")


if __name__ == "__main__":
    main()
