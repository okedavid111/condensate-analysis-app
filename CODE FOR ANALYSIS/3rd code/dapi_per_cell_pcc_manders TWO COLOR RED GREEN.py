#!/usr/bin/env python3
"""Recovered stage 3 two-color red/green PCC and Manders analysis."""

import os

import numpy as np
import pandas as pd
import tifffile as tiff
from skimage.filters import threshold_otsu


IMAGE_FOLDER = ""
OUTPUT_FOLDER = ""
OUTPUT_EXCEL = ""
MASK_FOLDER = ""
ROI_PREVIEW_FOLDER = ""


def list_tiffs(folder):
    out = []
    for dirpath, _, filenames in os.walk(folder):
        for name in filenames:
            if name.lower().endswith((".tif", ".tiff")) and not name.startswith("."):
                out.append(os.path.join(dirpath, name))
    return sorted(out)


def as_rgb(arr):
    arr = np.asarray(arr)
    if arr.ndim == 2:
        arr = np.stack([arr, arr, arr], axis=-1)
    elif arr.ndim == 3 and arr.shape[0] >= 3 and arr.shape[-1] != 3:
        arr = np.moveaxis(arr[:3], 0, -1)
    elif arr.ndim == 3 and arr.shape[-1] >= 3:
        arr = arr[..., :3]
    elif arr.ndim == 4:
        return as_rgb(np.squeeze(arr))
    return arr.astype(np.float32)


def positive_mask(channel):
    vals = channel[channel > 0]
    if vals.size < 8:
        return channel > 0
    try:
        thr = threshold_otsu(vals)
    except Exception:
        thr = np.percentile(vals, 75)
    return channel > thr


def metrics(path):
    rgb = as_rgb(tiff.imread(path))
    red = rgb[..., 0].astype(float)
    green = rgb[..., 1].astype(float)
    mask = (red > 0) | (green > 0)
    if np.count_nonzero(mask) < 3:
        pcc = np.nan
    else:
        rv = red[mask]
        gv = green[mask]
        pcc = float(np.corrcoef(rv, gv)[0, 1]) if np.std(rv) > 0 and np.std(gv) > 0 else np.nan
    red_pos = positive_mask(red) & mask
    green_pos = positive_mask(green) & mask
    red_sum = float(red[mask].sum())
    green_sum = float(green[mask].sum())
    overlap = red_pos & green_pos
    return {
        "Image": os.path.basename(path),
        "Path": path,
        "Pixels_Measured": int(np.count_nonzero(mask)),
        "Pearson_Red_Green": pcc,
        "M1_original": float(red[green_pos].sum() / red_sum) if red_sum else np.nan,
        "M2_original": float(green[red_pos].sum() / green_sum) if green_sum else np.nan,
        "M1_thresholded": float(red[overlap].sum() / red[red_pos].sum()) if red[red_pos].sum() else np.nan,
        "M2_thresholded": float(green[overlap].sum() / green[green_pos].sum()) if green[green_pos].sum() else np.nan,
    }


def main():
    global OUTPUT_EXCEL, MASK_FOLDER, ROI_PREVIEW_FOLDER
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    MASK_FOLDER = MASK_FOLDER or os.path.join(OUTPUT_FOLDER, "per_cell_masks")
    ROI_PREVIEW_FOLDER = ROI_PREVIEW_FOLDER or os.path.join(OUTPUT_FOLDER, "red_green_roi_preview_images")
    os.makedirs(MASK_FOLDER, exist_ok=True)
    os.makedirs(ROI_PREVIEW_FOLDER, exist_ok=True)
    OUTPUT_EXCEL = OUTPUT_EXCEL or os.path.join(OUTPUT_FOLDER, "Python_JACoP_Like_Enrichment_Red_Green_Positive_ROI_Results.xlsx")
    rows = [metrics(path) for path in list_tiffs(IMAGE_FOLDER)]
    with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="Per_ROI", index=False)
    print("DONE")
    print(f"Output Excel: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
