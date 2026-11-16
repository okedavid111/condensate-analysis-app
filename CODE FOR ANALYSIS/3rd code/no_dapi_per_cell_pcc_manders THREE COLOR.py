#!/usr/bin/env python3
"""Recovered stage 3 three-color PCC and Manders analysis."""

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


PAIRS = (("Red_Green", 0, 1), ("Red_Blue", 0, 2), ("Green_Blue", 1, 2))


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


def pair_metrics(path, rgb, pair_name, a_idx, b_idx):
    a = rgb[..., a_idx].astype(float)
    b = rgb[..., b_idx].astype(float)
    mask = (a > 0) | (b > 0)
    if np.count_nonzero(mask) < 3:
        pcc = np.nan
    else:
        av = a[mask]
        bv = b[mask]
        pcc = float(np.corrcoef(av, bv)[0, 1]) if np.std(av) > 0 and np.std(bv) > 0 else np.nan
    a_pos = positive_mask(a) & mask
    b_pos = positive_mask(b) & mask
    overlap = a_pos & b_pos
    return {
        "Image": os.path.basename(path),
        "Path": path,
        "Pair": pair_name,
        "Pixels_Measured": int(np.count_nonzero(mask)),
        "Pearson": pcc,
        "M1_original": float(a[b_pos].sum() / a[mask].sum()) if a[mask].sum() else np.nan,
        "M2_original": float(b[a_pos].sum() / b[mask].sum()) if b[mask].sum() else np.nan,
        "M1_thresholded": float(a[overlap].sum() / a[a_pos].sum()) if a[a_pos].sum() else np.nan,
        "M2_thresholded": float(b[overlap].sum() / b[b_pos].sum()) if b[b_pos].sum() else np.nan,
    }


def main():
    global OUTPUT_EXCEL, MASK_FOLDER, ROI_PREVIEW_FOLDER
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    MASK_FOLDER = MASK_FOLDER or os.path.join(OUTPUT_FOLDER, "per_cell_masks")
    ROI_PREVIEW_FOLDER = ROI_PREVIEW_FOLDER or os.path.join(OUTPUT_FOLDER, "roi_preview_images")
    os.makedirs(MASK_FOLDER, exist_ok=True)
    os.makedirs(ROI_PREVIEW_FOLDER, exist_ok=True)
    OUTPUT_EXCEL = OUTPUT_EXCEL or os.path.join(OUTPUT_FOLDER, "Python_JACoP_Like_Enrichment_Triple_Positive_ROI_Results.xlsx")
    rows = []
    for path in list_tiffs(IMAGE_FOLDER):
        rgb = as_rgb(tiff.imread(path))
        for pair in PAIRS:
            rows.append(pair_metrics(path, rgb, *pair))
    with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="Per_ROI", index=False)
    print("DONE")
    print(f"Output Excel: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
