# Analysis methods and calculations

This file summarizes calculations implemented in version 1.0.0. It is not a substitute for recording image-acquisition settings, biological replicates, exclusion criteria, and statistical methods in a manuscript.

## Derived cleaned images

For each retained ROI, the user assigns channel roles and selects background-suppression settings while inspecting an interactive raw-versus-cleaned preview. The app applies the selected settings reproducibly, saves the derived image, and records the parameters. Original TIFF files are preserved unchanged. The derived cleaned image can be selected for subsequent PCC/Manders and condensate count/size analysis. Raw-intensity workflows retain an explicit raw or untouched-acquisition source so that cleaning is not silently substituted for quantitative raw intensity.

## PCC and directional Manders coefficients

The app prepares reviewed inputs and programmatically invokes Fiji/ImageJ JACoP. Pearson's correlation coefficient describes the linear correlation between the paired pixel intensities. Directional Manders coefficients are retained separately because M1 and M2 are not interchangeable: each reports the fraction of one channel's signal overlapping the other under the selected JACoP settings. The exported workbook records channel order and pair labels.

## Condensate number and size

For each retained ROI, the selected condensate channel is converted to 8-bit. The user selects an ROI-specific lower threshold; the upper threshold is fixed at 255. The app makes a binary mask, invokes Fiji Watershed to separate touching objects, and runs Fiji Analyze Particles with the selected minimum particle area and no circularity restriction.

For pixel size `p` in micrometres per pixel and particle area `A_px` in pixels squared:

```text
A_um2 = A_px * p^2
d_eq  = 2 * sqrt(A_px / pi) * p
```

`A_um2` is physical area in square micrometres and `d_eq` is the equivalent circular diameter in micrometres. The cell-level mean diameter is the arithmetic mean of the individual-particle equivalent diameters; it is not the diameter calculated from the summed area.

## Manual ROI-based partition coefficient

The user traces a measurement ROI, one or more condensate ROIs, and a cell-free background ROI. All selected channels are measured independently in the same masks. The background-corrected partition coefficient is:

```text
PC = (mean_condensate - mean_background) /
     (mean_dilute_phase - mean_background)
```

The diffuse intracellular signal is the denominator. The selected boundary exclusion removes edge pixels from both condensate and dilute-phase masks. Raw ratios and background-corrected values are exported.

## Fixed-cell MS2/condensate analysis

The fixed-cell workflow supports manual measurement/focus annotations, saved channel thresholds, optional DAPI nuclear segmentation, intact Fiji count/size particle reuse, or manually traced condensates. Raw or matched untouched-acquisition images can supply fluorescence intensities. The app exports one row per processed cell, per-condensate measurements, background-source provenance, QC status, and overlay paths. Full calculation choices and masks are preserved in the workbook so that automatic fallbacks can be reviewed.
