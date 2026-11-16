# Reproducibility and release provenance

## Record for each analysis

Retain the following with the raw TIFF data:

- the exact software release and version-specific DOI;
- operating system, Python version, Fiji/ImageJ version, and JACoP installation;
- the job-state JSON and threshold files produced for the dataset;
- kept/rejected ROI decisions and the prespecified exclusion rule;
- channel labels and channel order;
- raw-versus-cleaned source selection for each analysis;
- cleaning settings, intensity thresholds, particle-size threshold, and pixel calibration;
- QC overlays and exported Excel workbooks; and
- any manual ROI, focus, condensate, or background annotations.

## Data safety

The application preserves source TIFFs and writes derived files to output folders. Use a separate output directory and retain an immutable copy of the acquisition data. Do not place sensitive or unpublished microscopy data inside a public Git repository.

## Public release contents

The public repository intentionally excludes experiment data, local paths, saved application state, logs, caches, result workbooks, bundled Fiji distributions, and Python installers. The `.gitignore` file provides an additional safeguard but is not a substitute for inspecting every commit before publication.

## Release validation

For version 1.0.0, all included Python files were syntax-compiled under Python 3.14.4, and the dependency versions were recorded. Before relying on a new installation, run the platform setup checker and confirm the numerical outputs with a representative dataset whose expected masks, particle counts, areas, PCC, and directional Manders values have been independently reviewed.
