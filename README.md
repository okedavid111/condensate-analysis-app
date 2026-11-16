# Condensate Analysis App

Version 1.0.0

The Condensate Analysis App is a cross-platform, local Python application for reproducible microscopy image review, cleaning, segmentation, colocalization, and condensate quantification. It replaces manual Photoshop background cleanup with documented NumPy, SciPy, and scikit-image processing and programmatically runs the required Fiji/ImageJ operations.

The app was developed for two- and three-channel fluorescence TIFF images. It provides:

- cell/ROI selection, cropping, and blinded keep/reject quality control;
- interactive channel cleaning with saved per-image parameters and raw-versus-cleaned previews;
- nuclear and cellular ROI segmentation, puncta detection, and watershed separation;
- Fiji/ImageJ JACoP Pearson correlation and directional Manders colocalization;
- Fiji Analyze Particles condensate count, area, and equivalent-diameter measurements;
- manual ROI-based partition-coefficient measurements;
- fixed-cell MS2/condensate intensity and distance measurements; and
- auditable QC overlays and Excel exports.

The original TIFF images are not overwritten. Cleaned images are saved as derived files and can be selected as inputs for subsequent PCC/Manders and condensate count/size analyses. Workflows that require raw intensity values retain a separately selectable raw or untouched-acquisition source.

## Availability and scope

This repository contains the versioned source code used for the public software release. It deliberately does **not** include Fiji/ImageJ, JACoP, a Python installer, microscopy data, experiment state, logs, thumbnails, generated results, or historical working copies. Fiji/ImageJ and JACoP are third-party software and must be installed separately.

The browser interface is served only on `127.0.0.1`; image data are processed locally and are not uploaded to an external service.

## Requirements

- macOS or 64-bit Windows
- Python 3.11 or newer
- Fiji/ImageJ with the JACoP plugin for PCC/Manders workflows
- the Python packages pinned in `requirements.txt`

Fiji is available from <https://fiji.sc/>. After installing Fiji, install JACoP through Fiji's updater or place `JACoP_.class` in the Fiji `plugins` directory. The setup checker verifies the Python dependencies, Fiji executable, and JACoP installation.

## Installation

### macOS

1. Install Python 3.11 or newer and Fiji/ImageJ with JACoP.
2. Download or clone this repository into a writable folder.
3. Double-click `Setup macOS.command` once.
4. Double-click `CHECK INSTALLATION.command`.
5. Double-click `OPEN CONDENSATE ANALYSIS APP.command`.

If Fiji is installed somewhere other than `/Applications/Fiji.app`, set `FIJI_EXECUTABLE` to the Fiji launcher before starting the app.

### Windows

1. Install 64-bit Python 3.11 or newer and Fiji/ImageJ with JACoP.
2. Download or clone this repository into a writable folder.
3. Run `setup_windows.bat` once.
4. Run `CHECK INSTALLATION WINDOWS.bat`.
5. Run `run_app_windows.bat`.

The app searches common Fiji locations. You can set the `FIJI_EXECUTABLE` environment variable to the complete Fiji executable path when needed.

### Command line

```bash
python3 -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
python -m pip install -r requirements.txt
python condensate_analysis_app.py
```

Open the local address printed in the terminal. Stop the app with `Ctrl+C`.

## Analysis overview

1. Select a raw-image folder and output location.
2. Run preprocessing and ROI cropping, or load an existing compatible job folder.
3. Review every ROI and mark it keep or reject before viewing downstream quantitative measurements.
4. Assign channel roles and, when cleaning is used, adjust the background-suppression controls while inspecting the preview. The app saves the cleaned derivative and its parameters.
5. Select the appropriate input source and run PCC/Manders, condensate count/size, partition coefficient, or fixed-cell MS2 analysis.
6. Inspect the QC overlays and exported workbook before statistical analysis.

See [docs/ANALYSIS_METHODS.md](docs/ANALYSIS_METHODS.md) for calculations and [docs/REPRODUCIBILITY.md](docs/REPRODUCIBILITY.md) for reporting and provenance guidance.

## Reproducibility notes

- Version 1.0.0 uses the exact Python package versions in `requirements.txt`.
- Analysis parameters, thresholds, channel assignments, ROI decisions, and source paths are stored in the job state and/or exported workbooks.
- The app does not resize or spatially resample images and does not perform a Z-projection.
- Data and generated workbooks are ignored by Git to reduce the risk of accidental public release.
- A syntax-level release check is provided by `python -m compileall .`; Fiji-dependent numerical validation requires representative microscopy images and a working Fiji/JACoP installation.

## Citation

Please cite the version-specific Zenodo record for the exact release used. Citation metadata are also provided in [`CITATION.cff`](CITATION.cff). The DOI will be added to the repository metadata after the first GitHub release is archived by Zenodo.

## Authorship and development

The software was designed and developed by David Oke using AI-assisted coding. Scientific decisions, workflow design, testing, parameter selection, and responsibility for the released software remain with the author.

## License

The release license will be recorded in the repository before public publication.

## Third-party software

Fiji/ImageJ and JACoP are not distributed in this repository. See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
