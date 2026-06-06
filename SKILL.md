---
name: shear-angle-slope-measurement
description: Measure shear-test or granular-slope angles from simulation/rendered images. Use when Codex needs to crop a local region from an EDEM or similar particle-flow image, extract the top envelope of flowed material, fit single or segmented outside-flow slope lines, report local-coordinate fitting equations with y positive upward, annotate cropped PNGs, and write Chinese Excel results. For curved outflow surfaces with a steep former/front section and either flatter or steeper latter/rear section, use former, latter, and total fits unless the latter angle is near-horizontal.
---

# Shear Angle Slope Measurement

Use this skill to turn a shear/flow image into reproducible slope outputs:

- cropped local PNG: `Local/<image_stem>_local.png`
- annotated cropped PNGs under `Local`
- Excel workbook in the root directory, with sheet name `<image_stem>` plus the Chinese suffix for "fitting result".
- aggregate Excel worksheets named with the Chinese labels for "shear angle" and "discharge distance".

## Coordinate Convention

Use local cropped-image coordinates in final reporting unless the user explicitly asks for full-image coordinates.

- `x` positive direction: right.
- `y` positive direction: upward.
- Local origin: lower-left corner of the cropped image.
- Report the fitting equation as `y = mx + b` in the local cropped coordinate system.
- Use the Chinese label for "shear angle", not the Chinese label for "slope angle".
- Keep abbreviations such as `ROI`, `R^2`, `deg`, and `EDEM`, and add Chinese explanations when useful.

## Recommended Workflow

1. Identify the source image and root directory.
2. Determine the ROI crop rectangle `(x0, y0, x1, y1)` in original image pixel coordinates.
3. Segment dark material pixels, extract the top envelope column by column, and fit the outside-flow material slope.
4. If the outflow surface is curved or has a clear slope change between former/front and latter/rear sections, first try former/latter/total fitting. This applies both when the latter/rear section becomes flatter and when it becomes steeper. Use automatic split selection unless the user gives a split x-coordinate. In segmented mode, fit the full point set in each section so the former line starts near the material outflow height, the former end equals the latter start, the total start equals the former start, and the total end equals the latter end.
5. If the latter/rear fitted angle is not greater than 2 degrees, treat the latter section as near-horizontal runout and keep only the former/front fit. Do not output latter or total rows/images in this case unless the user explicitly asks for them.
6. Convert fitted lines to local y-up coordinates.
7. Measure horizontal discharge distance from the leftmost valid material point inside the transparent box to the rear endpoint of the outside-flow material. Name this measurement with the Chinese label for "discharge distance". Use the transparent box length as the scale reference; default actual length is 100 mm.
8. Save:
   - original crop PNG
   - annotated crop PNGs with only:
     - fitting equation
     - shear angle
     - R^2
     - fit point count `n`
     - red dashed discharge-distance marker and label
   - Excel result table in one per-image worksheet, without secondary tables
   - aggregate worksheets in the same workbook for shear-angle data and discharge-distance data.
9. Verify file existence, image dimensions, and key workbook cells.

## Script

Use `scripts/measure_shear_slope.py` for the standard workflow. It requires Python packages `Pillow`, `numpy`, and `openpyxl` from the bundled runtime when available.

Example matching the original `r1.png` workflow:

```powershell
& '<bundled-python.exe>' '<skill-folder>\scripts\measure_shear_slope.py' `
  --image 'C:\Users\HP\Desktop\simulated_results_of_shear_angle\r1.png' `
  --root 'C:\Users\HP\Desktop\simulated_results_of_shear_angle' `
  --crop 350 250 1160 700 `
  --fit-x-min 675 `
  --fit-x-max 1120 `
  --fit-y-max 670 `
  --workbook-name 'simulated_results_of_shear_angle.xlsx'
```

Adjust crop and fit bounds for other images. If the user provides a different image but no bounds, inspect the image first and choose bounds that include the box and the flowed material while excluding timestamps, axes, logos, and unrelated whitespace.

For a curved outflow surface, add:

```powershell
  --segment-mode auto
```

This writes former, latter, and total rows to the same worksheet and saves three annotated images. Use `--split-x <value>` only when a manual front/rear split is requested.

If the computed latter angle is at or below 2 degrees, the script keeps only the former/front fit. Override the cutoff only when requested:

```powershell
  --flat-latter-angle-max 2
```

The script also writes distance columns using the transparent box as scale. The default reference is 100 mm:

```powershell
  --reference-length-mm 100
```

## Output Rules

- Save cropped and annotated images under `Local` in the root directory.
- Name files from the source image stem:
  - `<stem>_local.png`
  - single-fit mode: Chinese prefix for "total" + `<stem>` + `fit.png`
  - segmented mode: Chinese prefix for "former" + `<stem>` + `fit.png`, Chinese prefix for "latter" + `<stem>` + `fit.png`, and Chinese prefix for "total" + `<stem>` + `fit.png`
- Name the Excel worksheet `<stem>` plus the Chinese suffix for "fitting result".
- Keep all per-image measurement worksheets in the same root workbook.
- Rebuild two aggregate worksheets in the same workbook:
  - Chinese label for "shear angle": shear-angle equation, angle, R^2, fit point count, and ROI.
  - Chinese label for "discharge distance": total-region reference length, reference pixels, discharge-distance pixels, discharge-distance millimeters, and ROI.
- In both aggregate worksheets, the first column header is the Chinese label for "group number", and its value is the source image stem without extension.
- In the shear-angle aggregate worksheet, merge the first-column group-number cells for rows that share the same image stem.
- In the discharge-distance aggregate worksheet, record only the total-region row for each image, not former/front or latter/rear rows.
- Excel columns:
  - image name
  - region
  - fitting equation in local coordinates with y positive upward
  - shear angle in degrees
  - R^2
  - fit point count `n`
  - ROI coordinates
  - transparent box reference length in mm
  - reference length in pixels
  - discharge distance in pixels
  - discharge distance in mm
- In segmented mode, write three rows in this order when the latter angle is greater than 2 degrees: former, latter, total.
- When the latter angle is at or below 2 degrees, write only the former/front row and save only the former/front annotated image.

## Validation

After running the script:

- Confirm the local crop dimensions equal `(x1 - x0, y1 - y0)`.
- Confirm the annotated image has readable Chinese text, not `?`.
- Confirm the workbook sheet name is `<image_stem>` plus the Chinese suffix for "fitting result".
- Confirm the workbook does not contain the full-image equation or a secondary result table unless the user requested them.
- Confirm distance columns are present, use the transparent box reference length, and are labeled as discharge distance.
- Confirm each annotated crop shows a red dashed discharge-distance line from the left material point to the rear endpoint.
