# Climate_Change_Analysis_2 review hand-off

Workbook version 1.7.11. The computed content is the v1.7.4 file as saved by
Microsoft Excel 16.0300, with one later change to an input: v1.7.10 re-ingested the
RONI series from the published ascii file at two decimals, which moved 3,923 cells,
all on `Open Items`. Everything else since v1.7.4 has been text: v1.7.5 added three
reference tabs, v1.7.6 and v1.7.7 corrected them, v1.7.8 stated the terms of use,
v1.7.9 corrected the ClimateIndicator licence to MIT, and v1.7.11 closed the last of
the provenance items.

Start on `Summary`. Then, depending on what you need: `Change log` for the
revision history, `Acronyms` for any abbreviation or symbol in the file,
`Glossary` for the statistical and climate terms in the sense this workbook uses
them, and `Sources` for input provenance and the script inventory. If you are
about to edit a formula, read the last two `Glossary` entries first; they are the
two traps this file has already fallen into.

## What is in the package

| File | What it is |
|---|---|
| `Climate_Change_Analysis_2.xlsx` | the workbook, 22 sheets |
| `Climate_Change_Analysis_2.md` | full Markdown export: every cell, every formula |
| `Climate_Change_Analysis_2_audit.md` | audit extract: formulas, notes and headline cells, no bulk data |
| `convert_to_md.py` | produces both exports |
| `residual_diagnostics.py` | regression diagnostics behind the `Diagnostics` tab |
| `validate_kernels.py` | reconstructs the `Kernel` tab, including the blocked cross-validation |
| `erf_uncertainty.py` | reconstructs the `Sensitivity` tab, including the Monte Carlo |
| `open_items.py` | the `Open Items` tab: RONI, the aerosol-GHG split, the albedo candidate |
| `erf_inputs/` | `ERF_best_aggregates.csv`, `ERF_p05_aggregates.csv`, `ERF_p95_aggregates.csv` |
| `roni_v5_cpc.csv` | 917 monthly RONI values transcribed from the CPC HTML table |
| `MANIFEST.sha256` | checksums for the other files |
| `RONI.v5.ascii.txt` | the published RONI source, ERSSTv5 basis, as used |
| `RONI.ascii.txt` | the operational RONI series, shipped only so the two-series comparison can be re-run; NOT an input |
| `roni_v5_cpc_1dp_superseded.csv` | the 2026-10-22 one-decimal transcription, kept so the rounding check and the quantisation test can be reproduced |
| `verify_roni_ingest.py` | re-runs the ingestion checks: the CSV against the ascii file, the rounding convention against five candidates, and the two-series comparison |
| `map_roni_rows.py` | proves the row-to-season mapping of `Open Items` column C before anything writes to it |
| `roni_quantisation_test.py` | historical: what the one-decimal series was predicted to cost, kept as the record of that prediction against the outcome |
| `EXPECTED_v1710.md`, `EXPECTED_v1711.md` | the acceptance tests those versions were checked against, and their outcomes |
| `LICENSE` | MIT, for the five scripts |
| `LICENSE-CONTENT.md` | CC BY 4.0 for the workbook, and the terms of every input |
| `THIRD-PARTY-NOTICES` | upstream notices that must travel with copies |

## Check the package before you start

```
sha256sum -c MANIFEST.sha256
```

The layout matters for one command. The three ERF CSVs must sit in `erf_inputs/`,
not beside the workbook, or `erf_uncertainty.py` fails as written below. This is
the tree the commands in this document assume:

```
Climate_Change_Analysis_2.xlsx
Climate_Change_Analysis_2.md
Climate_Change_Analysis_2_audit.md
HANDOFF.md
MANIFEST.sha256
convert_to_md.py
erf_uncertainty.py
open_items.py
residual_diagnostics.py
validate_kernels.py
roni_v5_cpc.csv
erf_inputs/ERF_best_aggregates.csv
erf_inputs/ERF_p05_aggregates.csv
erf_inputs/ERF_p95_aggregates.csv
```

If the CSVs arrive flat instead, either move them into `erf_inputs/` or pass `.`
as the second argument. Moving them is better, because then this document and the
package agree, which is what went wrong the first time it was zipped.

## How to run them

Run from the directory holding the workbook. Two of the four need their inputs
named; the directory argument to `erf_uncertainty.py` is the one people get wrong,
because it wants the folder, not a file, and the default is not the working
directory.

```
python3 residual_diagnostics.py
python3 validate_kernels.py
python3 erf_uncertainty.py Climate_Change_Analysis_2.xlsx erf_inputs 20000
python3 open_items.py Climate_Change_Analysis_2.xlsx roni_v5_cpc.csv
```

The RONI helpers take no arguments and resolve everything relative to their own
location, so they run from wherever the package is unpacked:

```
python3 verify_roni_ingest.py
python3 map_roni_rows.py
python3 roni_quantisation_test.py 40
```

`verify_roni_ingest.py` is the one to run first if you doubt the RONI input: it
checks the shipped CSV against `RONI.v5.ascii.txt` value by value, identifies
which of five rounding conventions explains the superseded transcription, and
quantifies how far the operational series differs from the one used here.

The first two default to `Climate_Change_Analysis_2.xlsx` and take an optional
path argument. `erf_uncertainty.py` takes optional draws and seed after the
directory; the defaults are 20,000 and 20261022, and the Monte Carlo triplets in
`Sensitivity` reproduce at those values.

Only two of the four self-check. `validate_kernels.py` and `erf_uncertainty.py`
compare themselves against the workbook cell by cell and mark each one `ok` or
`DIFFERS`, so grep for `DIFFERS`. A clean run has exactly one: `Kernel!P69`, the
CO2-only pooled R² whose 0.006 gap has no recoverable provenance and which the
script is deliberately written to report rather than pass. It prints `DIFFERS BY`,
which the short pattern still catches. Any other `DIFFERS` is a finding.

`residual_diagnostics.py` and `open_items.py` do not self-check. They print their
own results next to the workbook baselines they started from and leave the
comparison to the reader. A silent run from those two means nothing was compared,
not that everything matched, so read their output rather than grepping it.

Versions used for the runs reported in the workbook: Python 3.13.16, numpy 2.5.3,
scipy 1.18.1, openpyxl 3.1.5, LibreOffice 24.2.7.2.

## Regenerating the Markdown

```
soffice --headless --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir recalc Climate_Change_Analysis_2.xlsx
python3 convert_to_md.py Climate_Change_Analysis_2.xlsx recalc/Climate_Change_Analysis_2.xlsx Climate_Change_Analysis_2.md full
python3 convert_to_md.py Climate_Change_Analysis_2.xlsx recalc/Climate_Change_Analysis_2.xlsx Climate_Change_Analysis_2_audit.md audit
```

The second argument is a recalculated copy, used only where a formula cell has no
cached value and the recalculation supplies one. At v1.7.11 no cell meets both
conditions, so nothing in the export comes from it, and the preamble says so. Note
what that does not mean: 393 formula cells do carry no cached value, on nine tabs.
They are blank-by-design guards whose recalculated value is blank too, so there is
nothing to fall back to. Keep passing the argument anyway, because adding formulas
and exporting without saving in Excel first would bring the fallback back, and the
preamble is what tells a reviewer whether it fired.

If you edit the workbook, save it in Excel before exporting, and prefer Excel over
LibreOffice for the save. The export is only as authoritative as the cache it
reads.

## Terms

The workbook, its notes and the Markdown exports are CC BY 4.0; the five scripts
are MIT. Neither covers the input data, which keeps the terms recorded in the
workbook's `Sources` Terms column, most of it United States Government work in
the public domain. `LICENSE-CONTENT.md` has the detail, including the four
attributions that survive onward reuse. As of v1.7.11 nothing is outstanding: every
`CHECK` marker the file carried has been resolved against a publisher record, an
agency document or a checksummed download, and `Sources` section F records each
closure. `Sources` section F states the same inside the workbook,
so a copy separated from this package still carries its terms.

## Two copies of every version, and which one is canonical

Each version of this workbook exists as two files with different checksums, and
confusing them has already cost a review cycle.

The **build** is the script output. It inherits Excel's cached values from the
v1.7.4 save, so `docProps/app.xml` reports Microsoft Excel even though a script
wrote the file. That field says where the numbers came from, not who wrote the
copy you are holding.

The **release** is that build opened and saved in Excel, which recomputes every
formula and writes its own cache. This is the canonical copy. It is what the
author distributes and what `MANIFEST.sha256` should describe.

From v1.7.9 the release is what ships in this package, so the manifest describes
it and a workbook checksum should match. If one ever fails, check this before
concluding anything: a build and a release of the same version differ by
construction, and neither is stale or wrong. The `Change log` records the
version; the checksum alone cannot tell you which of the two you have.

Two facts about what an Excel save does and does not change, both measured on the
v1.7.9 pair. It alters nothing that matters: all 107,465 formulas, all 120,565
numeric values and every text cell are identical between the build and the
release. But it renumbers the worksheet parts inside the package, so
`xl/worksheets/sheet18.xml` in a build is not the same sheet in a release.
Anything doing XML surgery on this file must locate parts by their content, never
by part name.

## Do not read calcPr through openpyxl

`openpyxl` reports `wb.calculation.fullCalcOnLoad == True` for every file,
including ones whose raw XML has `<calcPr calcId="191029"/>` with no such
attribute, because that is the default on its `CalcProperties` class rather than
something read from the file. It is therefore useless as a way to tell a build
from a release. Read `xl/workbook.xml` directly instead.

The same caution applies to counting uncached formulas through `openpyxl`:
`data_only` returns `None` both for a formula with no cached value and for one
cached as an empty string, so 393 cells on nine tabs look uncached either way.
They are blank-by-design guards, formulas such as `IF(x="","",...)` whose result
is empty. A raw-XML check distinguishes the two cases; `openpyxl` cannot.

## Read this before trusting a figure

The two engines are not interchangeable, and the file carries a scar to prove it.
Through v1.7.2 seven formulas in `Diagnostics` section C passed a lone boolean
array to `SUMPRODUCT`, as in `=SUMPRODUCT(($B$9:$B$822>=2016))`. LibreOffice
coerces such an array to numbers. Excel does not: it returned 0, and the eighteen
cells dividing by those seven showed `#DIV/0!` to every Excel user from v1.1
onward. The pasted figures, all four scripts and every Markdown export were right
the whole time, which is exactly why nobody caught it: each review checked
LibreOffice's output against LibreOffice-derived material. It surfaced only when
the workbook was opened and saved in Excel and that cache was compared against the
export.

That is now closed. v1.7.3 rewrote the seven as `SUMPRODUCT(--(...))`, which both
engines evaluate identically, and v1.7.4 is the Excel-saved file: every formula
that resolves to a value carries Excel's own, the export falls back to a
recalculation for nothing, and the two engines agree across 120,565 numeric cells,
which is every
computed cell in the file, with no difference beyond 1e-9 relative, and 107,465
formulas unchanged by the round trip. The workbook contains no error-valued cell.
The reference tabs added in v1.7.5 and corrected in v1.7.6 are text only, and all
120,565 of those cells still hold the value Excel cached at v1.7.4, checked cell by
cell after each rebuild.

Keep the habit anyway. A figure confirmed by one engine is unconfirmed, and this
package was three adversarial reviews deep before anyone looked at the file in the
program it was written for.

## Known open items, deliberately left open

These are judgements, not loose ends. Nothing here is unverified.

- `Kernel!P69`, the CO2-only pooled R², differs from the script by 0.006 with no
  recoverable provenance under any partition, row range or metric tried. The
  script is written to report it rather than pass it, so a clean run shows exactly
  one `DIFFERS`.
- RONI is carried as a structural uncertainty of +0.06 rather than adopted,
  because substituting it makes the fit worse. Since v1.7.10 the series is the
  published two-decimal one; the figure is +0.0582.
- The GHG-plus-other and aerosol responses come out indistinguishable: g = 0.409,
  h = 0.414, g − h = −0.005 with a 95% interval of −0.20 to +0.19 (`Open Items`
  rows 39, 40, 45, 46). The two series are nearly orthogonal over the record,
  r = −0.002 at row 35, so the regression does separate them, and the weak
  interval means the data give no reason to prefer different responses rather than
  confirming equal ones. What is not identified is the aerosol response per W/m²:
  scaling the aerosol series by s scales h by exactly 1/s and leaves h·s and the
  R² untouched (row 50 against row 40), and AR6 puts that magnitude uncertain by a
  factor of four at 5-95%. So the aerosol contribution is identified and the
  response per W/m² is not. This is not a collinearity problem; section B rules
  one out. The reading note is at `Open Items!AA51`.
- A 2016-onward hold-out refit would be a new test, not a correction, and has not
  been run.
- `Forcing` note numbering runs 1,2,3,4,5,8,7,6. Cosmetic, left alone.
- `Sources` row 36 is a placeholder for a low-cloud and albedo estimate the
  workbook does not cite. It carries no claim.
- The aerosol percentile range in `erf_inputs` is −1.6003 to −0.3891 against the
  IGCC paper's −1.58 to −0.40. That is a pointwise percentile envelope against an
  ensemble quantile, a distinction `Sensitivity!AA52` already draws, and it is
  recorded at `Sources!L13` rather than reconciled.
