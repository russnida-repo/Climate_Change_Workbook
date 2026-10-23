#!/usr/bin/env python3
"""
convert_to_md.py

Converts Climate_Change_Analysis_2.xlsx to a single Markdown file for reading by
a language model, carrying BOTH the computed value and the formula behind it.

Formula handling
----------------
A cell whose formula is part of a long fill down a column (the 822-month data
series) would make the file unreadable if the formula were repeated on every
row. Those columns are therefore described once, in a per-sheet "Formula
columns" legend, and their cells show values only. Every other formula cell -
the statistics panels, the headline links, the regression blocks - carries its
formula inline next to the value, since those are the cells an auditor needs to
check.

Values
------
Values are the ones Excel cached. From v1.7.4 the workbook ships Excel-saved, so
every formula carries one and nothing here is recalculated; the preamble each run
emits states which tabs, if any, fell back.

The fallback exists because formulas added after the last save have no cached
value. Through v1.7.3 that covered the whole Diagnostics and Open Items tabs,
which the builder shipped with fullCalcOnLoad set so Excel would compute them on
open; for those cells the value came from a LibreOffice recalculation of the same
file, separately checked against the scripts. Excel's own save drops that flag,
because with every value cached there is nothing to force. Keep the mechanism:
editing the workbook and exporting without saving in Excel first brings it
straight back.

A LibreOffice recalculation is not a stand-in for Excel. Through v1.7.2 the two
engines disagreed on 25 cells of Diagnostics section C, where seven formulas
passed a lone boolean array to SUMPRODUCT: LibreOffice coerces such an array to
numbers, Excel does not and returned 0, so those seven evaluated to zero in Excel
and the eighteen cells dividing by them showed #DIV/0!. A formula defect, not an
engine difference, fixed in v1.7.3 and confirmed by Excel in v1.7.4. The two
engines are now reconciled across the whole file: 120,565 numeric cells, which is
every computed cell in it, with no difference beyond 1e-9 relative.

Usage: python3 convert_to_md.py <workbook.xlsx> <recalculated.xlsx> <out.md>
"""

import re
import sys
import datetime
import openpyxl
from openpyxl.utils import get_column_letter

REF = re.compile(r'(?<![A-Za-z0-9_.$])(\$?)([A-Z]{1,3})(\$?)(\d+)')
MIN_RUN = 5          # a column fill of this many consecutive identical formulas
                     # goes to the legend instead of being repeated inline


def nonempty(v):
    return v is not None and (not isinstance(v, str) or v.strip() != "")


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour == 0 and v.minute == 0) \
            else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        if v != v or v in (float("inf"), float("-inf")):
            return str(v)
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return f"{v:.10g}"
    s = str(v).replace("\r\n", "\n").replace("\r", "\n")
    # leading spaces carry meaning in this workbook (they nest sub-items inside a
    # panel), so only the trailing end is trimmed
    return s.replace("|", "\\|").replace("\n", " / ").rstrip()


def esc(s):
    return s.replace("|", "\\|").replace("\n", " ")


def normalize(formula, row):
    """Row-relative form, so one fill down a column normalizes identically."""
    def rep(m):
        d1, col, d2, r = m.groups()
        if d2 == "$":
            return f"{d1}{col}$={r}"
        return f"{d1}{col}R{int(r) - row:+d}"
    return REF.sub(rep, formula)


def main(path_wb, path_recalc, out_path, mode="full"):
    audit = (mode == "audit")
    wbf = openpyxl.load_workbook(path_wb, data_only=False)
    wbv = openpyxl.load_workbook(path_wb, data_only=True)
    wbr = openpyxl.load_workbook(path_recalc, data_only=True)
    # version and date are read from the workbook's own cells, not hardcoded
    try:
        ver = str(wbv["Summary"]["C48"].value or "?")
        ver_date = str(wbv["Summary"]["C49"].value or "?")
    except Exception:
        ver, ver_date = "?", "?"
    stamp = f"workbook version {ver}, {ver_date}"
    # which cells ship without a cached value, so the preamble can say so accurately
    _nocache = {}
    for _ws in wbf.worksheets:
        _v = wbv[_ws.title]
        _n = 0
        _r = wbr[_ws.title]
        for _row in _ws.iter_rows():
            for _c in _row:
                # count only where the fallback actually supplied a value: a formula
                # with no cached value whose recalculation is not itself blank
                if (isinstance(_c.value, str) and _c.value.startswith("=")
                        and _v.cell(_c.row, _c.column).value is None
                        and _r.cell(_c.row, _c.column).value is not None):
                    _n += 1
        if _n:
            _nocache[_ws.title] = _n
    # count the formula cells that carry no cached value at all, so the preamble
    # can say how many rather than implying there are none
    _blank = 0
    for _ws in wbf.worksheets:
        _v = wbv[_ws.title]
        for _row in _ws.iter_rows():
            for _c in _row:
                if (isinstance(_c.value, str) and _c.value.startswith("=")
                        and _v.cell(_c.row, _c.column).value is None):
                    _blank += 1
    if _nocache:
        _parts = [f"`{k}` ({v:,} cells)" if v > 50 else f"`{k}` ({v})"
                  for k, v in sorted(_nocache.items(), key=lambda kv: -kv[1])]
        _nocache_text = ("The cells whose value here comes from that recalculation rather "
                         "than from the cache: " + "; ".join(_parts) + ". Everything else "
                         "is Excel's own cached value.")
    else:
        _nocache_text = ("No value in this export comes from the recalculation. Every formula "
                         "cell that carries a cached value is shown with Excel's own; the "
                         f"{_blank:,} that carry none are blank-by-design guards, formulas such as "
                         "IF(x=\"\",\"\",...) whose result is empty, so the recalculation is "
                         "blank for them too and there is nothing to fall back to.")

    L = []
    A = L.append
    if audit:
        A("# Climate Change Analysis 2: audit extract (panels, statistics and all formulas)")
        A("")
        A(f"Converted from `Climate_Change_Analysis_2.xlsx` ({stamp}) to")
        A(f"Markdown for reading by a language model. This is the **audit extract**: all "
          f"{len(wbf.worksheets)} worksheets,")
        A("every formula in the workbook, every assumption block, regression panel, statistic and")
        A("commentary cell, but **the long input data series are represented by their first and last")
        A("three rows only**. A row is kept whenever any of its cells sits in a column holding fewer")
        A("than 100 entries, which keeps all panel and label content and drops only the repetitive")
        A("month-by-month and year-by-year rows. Those rows are fully reproducible from the formulas")
        A("listed here plus the cited public datasets, and the complete file")
        A("`Climate_Change_Analysis_2.md` carries all of them if you need the raw series.")
        A("")
    else:
        A("# Climate Change Analysis 2: full workbook contents, with formulas")
        A("")
        A(f"Converted from `Climate_Change_Analysis_2.xlsx` ({stamp}) to")
        A(f"Markdown for reading by a language model. All {len(wbf.worksheets)} worksheets and "
          "all non-empty cells are")
        A("included, and every formula in the workbook is recorded.")
        A("")
    A("## How to read this file")
    A("")
    A("**Coordinates.** Each worksheet is a `## Sheet:` section, in workbook tab order, as one")
    A("Markdown table. The first table column, `row`, is the real Excel row number, and each header")
    A("cell is prefixed with its real column letter, e.g. `N: Residual`. Cross-references written")
    A("inside the workbook (`Volcanic!N39`, `Forcing!T33`, `Kernel!P62:Q70`, `Diagnostics!AR16`)")
    A("therefore resolve directly in this file. Columns that are empty across a whole sheet are")
    A("omitted, so header letters can skip (…`Q`, `S`…); row alignment is otherwise exactly as in")
    A("Excel, which matters because these sheets put a label in one column and its value a few")
    A("columns to the right. These sheets are sparse by design: a data series on the left, panels of")
    A("assumptions, regression statistics and commentary to the right, so many cells are")
    A("legitimately blank.")
    A("")
    A("**Formulas.** A cell that carries a formula shows it next to its value, in the form")
    A("`0.1512988  {=AVERAGEIFS(...)}`. The exception is a formula filled down a column for hundreds")
    A("of rows: repeating it on every row would bury the file, so each such fill is listed once in")
    A("that sheet's **Formula columns** table, which gives the column, the row range, and the")
    A("formula as written in the first row of the range (row numbers in it shift down the column in")
    A("the usual way). Cells inside a listed fill show values only. So: every formula in the")
    A("workbook is here, once per fill or once per cell.")
    A("")
    A("**Values.** Values are the ones Excel cached, printed to 10 significant figures, which")
    A("also clears binary-float noise (`0.08000000000000007` → `0.08`). Where a formula cell has")
    A("no cached value, the value shown comes instead from a LibreOffice recalculation of this")
    A("same file, checked cell by cell against the scripts listed in `Sources` section E.")
    A(_nocache_text)
    A("")
    A("**The two engines.** From v1.7.4 the workbook ships as saved by Microsoft Excel 16.0300, so")
    A("the values above are Excel's own. They were reconciled against a LibreOffice recalculation")
    A("of the same file: 120,565 numeric cells, which is every computed cell in it, with no")
    A("difference beyond 1e-9 relative, and 107,465 formulas unchanged by the round trip. The")
    A("three reference tabs added in v1.7.5 are text only and change neither figure. That")
    A("reconciliation is worth")
    A("stating because the engines did once disagree. Through v1.7.2, seven formulas in")
    A("`Diagnostics` section C passed a lone boolean array to `SUMPRODUCT`, as in")
    A("`=SUMPRODUCT(($B$9:$B$822>=2016))`. LibreOffice coerces such an array to numbers; Excel does")
    A("not, so in Excel those seven returned 0 and the eighteen cells dividing by them showed")
    A("`#DIV/0!`, from v1.1 onward, while the pasted figures, all four scripts and every Markdown")
    A("export stayed numerically right. It surfaced only when the file was opened and saved in")
    A("Excel and that cache was compared against this export. v1.7.3 rewrote the seven as")
    A("`SUMPRODUCT(--(...))`; v1.7.4 is the first export in which Excel has confirmed the fix. The")
    A("lesson generalizes: a figure confirmed by one engine is unconfirmed.")
    A("")
    A("**Conventions carried over from the workbook.** A value shown in blue in Excel was computed")
    A("outside the workbook and pasted; the four scripts responsible are listed in `Sources` section")
    A("E. Yellow cells are hand-maintained inputs. Blue and yellow do not survive into Markdown, so")
    A("use `Sources!B50:F55` and the assumption blocks to tell which is which.")
    A("")
    A("**Indentation.** Leading spaces inside a cell are preserved, because this workbook uses them")
    A("to nest sub-items under a panel heading (`Single-regressor fits`, then `  log2(CO2): °C per")
    A("doubling`). Read a leading-space cell as a child of the nearest unindented label above it.")
    A("")
    A("**Escapes.** `\\|` is a literal pipe inside a cell. A line break inside a cell is shown as")
    A("` / `. The two raw NOAA CO2 tabs keep their `#` comment header, reassembled into a code block")
    A("above the data (the original comment lines contained commas, so Excel split them across")
    A("columns).")
    A("")
    A("## Sheets in this workbook")
    A("")
    for ws in wbf.worksheets:
        A(f"- **{ws.title}**: {ws.max_row} rows x {ws.max_column} columns")
    A("")

    total_formulas = 0
    total_inline = 0
    total_runs = 0

    for ws in wbf.worksheets:
        wsv, wsr = wbv[ws.title], wbr[ws.title]
        nrow, ncol = ws.max_row, ws.max_column

        # ---- grids of formulas, values
        formula = {}
        value = {}
        for r in range(1, nrow + 1):
            for c in range(1, ncol + 1):
                fv = ws.cell(r, c).value
                if isinstance(fv, str) and fv.startswith("="):
                    formula[(r, c)] = fv[1:]
                    total_formulas += 1
                v = wsv.cell(r, c).value
                if v is None and (r, c) in formula:
                    v = wsr.cell(r, c).value
                    value[(r, c)] = (v, True)
                else:
                    value[(r, c)] = (v, False)

        # ---- find column fills
        in_run = set()
        runs = []
        for c in range(1, ncol + 1):
            r = 1
            while r <= nrow:
                if (r, c) not in formula:
                    r += 1
                    continue
                key = normalize(formula[(r, c)], r)
                r2 = r
                while (r2 + 1, c) in formula and \
                        normalize(formula[(r2 + 1, c)], r2 + 1) == key:
                    r2 += 1
                if r2 - r + 1 >= MIN_RUN:
                    runs.append((c, r, r2, formula[(r, c)]))
                    for rr in range(r, r2 + 1):
                        in_run.add((rr, c))
                r = r2 + 1
        total_runs += len(runs)

        keep = [c for c in range(1, ncol + 1)
                if any(nonempty(value[(r, c)][0]) or (r, c) in formula
                       for r in range(1, nrow + 1))]

        A("")
        A(f"## Sheet: {ws.title}")
        A("")
        A(f"Excel range {ws.dimensions}: {nrow} rows x {ncol} columns; "
          f"{len(keep)} columns contain data; {sum(1 for k in formula)} formula cells, "
          f"of which {len(in_run)} belong to {len(runs)} column fills listed below.")
        A("")
        if not keep:
            A("_(empty sheet)_")
            A("")
            continue

        # ---- formula-columns legend
        if runs:
            A("### Formula columns")
            A("")
            A("Each row is one formula filled down a column. The formula is shown as written in the")
            A("first row of the range; its relative row references shift down the column. Cells in")
            A("these ranges show values only in the table below.")
            A("")
            A("| Column | Rows | Formula (as written in the first row) |")
            A("|---|---|---|")
            for c, r1, r2, f in sorted(runs):
                mark = " [recalc]" if value[(r1, c)][1] else ""
                A(f"| {get_column_letter(c)} | {r1}-{r2}{mark} | `={esc(f)}` |")
            A("")

        # ---- build table rows
        rows = []
        for r in range(1, nrow + 1):
            seg = []
            present = False
            for c in keep:
                v, recalced = value[(r, c)]
                txt = fmt(v)
                if (r, c) in formula and (r, c) not in in_run:
                    f = esc(formula[(r, c)])
                    txt = (txt + f"  {{={f}}}").strip()
                    total_inline += 1
                if txt:
                    present = True
                seg.append(txt)
            if present:
                rows.append((r, seg))

        # ---- raw-data tabs: lift the "#" comment header out
        cmt = [(r, s) for r, s in rows if s and s[0].startswith("#")]
        if len(cmt) > 3:
            A(f"Comment header from the original downloaded data file "
              f"(sheet rows {cmt[0][0]}-{cmt[-1][0]}), reassembled:")
            A("")
            A("```")
            for r, s in cmt:
                A(", ".join(x for x in s if x).replace("\\|", "|"))
            A("```")
            A("")
            rows = [(r, s) for r, s in rows if not (s and s[0].startswith("#"))]

        # ---- audit mode: drop the repetitive input-series rows, keeping the header row
        if audit:
            fill_count = {c: sum(1 for r in range(1, nrow + 1)
                                 if nonempty(value[(r, c)][0]) or (r, c) in formula)
                          for c in keep}
            short = {c for c in keep if fill_count[c] < 100}
            allrows = [r for r, _ in rows]
            edge = set(allrows[:4]) | set(allrows[-3:])
            kept, dropped = [], 0
            for r, seg in rows:
                structural = any(
                    (nonempty(value[(r, c)][0]) or (r, c) in formula)
                    for c in keep if c in short)
                if structural or r in edge:
                    kept.append((r, seg))
                else:
                    dropped += 1
            rows = kept
            if dropped:
                A(f"Audit extract: {dropped} rows of the input series omitted here (every cell in "
                  f"them belonged to a column holding 100 or more entries). The rows below carry "
                  f"all panel, statistic, label and one-off formula content, plus the header and "
                  f"the first and last rows of the series.")
                A("")

        letters = [get_column_letter(c) for c in keep]
        hr, hdr = rows[0]
        src = [ws.cell(hr, c).value for c in keep]
        is_hdr = (any(isinstance(v, str) and v.strip() and not v.startswith("=") for v in src)
                  and not any(isinstance(v, (int, float)) and not isinstance(v, bool)
                              for v in src))
        if is_hdr:
            head = ["row"] + [f"{letters[i]}: {hdr[i]}" if hdr[i] else letters[i]
                              for i in range(len(letters))]
            body = rows[1:]
            A(f"### Cells (header row: Excel row {hr})")
        else:
            head = ["row"] + letters
            body = rows
            A("### Cells (no single header row; columns identified by letter)")
        A("")
        A("| " + " | ".join(head) + " |")
        A("|" + "|".join(["---"] * len(head)) + "|")
        for r, seg in body:
            A(f"| {r} | " + " | ".join(seg) + " |")
        A("")

    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    print(f"formula cells: {total_formulas:,}")
    print(f"  in column fills (legend): {total_formulas - total_inline:,} "
          f"across {total_runs} fills")
    print(f"  shown inline: {total_inline:,}")
    print(f"wrote {out_path}")


if __name__ == "__main__":
    main(*sys.argv[1:5])
