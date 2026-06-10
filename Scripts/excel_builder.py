"""
Excel Output Builder
Takes MUS results and writes the formatted workpaper Excel file.

Tabs: Cover, Population Reconciliation, Exclusions, Positive Population,
Negative Population, Positive Sample, Negative Sample.

Value parsing uses the single shared clean_value from mus_engine so the workpaper
cells always match the engine result.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

from mus_engine import clean_value


# ── Colours ───────────────────────────────────────────────────────────────────
BLUE_FILL    = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
LIGHT_FILL   = PatternFill("solid", start_color="D6E4F0", end_color="D6E4F0")
IS_FILL      = PatternFill("solid", start_color="FCE4D6", end_color="FCE4D6")
MUS_FILL     = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
GREY_FILL    = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
WHITE_FILL   = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
YELLOW_FILL  = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")
EXCL_FILL    = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")

WHITE_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BOLD_FONT    = Font(name="Arial", bold=True, size=10)
NORMAL_FONT  = Font(name="Arial", size=10)
LABEL_FONT   = Font(name="Arial", bold=True, size=10, color="1F4E79")

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def _auto_width(ws, min_width=10, max_width=50):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)


def build_cover_tab(wb, pos_workings, neg_workings, pm, risk_factor, filename,
                    popnum_col, value_col, decimal_system):
    ws = wb.create_sheet("Cover", 0)
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:D1")
    title = ws["A1"]
    title.value = "Monetary Unit Sampling \u2013 Audit Workpaper"
    title.font = Font(name="Arial", bold=True, size=14, color="1F4E79")
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:D2")
    subtitle = ws["A2"]
    subtitle.value = f"Generated: {datetime.now().strftime('%d %B %Y %H:%M')}"
    subtitle.font = Font(name="Arial", size=10, color="808080")
    subtitle.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 16

    def section(row, label):
        ws.merge_cells(f"A{row}:D{row}")
        c = ws[f"A{row}"]
        c.value = label
        c.font = WHITE_FONT
        c.fill = BLUE_FILL
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 18

    def row_pair(row, label, value, fmt=None):
        ws[f"A{row}"].value = label
        ws[f"A{row}"].font = NORMAL_FONT
        ws[f"A{row}"].fill = LIGHT_FILL
        ws[f"A{row}"].alignment = Alignment(vertical="center", indent=1)
        ws[f"A{row}"].border = BORDER
        ws.merge_cells(f"A{row}:C{row}")

        c = ws[f"D{row}"]
        c.value = value
        c.font = BOLD_FONT
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = BORDER
        if fmt and not isinstance(value, str):
            c.number_format = fmt
        ws.row_dimensions[row].height = 16

    # Inputs section
    section(4, "Sampling Inputs")
    row_pair(5, "Source file", filename)
    row_pair(6, "Population number column", popnum_col)
    row_pair(7, "Value column", value_col)
    row_pair(8, "Decimal system", "EUR (1.234,56)" if decimal_system == "eur" else "Normal (1,234.56)")
    row_pair(9, "Performance materiality", pm, '#,##0.00')
    row_pair(10, "Risk factor", risk_factor, '0.0000')
    row_pair(11, "Random seed (fixed)", 42)
    row_pair(12, "Sample size rounding", "Always round up (ceiling)")

    def population_block(start_row, w, label, abs_note=""):
        section(start_row, label)
        row_pair(start_row + 1, "Total items in population", w["total_count"])
        row_pair(start_row + 2, f"Total population value{abs_note}", w["total_value"], '#,##0.00')
        row_pair(start_row + 3, f"Individually significant items ({abs_note or '>= PM'}) \u2013 count", w["is_count"])
        row_pair(start_row + 4, "Individually significant items \u2013 value", w["is_value"], '#,##0.00')
        row_pair(start_row + 5, "Residual population \u2013 count", w["remaining_count"])
        row_pair(start_row + 6, "Residual population \u2013 value", w["remaining_total"], '#,##0.00')
        row_pair(start_row + 7, "Raw sample size (Residual Total \u00d7 Risk Factor \u00f7 PM)", w["raw_sample_size"], '0.0000')
        row_pair(start_row + 8, "Sample size (rounded up)", w["sample_size"])
        if w["full_residual_selected"]:
            row_pair(start_row + 9, "Sampling interval", "N/A - full residual selected")
            row_pair(start_row + 10, "Random start (seed 42)", "N/A - full residual selected")
        else:
            row_pair(start_row + 9, "Sampling interval", w["interval"], '#,##0.00')
            row_pair(start_row + 10, "Random start (seed 42)", w["random_start"], '#,##0.00')
        row_pair(start_row + 11, "MUS items selected", w["mus_selected_count"])
        row_pair(start_row + 12, "Total items selected (IS + MUS)", w["total_selected"])
        return start_row + 13

    next_row = 14
    if pos_workings:
        next_row = population_block(next_row, pos_workings, "Positive Population Summary") + 1
    if neg_workings:
        population_block(next_row, neg_workings, "Negative Population Summary", abs_note=" (absolute)")

    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 5
    ws.column_dimensions["D"].width = 24


def build_reconciliation_tab(wb, reconciliation):
    """Population completeness reconciliation: source = pos + neg + zero + skipped."""
    ws = wb.create_sheet("Population Reconciliation")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:C1")
    t = ws["A1"]
    t.value = "Population Reconciliation"
    t.font = Font(name="Arial", bold=True, size=14, color="1F4E79")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:C2")
    n = ws["A2"]
    n.value = "Source rows must equal Positives + Negatives + Zeros + Excluded."
    n.font = Font(name="Arial", italic=True, size=9, color="595959")
    n.alignment = Alignment(horizontal="center")

    headers = ["Category", "Count", "Absolute Value"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = WHITE_FONT
        c.fill = BLUE_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER

    r = reconciliation
    rows = [
        ("Positives", r["pos_count"], r["pos_value"]),
        ("Negatives", r["neg_count"], r["neg_value"]),
        ("Zeros (excluded)", r["zero_count"], r["zero_value"]),
        ("Unparseable / blank (excluded)", r["skipped"], None),
    ]
    row_i = 5
    first_data = row_i
    for label, count, val in rows:
        ws.cell(row=row_i, column=1, value=label).border = BORDER
        ws.cell(row=row_i, column=1).font = NORMAL_FONT
        cc = ws.cell(row=row_i, column=2, value=count)
        cc.border = BORDER
        cc.font = NORMAL_FONT
        cc.alignment = Alignment(horizontal="right")
        vc = ws.cell(row=row_i, column=3, value=val if val is not None else "")
        vc.border = BORDER
        vc.font = NORMAL_FONT
        vc.alignment = Alignment(horizontal="right")
        if val is not None:
            vc.number_format = '#,##0.00'
        row_i += 1
    last_data = row_i - 1

    # Computed total row (live formula)
    tc_label = ws.cell(row=row_i, column=1, value="Computed total (sum of above)")
    tc_label.font = BOLD_FONT
    tc_label.fill = LIGHT_FILL
    tc_label.border = BORDER
    tc_count = ws.cell(row=row_i, column=2, value=f"=SUM(B{first_data}:B{last_data})")
    tc_count.font = BOLD_FONT
    tc_count.fill = LIGHT_FILL
    tc_count.border = BORDER
    tc_count.alignment = Alignment(horizontal="right")
    computed_row = row_i
    row_i += 1

    # Source rows from file
    src_label = ws.cell(row=row_i, column=1, value="Source rows in file")
    src_label.font = BOLD_FONT
    src_label.fill = LIGHT_FILL
    src_label.border = BORDER
    src_count = ws.cell(row=row_i, column=2, value=r["source_rows"])
    src_count.font = BOLD_FONT
    src_count.fill = LIGHT_FILL
    src_count.border = BORDER
    src_count.alignment = Alignment(horizontal="right")
    source_row = row_i
    row_i += 1

    # Reconciliation check (live)
    chk_label = ws.cell(row=row_i, column=1, value="Reconciliation check")
    chk_label.font = BOLD_FONT
    chk_label.border = BORDER
    chk = ws.cell(row=row_i, column=2, value=f'=IF(B{computed_row}=B{source_row},"OK","CHECK")')
    chk.font = Font(name="Arial", bold=True, size=10)
    chk.border = BORDER
    chk.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 20


def build_exclusions_tab(wb, exclusions):
    """Lists every excluded row (zeros and unparseable/blank) with the reason."""
    ws = wb.create_sheet("Exclusions")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value = "Exclusions \u2013 rows not sampled"
    t.font = Font(name="Arial", bold=True, size=14, color="1F4E79")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    if not exclusions:
        ws["A3"] = "No exclusions. Every row was loaded into a population."
        ws["A3"].font = NORMAL_FONT
        ws.column_dimensions["A"].width = 50
        return

    headers = ["Source Row", "Population Number", "Raw Value", "Type", "Reason"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = WHITE_FONT
        c.fill = BLUE_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER

    for offset, ex in enumerate(exclusions):
        r = 4 + offset
        cells = [
            ex.get("row", ""),
            ex.get("popnum", ""),
            "" if ex.get("raw", "") is None else str(ex.get("raw", "")),
            ex.get("type", ""),
            ex.get("reason", ""),
        ]
        for i, v in enumerate(cells, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = NORMAL_FONT
            c.fill = EXCL_FILL
            c.border = BORDER
            c.alignment = Alignment(vertical="center")

    ws.freeze_panes = "A4"
    _auto_width(ws)


def build_population_tab(wb, tab_name, annotated, original_headers, value_col, decimal_system):
    ws = wb.create_sheet(tab_name)
    ws.sheet_view.showGridLines = False

    if not annotated:
        ws["A1"] = "No transactions in this population."
        return

    extra_headers = ["Selection Status", "Running Total (Residual Pop.)"]
    all_headers = original_headers + extra_headers

    for col_idx, h in enumerate(all_headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = WHITE_FONT
        cell.fill = BLUE_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 30

    legend_row = 2
    ws.merge_cells(f"A{legend_row}:{get_column_letter(len(all_headers))}{legend_row}")
    legend_cell = ws[f"A{legend_row}"]
    legend_cell.value = (
        "IS = Individually Significant (|value| >= Performance Materiality, 100% selected)   "
        "MUS = Selected by Monetary Unit Sampling   "
        "Running total applies to residual population only (IS items excluded from MUS walk)"
    )
    legend_cell.font = Font(name="Arial", italic=True, size=9, color="595959")
    legend_cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[legend_row].height = 28

    data_start_row = 3
    status_col_idx = len(original_headers) + 1
    running_col_idx = len(original_headers) + 2
    value_col_idx = original_headers.index(value_col) + 1

    running_col_letter = get_column_letter(running_col_idx)
    value_col_letter = get_column_letter(value_col_idx)
    status_col_letter = get_column_letter(status_col_idx)

    for row_offset, item in enumerate(annotated):
        excel_row = data_start_row + row_offset
        t = item["transaction"]
        status = item["status"]

        if status == "IS":
            row_fill = IS_FILL
        elif status == "MUS":
            row_fill = MUS_FILL
        elif row_offset % 2 == 0:
            row_fill = WHITE_FILL
        else:
            row_fill = GREY_FILL

        for col_idx, h in enumerate(original_headers, start=1):
            val = t.original_row.get(h, "")
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.font = NORMAL_FONT
            cell.fill = row_fill
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")

            if h == value_col:
                parsed = clean_value(val, decimal_system)
                if parsed is not None:
                    cell.value = parsed
                    cell.number_format = '#,##0.00'

        status_cell = ws.cell(row=excel_row, column=status_col_idx)
        status_cell.value = status if status else ""
        status_cell.font = Font(name="Arial", bold=True, size=10)
        status_cell.fill = row_fill
        status_cell.border = BORDER
        status_cell.alignment = Alignment(horizontal="center", vertical="center")

        rt_cell = ws.cell(row=excel_row, column=running_col_idx)
        rt_cell.border = BORDER
        rt_cell.fill = row_fill
        rt_cell.alignment = Alignment(horizontal="right", vertical="center")
        rt_cell.font = NORMAL_FONT

        if status == "IS":
            rt_cell.value = "N/A - IS"
            rt_cell.font = Font(name="Arial", italic=True, size=10, color="808080")
        else:
            rt_cell.value = (
                f'=SUMPRODUCT(({status_col_letter}${data_start_row}:'
                f'{status_col_letter}{excel_row}<>"IS")*'
                f'ABS({value_col_letter}${data_start_row}:{value_col_letter}{excel_row}))'
            )
            rt_cell.number_format = '#,##0.00'

    ws.freeze_panes = f"A{data_start_row}"
    _auto_width(ws)
    ws.column_dimensions[running_col_letter].width = 30


def build_sample_tab(wb, tab_name, annotated, original_headers, value_col, decimal_system):
    ws = wb.create_sheet(tab_name)
    ws.sheet_view.showGridLines = False

    selected = [item for item in annotated if item["status"] in ("IS", "MUS")]

    if not selected:
        ws["A1"] = "No items selected in this population."
        return

    all_headers = original_headers + ["Selection Status"]

    for col_idx, h in enumerate(all_headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = WHITE_FONT
        cell.fill = BLUE_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 30

    value_col_idx = original_headers.index(value_col) + 1
    status_col_idx = len(original_headers) + 1

    for row_offset, item in enumerate(selected):
        excel_row = 2 + row_offset
        t = item["transaction"]
        status = item["status"]
        row_fill = IS_FILL if status == "IS" else MUS_FILL

        for col_idx, h in enumerate(original_headers, start=1):
            val = t.original_row.get(h, "")
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.font = NORMAL_FONT
            cell.fill = row_fill
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")
            if h == value_col:
                parsed = clean_value(val, decimal_system)
                if parsed is not None:
                    cell.value = parsed
                    cell.number_format = '#,##0.00'

        sc = ws.cell(row=excel_row, column=status_col_idx)
        sc.value = status
        sc.font = Font(name="Arial", bold=True, size=10)
        sc.fill = row_fill
        sc.border = BORDER
        sc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[excel_row].height = 16

    total_row = 2 + len(selected)
    # Label goes one column left of the value column, unless value is column 1
    label_col = value_col_idx - 1 if value_col_idx > 1 else value_col_idx + 1
    lc = ws.cell(row=total_row, column=label_col)
    lc.value = "TOTAL SELECTED"
    lc.font = BOLD_FONT
    lc.fill = LIGHT_FILL
    lc.border = BORDER

    val_col_letter = get_column_letter(value_col_idx)
    tc = ws.cell(row=total_row, column=value_col_idx)
    tc.value = f"=SUM({val_col_letter}2:{val_col_letter}{total_row - 1})"
    tc.font = BOLD_FONT
    tc.fill = LIGHT_FILL
    tc.border = BORDER
    tc.number_format = '#,##0.00'

    ws.freeze_panes = "A2"
    _auto_width(ws)


def build_workbook(
    output_path,
    pos_results,
    neg_results,
    original_headers,
    popnum_col,
    value_col,
    pm,
    risk_factor,
    filename,
    decimal_system,
    reconciliation,
    exclusions,
):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    pos_workings = pos_results["workings"] if pos_results else None
    neg_workings = neg_results["workings"] if neg_results else None

    # Tab 1: Cover
    build_cover_tab(wb, pos_workings, neg_workings, pm, risk_factor, filename,
                    popnum_col, value_col, decimal_system)

    # Tab 2: Population Reconciliation
    build_reconciliation_tab(wb, reconciliation)

    # Tab 3: Exclusions
    build_exclusions_tab(wb, exclusions)

    # Tab 4: Positive population
    if pos_results and pos_results["annotated"]:
        build_population_tab(wb, "Positive Population", pos_results["annotated"],
                             original_headers, value_col, decimal_system)
    else:
        ws = wb.create_sheet("Positive Population")
        ws["A1"] = "No positive transactions in population."

    # Tab 5: Negative population
    if neg_results and neg_results["annotated"]:
        build_population_tab(wb, "Negative Population", neg_results["annotated"],
                             original_headers, value_col, decimal_system)
    else:
        ws = wb.create_sheet("Negative Population")
        ws["A1"] = "No negative transactions in population."

    # Tab 6: Positive sample
    if pos_results and pos_results["annotated"]:
        build_sample_tab(wb, "Positive Sample", pos_results["annotated"],
                         original_headers, value_col, decimal_system)
    else:
        ws = wb.create_sheet("Positive Sample")
        ws["A1"] = "No positive transactions selected."

    # Tab 7: Negative sample
    if neg_results and neg_results["annotated"]:
        build_sample_tab(wb, "Negative Sample", neg_results["annotated"],
                         original_headers, value_col, decimal_system)
    else:
        ws = wb.create_sheet("Negative Sample")
        ws["A1"] = "No negative transactions selected."

    wb.save(output_path)
    print(f"Workbook saved: {output_path}")
