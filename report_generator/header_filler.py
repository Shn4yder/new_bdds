from pathlib import Path
from typing import List, Tuple
from copy import copy as shallow_copy
from report_generator.utils import ensure_openpyxl, copy_cell_style, load_workbook, get_column_letter, Alignment


def process_header_template(template_path: Path, output_path: Path, months: List[Tuple[int, int]]):
    ensure_openpyxl()
    wb = load_workbook(filename=str(template_path), data_only=False)
    ws = wb.active

    base_col_idx = 4
    insert_count = max(0, len(months) - 1)

    src_col_L = 12
    saved_values = []
    saved_styles = []
    for r in range(1, 5):
        cell = ws.cell(row=r, column=src_col_L)
        saved_values.append(cell.value)
        saved_styles.append(
            {
                "font": shallow_copy(cell.font),
                "fill": shallow_copy(cell.fill),
                "border": shallow_copy(cell.border),
                "alignment": shallow_copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": shallow_copy(cell.protection),
            }
        )

    if insert_count > 0:
        ws.insert_cols(base_col_idx + 1, amount=insert_count)

    width_d = ws.column_dimensions["D"].width
    max_row = ws.max_row
    for i in range(1, insert_count + 1):
        col_letter = get_column_letter(base_col_idx + i)
        ws.column_dimensions[col_letter].width = width_d
        for r in range(1, max_row + 1):
            src = ws.cell(row=r, column=base_col_idx)
            dst = ws.cell(row=r, column=base_col_idx + i)
            copy_cell_style(src, dst)

    last_col = 4 + len(months)
    shifted_L_col = src_col_L + insert_count
    for r in range(1, 5):
        dst = ws.cell(row=r, column=last_col)
        dst.value = saved_values[r - 1]
        style = saved_styles[r - 1]
        dst.font = style["font"]
        dst.fill = style["fill"]
        dst.border = style["border"]
        dst.alignment = style["alignment"]
        dst.number_format = style["number_format"]
        dst.protection = style["protection"]
        ws.cell(row=r, column=shifted_L_col).value = None

    for row in (5, 6):
        ranges = list(ws.merged_cells.ranges)
        for rng in ranges:
            if rng.min_row <= row <= rng.max_row:
                ws.unmerge_cells(str(rng))
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
        if Alignment is not None:
            ws.cell(row=row, column=2).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = None

    wb.save(str(output_path))


def build_header_workbook(template_path: Path, months: List[Tuple[int, int]]):
    ensure_openpyxl()
    wb = load_workbook(filename=str(template_path), data_only=False)
    ws = wb.active

    base_col_idx = 4
    insert_count = max(0, len(months) - 1)

    src_col_L = 12
    saved_values = []
    saved_styles = []
    for r in range(1, 5):
        cell = ws.cell(row=r, column=src_col_L)
        saved_values.append(cell.value)
        saved_styles.append(
            {
                "font": shallow_copy(cell.font),
                "fill": shallow_copy(cell.fill),
                "border": shallow_copy(cell.border),
                "alignment": shallow_copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": shallow_copy(cell.protection),
            }
        )

    if insert_count > 0:
        ws.insert_cols(base_col_idx + 1, amount=insert_count)

    width_d = ws.column_dimensions["D"].width
    max_row = ws.max_row
    for i in range(1, insert_count + 1):
        col_letter = get_column_letter(base_col_idx + i)
        ws.column_dimensions[col_letter].width = width_d
        for r in range(1, max_row + 1):
            src = ws.cell(row=r, column=base_col_idx)
            dst = ws.cell(row=r, column=base_col_idx + i)
            copy_cell_style(src, dst)

    last_col = 4 + len(months)
    shifted_L_col = src_col_L + insert_count
    for r in range(1, 5):
        dst = ws.cell(row=r, column=last_col)
        dst.value = saved_values[r - 1]
        style = saved_styles[r - 1]
        dst.font = style["font"]
        dst.fill = style["fill"]
        dst.border = style["border"]
        dst.alignment = style["alignment"]
        dst.number_format = style["number_format"]
        dst.protection = style["protection"]
        ws.cell(row=r, column=shifted_L_col).value = None

    for row in (5, 6):
        ranges = list(ws.merged_cells.ranges)
        for rng in ranges:
            if rng.min_row <= row <= rng.max_row:
                ws.unmerge_cells(str(rng))
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
        if Alignment is not None:
            ws.cell(row=row, column=2).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = None

    return wb
