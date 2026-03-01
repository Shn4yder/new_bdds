from pathlib import Path
from typing import List, Tuple
from report_generator.utils import (
    ensure_openpyxl,
    copy_cell_style,
    load_workbook,
    Workbook,
    get_column_letter,
)


def _autosize_columns(ws, cols: List[int]):
    for col in cols:
        max_len = 0
        for row in ws.iter_rows(min_col=col, max_col=col, values_only=True):
            v = row[0]
            if v is None:
                continue
            s = str(v)
            if len(s) > max_len:
                max_len = len(s)
        letter = get_column_letter(col)
        if max_len > 0:
            ws.column_dimensions[letter].width = max(8, min(60, max_len * 1.2))
        else:
            ws.column_dimensions[letter].width = 8


def _copy_sheet_range(src_ws, dst_ws, start_row_dst: int, last_col: int):
    max_row_src = src_ws.max_row
    for r in range(1, max_row_src + 1):
        dst_r = start_row_dst + r - 1
        dst_ws.row_dimensions[dst_r].height = src_ws.row_dimensions[r].height
        for c in range(1, last_col + 1):
            src = src_ws.cell(row=r, column=c)
            dst = dst_ws.cell(row=dst_r, column=c)
            dst.value = src.value
            copy_cell_style(src, dst)
    for rng in list(src_ws.merged_cells.ranges):
        sr, sc = rng.min_row, rng.min_col
        er, ec = rng.max_row, rng.max_col
        sr += start_row_dst - 1
        er += start_row_dst - 1
        if sc > last_col:
            continue
        if ec > last_col:
            ec = last_col
        dst_ws.merge_cells(start_row=sr, start_column=sc, end_row=er, end_column=ec)


def build_report(header_path: Path, body_path: Path, footer_path: Path, output_path: Path, months: List[Tuple[int, int]]):
    ensure_openpyxl()
    last_col = 4 + len(months)
    wb_header = load_workbook(filename=str(header_path), data_only=False)
    wb_body = load_workbook(filename=str(body_path), data_only=False)
    wb_footer = load_workbook(filename=str(footer_path), data_only=False)
    ws_header = wb_header.active
    ws_body = wb_body.active
    ws_footer = wb_footer.active

    wb_out = Workbook()
    ws_out = wb_out.active

    for c in range(1, last_col + 1):
        letter = get_column_letter(c)
        dim = ws_header.column_dimensions.get(letter)
        if dim is not None and dim.width is not None:
            ws_out.column_dimensions[letter].width = dim.width

    _copy_sheet_range(ws_header, ws_out, 1, last_col)
    offset = ws_header.max_row + 1
    _copy_sheet_range(ws_body, ws_out, offset, last_col)
    offset = offset + ws_body.max_row
    _copy_sheet_range(ws_footer, ws_out, offset, last_col)
    _autosize_columns(ws_out, [2, 3])
    ws_out.column_dimensions["B"].width = 60

    wb_out.save(str(output_path))


def build_report_from_workbooks(header_wb, body_wb, footer_wb, output_path: Path, months: List[Tuple[int, int]]):
    ensure_openpyxl()
    last_col = 4 + len(months)
    ws_header = header_wb.active
    ws_body = body_wb.active
    ws_footer = footer_wb.active
    wb_out = Workbook()
    ws_out = wb_out.active
    for c in range(1, last_col + 1):
        letter = get_column_letter(c)
        dim = ws_header.column_dimensions.get(letter)
        if dim is not None and dim.width is not None:
            ws_out.column_dimensions[letter].width = dim.width
    _copy_sheet_range(ws_header, ws_out, 1, last_col)
    offset = ws_header.max_row + 1
    _copy_sheet_range(ws_body, ws_out, offset, last_col)
    offset = offset + ws_body.max_row
    _copy_sheet_range(ws_footer, ws_out, offset, last_col)
    _autosize_columns(ws_out, [2, 3])
    ws_out.column_dimensions["B"].width = 60
    wb_out.save(str(output_path))
