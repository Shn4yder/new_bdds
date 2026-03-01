from datetime import datetime, date
from pathlib import Path
from typing import List, Tuple
from report_generator.body_filler import build_body_workbook
from report_generator.footer_filler import build_footer_workbook
from report_generator.header_filler import build_header_workbook
from report_generator.report_builder import build_report_from_workbooks
from extractor import aggregate_baseline_by_month


def months_between(start: date, end: date) -> List[Tuple[int, int]]:
    start_m = date(start.year, start.month, 1)
    end_m = date(end.year, end.month, 1)
    months: List[Tuple[int, int]] = []
    y, m = start_m.year, start_m.month
    while True:
        months.append((y, m))
        if y == end_m.year and m == end_m.month:
            break
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def set_value_preserve_merge(ws, row: int, col: int, value):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            ws.cell(row=rng.min_row, column=rng.min_col, value=value)
            return
    ws.cell(row=row, column=col, value=value)


def prompt_mpp_path(prompt: str) -> Path:
    while True:
        s = input(prompt).strip().strip('"').strip("'")
        p = Path(s).resolve()
        if p.exists() and p.suffix.lower() == ".mpp":
            return p
        print("Укажите корректный путь к .mpp файлу.")

def prompt_complexity(prompt: str) -> int:
    while True:
        s = input(prompt).strip()
        if s in ("0", "1"):
            return int(s)
        print("Введите 0 для простого или 1 для сложного проекта.")

def main():

    mpp_path = prompt_mpp_path("Введите путь к .mpp файлу: ")
    complexity = prompt_complexity("Проект сложный? (1) или простой? (0): ")
    project_name, month_dates, rows = aggregate_baseline_by_month(str(mpp_path))
    start_date = month_dates[0]
    end_date = month_dates[-1]
    resource_names = [r[0] for r in rows]
    # Build maps keyed by (y,m)
    work_by_res = {
        r[0]: { (md.year, md.month): float(r[2].get(f"{md.year:04d}-{md.month:02d}", 0.0)) for md in month_dates }
        for r in rows
    }
    cost_by_res = {
        r[0]: { (md.year, md.month): float(r[3].get(f"{md.year:04d}-{md.month:02d}", 0.0)) for md in month_dates }
        for r in rows
    }
    if end_date < start_date:
        raise ValueError("Дата окончания раньше даты начала")

    months = months_between(start_date, end_date)

    body_template_path = Path("templates") / "body.xlsx"
    body_template_path = body_template_path.resolve()
    if not body_template_path.exists():
        raise FileNotFoundError(f"Не найден шаблон: {body_template_path}")

    start_str = start_date.strftime("%d.%m.%Y")
    end_str = end_date.strftime("%d.%m.%Y")
    report_dir = body_template_path.parent

    body_wb = build_body_workbook(body_template_path, months, resource_names)

    footer_template_path = (Path("templates") / "footer.xlsx").resolve()
    if not footer_template_path.exists():
        raise FileNotFoundError(f"Не найден шаблон подвала: {footer_template_path}")
    footer_wb = build_footer_workbook(footer_template_path, months)

    header_template_path = (Path("templates") / "header.xlsx").resolve()
    if not header_template_path.exists():
        raise FileNotFoundError(f"Не найден шаблон хедера: {header_template_path}")
    header_wb = build_header_workbook(header_template_path, months)
    ws_header = header_wb.active
    set_value_preserve_merge(ws_header, 6, 2, project_name)
    ws_body = body_wb.active
    base_col_idx = 4
    last_col = 4 + len(months)
    for idx, res in enumerate(resource_names):
        row_work = 3 + idx * 2
        row_cost = row_work + 1
        row_work_sum = 0.0
        row_cost_sum = 0.0
        for j, (y, m) in enumerate(months):
            val_work = float(work_by_res.get(res, {}).get((y, m), 0.0))
            val_cost = float(cost_by_res.get(res, {}).get((y, m), 0.0))
            ws_body.cell(row=row_work, column=base_col_idx + j, value=val_work)
            ws_body.cell(row=row_cost, column=base_col_idx + j, value=val_cost)
            row_work_sum += val_work
            row_cost_sum += val_cost
        ws_body.cell(row=row_work, column=last_col, value=row_work_sum)
        ws_body.cell(row=row_cost, column=last_col, value=row_cost_sum)
    ws_footer = footer_wb.active
    sum_work_total = 0.0
    for j, (y, m) in enumerate(months):
        col = base_col_idx + j
        month_sum = 0.0
        for res in resource_names:
            month_sum += float(work_by_res.get(res, {}).get((y, m), 0.0))
        ws_footer.cell(row=1, column=col, value=month_sum)
        sum_work_total += month_sum
    ws_footer.cell(row=1, column=last_col, value=sum_work_total)
    sum_cost_total = 0.0
    for j, (y, m) in enumerate(months):
        col = base_col_idx + j
        month_sum = 0.0
        for res in resource_names:
            month_sum += float(cost_by_res.get(res, {}).get((y, m), 0.0))
        ws_footer.cell(row=2, column=col, value=month_sum)
        sum_cost_total += month_sum
    ws_footer.cell(row=2, column=last_col, value=sum_cost_total)
    mgmt_val = 20 if complexity == 1 else 30
    mgmt_sum_total = 0.0
    for j, (_y, _m) in enumerate(months):
        col = base_col_idx + j
        ws_footer.cell(row=3, column=col, value=mgmt_val)
        mgmt_sum_total += mgmt_val
    ws_footer.cell(row=3, column=last_col, value=mgmt_sum_total)
    report_output_path = report_dir / f"report_{start_str}_{end_str}.xlsx"
    build_report_from_workbooks(header_wb, body_wb, footer_wb, report_output_path, months)
    print(f"Готово. Отчёт сохранён: {report_output_path}")


if __name__ == "__main__":
    main()
