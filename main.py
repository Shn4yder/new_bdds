from datetime import datetime, date
from pathlib import Path
from typing import List, Tuple
from report_generator.body_filler import build_body_workbook
from report_generator.footer_filler import build_footer_workbook
from report_generator.header_filler import build_header_workbook
from report_generator.report_builder import build_report_from_workbooks


def parse_date_ru(value: str) -> date:
    return datetime.strptime(value.strip(), "%d.%m.%Y").date()


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


def prompt_date_ru(prompt: str) -> date:
    while True:
        s = input(prompt).strip()
        try:
            return parse_date_ru(s)
        except ValueError:
            print("Неверный формат даты. Ожидается ДД.ММ.ГГГГ. Попробуйте снова.")


def main():

    start_date = prompt_date_ru("Введите дату начала (ДД.ММ.ГГГГ): ")
    while True:
        end_date = prompt_date_ru("Введите дату окончания (ДД.ММ.ГГГГ): ")
        if end_date < start_date:
            print("Дата окончания раньше даты начала. Введите корректную дату окончания.")
        else:
            break

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

    body_wb = build_body_workbook(body_template_path, months, (Path("templates") / "resources.txt").resolve())

    footer_template_path = (Path("templates") / "footer.xlsx").resolve()
    if not footer_template_path.exists():
        raise FileNotFoundError(f"Не найден шаблон подвала: {footer_template_path}")
    footer_wb = build_footer_workbook(footer_template_path, months)

    header_template_path = (Path("templates") / "header.xlsx").resolve()
    if not header_template_path.exists():
        raise FileNotFoundError(f"Не найден шаблон хедера: {header_template_path}")
    header_wb = build_header_workbook(header_template_path, months)
    report_output_path = report_dir / f"report_{start_str}_{end_str}.xlsx"
    build_report_from_workbooks(header_wb, body_wb, footer_wb, report_output_path, months)
    print(f"Готово. Отчёт сохранён: {report_output_path}")


if __name__ == "__main__":
    main()
