import argparse
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

from domrf import DomRfClient

REGION = 73
BANDS = ("До 25%", "25–75%", "Более 75%")
SUMMARY_HEADER_ROW = 2
type DeveloperRow = dict
type MonthlyRows = dict[str, list[DeveloperRow]]
type RowsByYear = dict[str, MonthlyRows]
type GroupTotals = tuple[list[float], list[float]]

DETAIL_HEADINGS = (
    "Девелопер",
    "Площадь, м²",
    "Распроданность",
    "Стройготовность",
    *(f"Продано, м² · {band.lower()}" for band in BANDS),
)


def workbook_data(months: int) -> tuple[list[str], RowsByYear]:
    periods = []
    by_year: RowsByYear = defaultdict(dict)
    with DomRfClient() as client:
        for period in client.report_periods()[:months]:
            label = f"{period.year}-{period.month:02d}"
            periods.append(label)
            for year in client.completion_years(period, REGION):
                bucket = f"{year}+" if year == period.year + 5 else str(year)
                by_year[bucket][label] = client.developers(
                    period, REGION, completion_years=[year]
                )
    if not periods or not by_year:
        raise RuntimeError("API не вернул данные для книги")
    return periods, dict(sorted(by_year.items()))


def band_for(readiness: float) -> int:
    if readiness < 25:
        return 0
    if readiness <= 75:
        return 1
    return 2


def group_totals(rows: list[DeveloperRow]) -> GroupTotals:
    areas = [0.0] * len(BANDS)
    sold = [0.0] * len(BANDS)
    for entry in rows:
        area, sale, readiness = (
            entry.get(key) for key in ("squareSum", "soldPerc", "readyPerc")
        )
        if area is None or sale is None or readiness is None:
            continue
        band = band_for(readiness)
        areas[band] += area
        sold[band] += area * sale / 100
    return sold, areas


@dataclass(frozen=True)
class SheetFormats:
    header: Format
    selector: Format
    percent: Format
    area: Format


def sheet_formats(book: xlsxwriter.Workbook) -> SheetFormats:
    header = book.add_format(
        {
            "bold": True,
            "bg_color": "#143849",
            "font_color": "white",
            "text_wrap": True,
            "valign": "vcenter",
        }
    )
    return SheetFormats(
        header,
        book.add_format({"bold": True, "bg_color": "#D8ECD7"}),
        book.add_format({"num_format": "0.00%"}),
        book.add_format({"num_format": "#,##0.00"}),
    )


def prepare_data_sheet(
    book: xlsxwriter.Workbook,
    data: Worksheet,
    periods: list[str],
    cutoff_years: list[int],
    formats: SheetFormats,
) -> None:
    data.hide()
    data.write_row(0, 0, ["Период", *DETAIL_HEADINGS[:4]], formats.header)
    data.write_column(0, 8, periods)
    data.write_column(0, 9, cutoff_years)
    book.define_name("ReportPeriods", f"='Данные'!$I$1:$I${len(periods)}")
    book.define_name("CutoffYears", f"='Данные'!$J$1:$J${len(cutoff_years)}")


def write_source_rows(data: Worksheet, monthly_rows: MonthlyRows, data_row: int) -> int:
    for period, entries in monthly_rows.items():
        for entry in entries:
            data.write_row(data_row, 0, [period, entry["key"]])
            for col, key in enumerate(("squareSum", "soldPerc", "readyPerc"), 2):
                value = entry.get(key)
                if value is None:
                    data.write_string(data_row, col, "")
                else:
                    data.write_number(data_row, col, value if col == 2 else value / 100)
            data_row += 1
    return data_row


def write_year_rows(
    sheet: Worksheet,
    monthly_rows: MonthlyRows,
    data: Worksheet,
    data_row: int,
    latest: list[DeveloperRow],
    area: Format,
) -> tuple[int, int]:
    first = data_row + 1
    data_row = write_source_rows(data, monthly_rows, data_row)
    if data_row >= first:
        sheet.write_dynamic_array_formula(
            "A5",
            f"=FILTER(Данные!$B${first}:$E${data_row},"
            f'Данные!$A${first}:$A${data_row}=ReportPeriod,"")',
            None,
            latest[0]["key"] if latest else "",
        )

    end = 4 + max(1, max(map(len, monthly_rows.values())))
    for n in range(5, end + 1):
        conditions = (
            f"D{n}<25%",
            f"AND(D{n}>=25%,D{n}<=75%)",
            f"D{n}>75%",
        )
        for col, condition in enumerate(conditions, 4):
            sheet.write_formula(
                n - 1,
                col,
                f'=IF(AND(COUNT(B{n}:D{n})=3,{condition}),B{n}*C{n},"")',
                area,
                "",
            )
    return data_row, end


def write_year_totals(
    sheet: Worksheet,
    data: Worksheet,
    index_row: int,
    year: str,
    end: int,
    latest: list[DeveloperRow],
    formats: SheetFormats,
) -> GroupTotals:
    header, area, pct = formats.header, formats.area, formats.percent
    sold_row, area_row, percent_row = end + 2, end + 3, end + 4
    sheet.write(f"A{sold_row}", "Всего продано, м²", header)
    sheet.write(f"A{area_row}", "Всего в строительстве, м²", header)
    sheet.write(f"A{percent_row}", "Распроданность", header)
    sold_totals, area_totals = group_totals(latest)
    all_area = sum((entry.get("squareSum") or 0) for entry in latest)
    sheet.write_formula(f"B{area_row}", f"=SUM(B5:B{end})", area, all_area)
    for index, letter in enumerate("EFG"):
        group = f"{letter}5:{letter}{end}"
        sold = f"{letter}{sold_row}"
        total = f"{letter}{area_row}"
        cached = sold_totals[index] / area_totals[index] if area_totals[index] else ""
        sheet.write_formula(sold, f"=SUM({group})", area, sold_totals[index])
        sheet.write_formula(
            total,
            f"=SUMPRODUCT(--ISNUMBER({group}),B5:B{end})",
            area,
            area_totals[index],
        )
        sheet.write_formula(
            f"{letter}{percent_row}", f'=IF({total}=0,"",{sold}/{total})', pct, cached
        )
        data.write_formula(
            index_row, 11 + index, f"='{year}'!{sold}", None, sold_totals[index]
        )
        data.write_formula(
            index_row, 14 + index, f"='{year}'!{total}", None, area_totals[index]
        )
    return sold_totals, area_totals


def write_year_sheet(
    sheet: Worksheet,
    data: Worksheet,
    index_row: int,
    data_row: int,
    year: str,
    monthly_rows: MonthlyRows,
    latest_period: str,
    formats: SheetFormats,
) -> tuple[int, GroupTotals]:
    header, area, pct, selector = (
        formats.header,
        formats.area,
        formats.percent,
        formats.selector,
    )
    data.write_number(index_row, 10, int(year.rstrip("+")))
    present = ",".join(f'ReportPeriod="{month}"' for month in monthly_rows)
    data.write_formula(
        index_row, 17, f"=--OR({present})", None, int(latest_period in monthly_rows)
    )
    sheet.set_column("A:A", 34)
    sheet.set_column("B:B", 17, area)
    sheet.set_column("C:D", 17, pct)
    sheet.set_column("E:G", 24, area)
    sheet.write("A1", f"Ульяновская область · {year}", header)
    sheet.write("A2", "Отчётный период")
    sheet.write_formula("B2", "=ReportPeriod", selector, latest_period)
    sheet.write_row("A4", DETAIL_HEADINGS, header)
    sheet.set_row(3, 32)
    sheet.freeze_panes(4, 0)
    latest = monthly_rows.get(latest_period, [])
    data_row, end = write_year_rows(sheet, monthly_rows, data, data_row, latest, area)
    totals = write_year_totals(sheet, data, index_row, year, end, latest, formats)
    return data_row, totals


def write_cutoff_sheet(
    book: xlsxwriter.Workbook,
    cutoff: Worksheet,
    last: int,
    latest_period: str,
    latest_cutoff: int,
    default_totals: GroupTotals,
    formats: SheetFormats,
) -> int:
    header, area, pct, selector = (
        formats.header,
        formats.area,
        formats.percent,
        formats.selector,
    )
    year_range, presence = year_ranges(last)
    cutoff.set_column("A:A", 25)
    cutoff.set_column("B:G", 17, area)
    cutoff.write("A1", "Объединённая строка", header)
    cutoff.write_formula("B1", '=EffectiveCutoff&"+"', selector, f"{latest_cutoff}+")
    cutoff.write("A2", "Отчётный период")
    cutoff.write_formula("B2", "=ReportPeriod", selector, latest_period)
    cutoff.merge_range("B3:D3", "Продано, м²", header)
    cutoff.merge_range("E3:G3", "В строительстве, м²", header)
    cutoff.write_row("A4", ["Год ввода", *BANDS, *BANDS], header)
    cutoff.freeze_panes(4, 0)
    cutoff.write_dynamic_array_formula(
        "A5",
        f"=FILTER(Данные!$K$1:$Q${last},"
        f'({year_range}>=EffectiveCutoff)*({presence}=1),"")',
        None,
        latest_cutoff,
    )
    detail_end = 4 + last
    cutoff.conditional_format(
        f"A5:A{detail_end}",
        {
            "type": "formula",
            "criteria": "=A5=LastCompletionYear",
            "format": book.add_format({"num_format": '0"+"'}),
        },
    )
    totals_row, ratio_row = detail_end + 2, detail_end + 3
    cutoff.write(f"A{totals_row}", "Суммы по группам", header)
    cutoff.write(f"A{ratio_row}", "Распроданность", header)
    cutoff_sold, cutoff_areas = default_totals
    amounts = (*cutoff_sold, *cutoff_areas)
    for col, letter in enumerate("BCDEFG"):
        cutoff.write_formula(
            f"{letter}{totals_row}",
            f"=SUM({letter}5:{letter}{detail_end})",
            area,
            amounts[col],
        )
    for col, (sold_col, area_col) in enumerate(zip("BCD", "EFG")):
        sold = f"{sold_col}{totals_row}"
        total = f"{area_col}{totals_row}"
        cached = cutoff_sold[col] / cutoff_areas[col] if cutoff_areas[col] else ""
        cutoff.write_formula(
            f"{sold_col}{ratio_row}", f'=IF({total}=0,"",{sold}/{total})', pct, cached
        )
    return ratio_row


def year_ranges(last: int) -> tuple[str, str]:
    return f"Данные!$K$1:$K${last}", f"Данные!$R$1:$R${last}"


def define_effective_cutoff(
    book: xlsxwriter.Workbook, data: Worksheet, last: int, latest_cutoff: int
) -> None:
    year_range, presence = year_ranges(last)
    data.write_formula(
        "S2", f"=_xlfn.MAXIFS({year_range},{presence},1)", None, latest_cutoff
    )
    book.define_name("LastCompletionYear", "='Данные'!$S$2")
    data.write_formula("S1", "=MIN(CutoffYear,LastCompletionYear)", None, latest_cutoff)
    book.define_name("EffectiveCutoff", "='Данные'!$S$1")


def summary_formats(
    book: xlsxwriter.Workbook,
) -> tuple[list[Format], dict[tuple[bool, int], Format]]:
    headings = []
    cells = {}
    for col in range(4):
        left, right = (5 if col == 0 else 1), (5 if col == 3 else 1)
        headings.append(
            book.add_format(
                {
                    "bold": True,
                    "bg_color": "#143849",
                    "font_color": "white",
                    "text_wrap": True,
                    "align": "center",
                    "valign": "vcenter",
                    "left": left,
                    "right": right,
                    "top": 5,
                    "bottom": 5,
                }
            )
        )
        for is_last in (False, True):
            options = {
                "align": "center",
                "valign": "vcenter",
                "left": left,
                "right": right,
                "bottom": 5 if is_last else 1,
            }
            if col:
                options["num_format"] = "0.00%"
            cells[is_last, col] = book.add_format(options)
    return headings, cells


def write_summary_sheet(
    book: xlsxwriter.Workbook,
    report: Worksheet,
    latest_period: str,
    cutoff_years: list[int],
    latest_cutoff: int,
    last: int,
    ratio_row: int,
    latest_totals: dict[int, GroupTotals],
    formats: SheetFormats,
) -> None:
    report.hide_gridlines(2)
    report.set_column("A:A", 33)
    report.set_column("B:D", 27)
    report.write("A1", "Ульяновская область", formats.header)
    report.write("B1", "Отчётный период")
    report.write("C1", latest_period, formats.selector)
    report.data_validation("C1", {"validate": "list", "source": "=ReportPeriods"})
    book.define_name("ReportPeriod", "='Сводка'!$C$1")

    headings, body_formats = summary_formats(book)
    titles = [
        "Год ввода",
        *(f"Распроданность при стройготовности {band.lower()}" for band in BANDS),
    ]
    for col, title in enumerate(titles):
        report.write(SUMMARY_HEADER_ROW, col, title, headings[col])
    report.set_row(SUMMARY_HEADER_ROW, 44)
    report.freeze_panes(SUMMARY_HEADER_ROW + 1, 0)

    cutoff_row = SUMMARY_HEADER_ROW + len(cutoff_years) + 2
    report.write(cutoff_row, 0, "Объединить начиная с года")
    report.write_number(cutoff_row, 1, latest_cutoff, formats.selector)
    report.data_validation(
        cutoff_row, 1, cutoff_row, 1, {"validate": "list", "source": "=CutoffYears"}
    )
    book.define_name("CutoffYear", f"='Сводка'!$B${cutoff_row + 1}")

    year_range, presence = year_ranges(last)
    for row, year in enumerate(cutoff_years, SUMMARY_HEADER_ROW + 1):
        is_last = row == SUMMARY_HEADER_ROW + len(cutoff_years)
        available = f"SUMIF({year_range},{year},{presence})"
        later = f'SUMIF({year_range},">="&{year},{presence})'
        visible_year = latest_totals.get(year)
        label = (
            f"{year}+" if year == latest_cutoff else str(year) if visible_year else ""
        )
        report.write_formula(
            row,
            0,
            f'=IF(EffectiveCutoff={year},IF({later}>0,"{year}+",""),'
            f'IF(AND(EffectiveCutoff>{year},{available}>0),"{year}",""))',
            body_formats[is_last, 0],
            label,
        )
        for index, (sold_col, area_col) in enumerate(zip("LMN", "OPQ"), 1):
            sold = f"SUMIF({year_range},{year},Данные!${sold_col}$1:${sold_col}${last})"
            total = (
                f"SUMIF({year_range},{year},Данные!${area_col}$1:${area_col}${last})"
            )
            cached = ""
            if visible_year:
                sums, areas = visible_year
                cached = sums[index - 1] / areas[index - 1] if areas[index - 1] else "-"
            cutoff_ref = f"'Расчёт ГГГГ+'!{'BCD'[index - 1]}{ratio_row}"
            report.write_formula(
                row,
                index,
                f'=IF(A{row + 1}="","",IF(EffectiveCutoff={year},'
                f'IF(ISNUMBER({cutoff_ref}),{cutoff_ref},"-"),'
                f'IF({total}=0,"-",{sold}/{total})))',
                body_formats[is_last, index],
                cached,
            )


def build_workbook(output: Path, months: int) -> None:
    periods, by_year = workbook_data(months)
    latest_period = periods[0]
    cutoff_years = sorted({int(year.rstrip("+")) for year in by_year})
    latest_cutoff = max(
        int(year.rstrip("+")) for year, rows in by_year.items() if latest_period in rows
    )

    with xlsxwriter.Workbook(
        output, {"strings_to_formulas": False, "strings_to_urls": False}
    ) as book:
        report = book.add_worksheet("Сводка")
        data = book.add_worksheet("Данные")
        cutoff = book.add_worksheet("Расчёт ГГГГ+")
        formats = sheet_formats(book)
        prepare_data_sheet(book, data, periods, cutoff_years, formats)

        data_row = 1
        latest_totals = {}
        for index_row, (year, monthly_rows) in enumerate(by_year.items()):
            year_sheet = book.add_worksheet(year)
            data_row, totals = write_year_sheet(
                year_sheet,
                data,
                index_row,
                data_row,
                year,
                monthly_rows,
                latest_period,
                formats,
            )
            if latest_period in monthly_rows:
                latest_totals[int(year.rstrip("+"))] = totals

        last = len(by_year)
        define_effective_cutoff(book, data, last, latest_cutoff)
        ratio_row = write_cutoff_sheet(
            book,
            cutoff,
            last,
            latest_period,
            latest_cutoff,
            latest_totals[latest_cutoff],
            formats,
        )
        write_summary_sheet(
            book,
            report,
            latest_period,
            cutoff_years,
            latest_cutoff,
            last,
            ratio_row,
            latest_totals,
            formats,
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Собрать интерактивный отчёт Excel")
    parser.add_argument("--months", type=int, default=3)
    parser.add_argument(
        "output",
        nargs="?",
        type=Path,
        default=Path("developers.xlsx"),
        help="Выходной файл Excel (по умолчанию developers.xlsx)",
    )
    args = parser.parse_args()
    if args.months < 1:
        parser.error("--months должен быть положительным")
    build_workbook(args.output, args.months)
    print(f"Книга сохранена в {args.output}")


if __name__ == "__main__":
    main()
