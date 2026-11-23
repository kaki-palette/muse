"""在庫・発注管理の動作確認用ブックを生成する。"""

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# 商品名、単価、発注点、安全在庫、リードタイム、現在庫、発注残、入庫予定日
PRODUCTS = (
    ("試作商品01", 120, 20, 5, 3, 12, 3, date(2030, 1, 8)),
    ("試作商品02", 240, 15, 4, 5, 15, 0, None),
    ("試作商品03", 360, 18, 6, 7, 11, 7, date(2030, 1, 10)),
    ("試作商品04", 480, 24, 8, 4, 26, 0, None),
    ("試作商品05", 150, 12, 3, 2, 0, 0, None),
    ("試作商品06", 270, 30, 10, 6, 24, 4, date(2030, 1, 12)),
    ("試作商品07", 390, 16, 4, 3, 8, 11, date(2030, 1, 9)),
    ("試作商品08", 510, 9, 2, 8, 3, 0, None),
    ("試作商品09", 630, 21, 7, 5, 6, 15, date(2030, 1, 11)),
    ("試作商品10", 750, 14, 5, 4, 10, 1, date(2030, 1, 8)),
)


def format_sheet(sheet, widths):
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 25
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="24465B")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center")
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def main():
    book = Workbook()
    book.properties.creator = ""
    book.calculation.calcMode = "auto"
    book.calculation.fullCalcOnLoad = True

    master = book.active
    master.title = "商品マスタ"
    ledger = book.create_sheet("在庫台帳")
    orders = book.create_sheet("発注リスト")

    master.append(("商品名", "単価", "発注点", "安全在庫", "リードタイム日数"))
    ledger.append(("商品名", "現在庫", "発注残", "入庫予定日"))
    orders.append(("商品名", "現在庫", "発注点", "発注要否", "発注数量"))

    last_row = len(PRODUCTS) + 1
    ledger_range = f"'在庫台帳'!$A$2:$C${last_row}"
    master_range = f"'商品マスタ'!$A$2:$C${last_row}"

    for row, product in enumerate(PRODUCTS, start=2):
        name, price, reorder_point, safety_stock, lead_time, stock, pending, due = product
        master.append((name, price, reorder_point, safety_stock, lead_time))
        ledger.append((name, stock, pending, due))
        ledger[f"D{row}"].number_format = "yyyy/mm/dd"
        master[f"B{row}"].number_format = "#,##0"

        pending_lookup = f"VLOOKUP(A{row},{ledger_range},3,FALSE)"
        orders.append((
            f"='商品マスタ'!A{row}",
            f"=VLOOKUP(A{row},{ledger_range},2,FALSE)",
            f"=VLOOKUP(A{row},{master_range},3,FALSE)",
            f'=IF(B{row}+{pending_lookup}<C{row},"要","不要")',
            f"=MAX(0,C{row}-B{row}-{pending_lookup})",
        ))

    format_sheet(master, (20, 12, 12, 12, 21))
    format_sheet(ledger, (20, 12, 12, 17))
    format_sheet(orders, (20, 12, 12, 14, 14))
    book.save(Path(__file__).with_name("inventory-order.xlsx"))


if __name__ == "__main__":
    main()
