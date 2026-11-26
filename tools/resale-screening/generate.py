"""フリマ価格差せどりの候補判定・監視リストを生成する。"""

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


CANDIDATES = (
    ("オレンズネロ限定セット グリーンブラック0.3mm", "Yahoo!フリマ", 7000, 0, 10000),
    ("クルトガダイブ オーロラパープル", "Yahoo!フリマ", 6950, 0, 9400),
    ("クルトガダイブ アビスブルー", "メルカリ", 4300, 0, 6200),
)

WATCH_ITEMS = (
    ("クルトガダイブ 標準色（各種）", 6500, 9500, 0, None),
    ("クルトガダイブ フルブラック", 9800, 9800, 0, None),
    ("オレンズネロ限定セット", 7200, 10000, 0, None),
    ("にじさんじ誕生日 アクリルブロックセット（現行仕様）", None, None, 682,
     "分岐10,992円。2027年春に実績再確認"),
)

CHECKED_ON = date(2026, 11, 26)


def format_sheet(sheet, widths):
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 34
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="24465B")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def main():
    book = Workbook()
    book.properties.creator = ""
    book.calculation.calcMode = "auto"
    book.calculation.fullCalcOnLoad = True

    candidates = book.active
    candidates.title = "候補判定"
    candidates.append((
        "商品名", "仕入れ先", "仕入れ先URL", "仕入れ価格", "仕入れ送料",
        "想定売却価格（メルカリ売却済み中央値）", "売却実績URL", "メモ",
        "販売手数料", "利益", "利益率", "分岐売却価格", "判定",
    ))

    for row, (name, source, cost, shipping, sale) in enumerate(CANDIDATES, start=2):
        candidates.append((
            name, source, None, cost, shipping, sale, None, None,
            f'=IF(ISNUMBER(F{row}),F{row}*10%,"")',
            f'=IF(COUNT(D{row}:F{row})=3,F{row}-I{row}-210-D{row}-E{row},"")',
            f'=IF(AND(ISNUMBER(J{row}),F{row}<>0),J{row}/F{row},"")',
            f'=IF(COUNT(D{row}:E{row})=2,CEILING((D{row}+E{row}+210+1000)/0.9,1),"")',
            f'=IF(ISNUMBER(J{row}),IF(J{row}>=1000,"○",IF(J{row}>=0,"△","×")),"")',
        ))
        for column in ("D", "E", "F", "I", "J", "L"):
            candidates[f"{column}{row}"].number_format = "#,##0.##"
        candidates[f"K{row}"].number_format = "0.0%"
    format_sheet(candidates, (49, 17, 28, 14, 14, 25, 28, 28, 14, 12, 12, 16, 10))

    watch = book.create_sheet("監視リスト")
    watch.append((
        "対象商品", "メルカリ売却済み中央値", "メルカリ売却済み最高値",
        "仕入れ送料の想定", "買いライン", "最終確認日", "メモ",
    ))
    for row, (name, median, highest, shipping, note) in enumerate(WATCH_ITEMS, start=2):
        watch.append((
            name, median, highest, shipping,
            f'=IF(AND(ISNUMBER(B{row}),ISNUMBER(D{row})),FLOOR(B{row}*0.9-210-1000-D{row},1),"")',
            CHECKED_ON, note,
        ))
        for column in ("B", "C", "D", "E"):
            watch[f"{column}{row}"].number_format = "#,##0"
        watch[f"F{row}"].number_format = "yyyy/mm/dd"
    format_sheet(watch, (56, 23, 23, 20, 15, 15, 42))

    book.save(Path(__file__).with_name("resale-screening.xlsx"))


if __name__ == "__main__":
    main()
