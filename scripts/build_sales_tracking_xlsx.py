#!/usr/bin/env python3
"""Build the "Suivi ventes" Excel workbook from a Shopify orders JSON dump.

Reproduces the layout of the reference template: one block of 6 columns
per calendar month (Date, Commande, Mt brut, tvq, tps, Expedition), one row
per calendar day, and a totals row per month (computed in Python, written
as plain numbers so they display correctly without depending on Excel
recalculation), plus a grand-total row summing every month. Blocks are
laid out left to right, one per month, covering the same date range as the
orders JSON (typically the 90-day rolling extraction window).

If more than one order falls on the same calendar day, their amounts are
summed onto that day's single row and their order numbers are joined with
", " in the Commande cell (the template only has room for one row per day).

Usage:
  python scripts/build_sales_tracking_xlsx.py [--input path/to/orders.json] [--output path/to/output.xlsx]

Defaults: --input is the most recent data/orders/YYYY-MM-DD.json file,
--output is data/orders/<same-date>_suivi_ventes.xlsx.
"""

from __future__ import annotations

import argparse
import calendar
import json
from collections import defaultdict
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "orders"

MONTH_ABBR = {
    1: "janv", 2: "févr", 3: "mars", 4: "avr", 5: "mai", 6: "juin",
    7: "juil", 8: "août", 9: "sept", 10: "oct", 11: "nov", 12: "déc",
}

HEADERS = ["Date", "Commande", "Mt brut", "tvq", "tps", "Expedition"]
BLOCK_WIDTH = 7  # 6 data columns + 1 blank separator

THIN_GRAY = Side(style="thin", color="FFBFBFBF")
CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)

HEADER_FILL = PatternFill("solid", fgColor="FFF8CBAD")
EXPEDITION_HEADER_FILL = PatternFill("solid", fgColor="FFFFFF00")
DAY_WITH_ORDER_FILL = PatternFill("solid", fgColor="FFE2EFDA")
TOTAL_LABEL_FILL = PatternFill("solid", fgColor="FF00B0F0")
TOTAL_AMOUNT_FILL = PatternFill("solid", fgColor="FFCCC0DA")
TOTAL_TAX_FILL = PatternFill("solid", fgColor="FFFF0000")

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")

# Escaped period so Excel always displays a decimal point, regardless of the
# spreadsheet locale (a plain "0.00" format renders with a comma under a
# French locale).
NUMBER_FORMAT = '0"."00'


def load_orders(path: Path) -> list[dict]:
    return json.loads(path.read_text(encoding="utf-8"))


def find_latest_input() -> Path:
    candidates = sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9].json"))
    if not candidates:
        raise SystemExit(f"No dated orders JSON file found in {DATA_DIR}")
    return candidates[-1]


def order_date(order: dict) -> date:
    return date.fromisoformat(order["created_at"][:10])


def split_taxes(order: dict) -> tuple[float, float]:
    tps = 0.0
    tvq = 0.0
    for tax_line in order.get("tax_lines") or []:
        title = (tax_line.get("title") or "").upper()
        price = float(tax_line.get("price") or 0)
        if "GST" in title or "TPS" in title:
            tps += price
        elif "QST" in title or "TVQ" in title:
            tvq += price
    return tps, tvq


def shipping_fee(order: dict) -> float:
    amount = (order.get("total_shipping_price_set") or {}).get("shop_money", {}).get("amount")
    return float(amount) if amount is not None else 0.0


def group_by_day(orders: list[dict]) -> dict[date, list[dict]]:
    grouped: dict[date, list[dict]] = defaultdict(list)
    for order in orders:
        grouped[order_date(order)].append(order)
    return grouped


def months_covered(orders: list[dict]) -> list[tuple[int, int]]:
    days = sorted({order_date(o) for o in orders})
    if not days:
        raise SystemExit("No orders to build a workbook from")
    start, end = days[0].replace(day=1), days[-1].replace(day=1)
    months = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        months.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def write_block(ws, start_col: int, year: int, month: int, day_orders: dict[date, list[dict]]) -> tuple[int, dict[str, float]]:
    days_in_month = calendar.monthrange(year, month)[1]
    abbr = MONTH_ABBR[month]

    for offset, header in enumerate(HEADERS):
        cell = ws.cell(row=1, column=start_col + offset, value=header)
        cell.fill = EXPEDITION_HEADER_FILL if header == "Expedition" else HEADER_FILL
        cell.font = Font(bold=True, size=10)
        cell.alignment = CENTER
        cell.border = CELL_BORDER

    totals = {"mt_brut": 0.0, "tvq": 0.0, "tps": 0.0, "expedition": 0.0}

    for day in range(1, days_in_month + 1):
        row = 1 + day
        d = date(year, month, day)
        date_cell = ws.cell(row=row, column=start_col, value=f"{day:02d}-{abbr}")
        date_cell.alignment = LEFT
        date_cell.border = CELL_BORDER
        date_cell.font = Font(size=10)

        orders_today = day_orders.get(d, [])
        cells = [ws.cell(row=row, column=start_col + i) for i in range(1, 6)]
        for c in cells:
            c.alignment = CENTER
            c.border = CELL_BORDER
            c.font = Font(size=10)

        if orders_today:
            commande = ", ".join(o.get("name", "") for o in orders_today)
            mt_brut = round(sum(float(o.get("subtotal_price") or 0) for o in orders_today), 2)
            tps_total = round(sum(split_taxes(o)[0] for o in orders_today), 2)
            tvq_total = round(sum(split_taxes(o)[1] for o in orders_today), 2)
            expedition = round(sum(shipping_fee(o) for o in orders_today), 2)

            cells[0].value = commande
            cells[1].value = mt_brut
            cells[2].value = tvq_total
            cells[3].value = tps_total
            cells[4].value = expedition
            for c in [date_cell, *cells]:
                c.fill = DAY_WITH_ORDER_FILL

            totals["mt_brut"] += mt_brut
            totals["tvq"] += tvq_total
            totals["tps"] += tps_total
            totals["expedition"] += expedition

        for c in cells[1:]:
            c.number_format = NUMBER_FORMAT

    total_row = 2 + days_in_month
    label_cell = ws.cell(row=total_row, column=start_col, value=f"{abbr}-{str(year)[-2:]}")
    label_cell.fill = TOTAL_LABEL_FILL
    label_cell.font = Font(bold=True, size=16, color="FFFFFFFF")
    label_cell.alignment = CENTER

    col_letters = [get_column_letter(start_col + i) for i in range(1, 6)]
    fills = [TOTAL_AMOUNT_FILL, TOTAL_TAX_FILL, TOTAL_TAX_FILL, TOTAL_AMOUNT_FILL]
    values = [round(totals["mt_brut"], 2), round(totals["tvq"], 2), round(totals["tps"], 2), round(totals["expedition"], 2)]
    for letter, fill, value in zip(col_letters[1:], fills, values):
        cell = ws[f"{letter}{total_row}"]
        cell.value = value
        cell.number_format = NUMBER_FORMAT
        cell.fill = fill
        cell.font = Font(bold=True, size=10)
        cell.alignment = CENTER
        cell.border = CELL_BORDER

    ws.row_dimensions[total_row].height = 25.5

    widths = {0: 9, 1: 11, 2: 9, 3: 8, 4: 8, 5: 10}
    for offset, width in widths.items():
        ws.column_dimensions[get_column_letter(start_col + offset)].width = width
    ws.column_dimensions[get_column_letter(start_col + BLOCK_WIDTH - 1)].width = 3

    return total_row, totals


def write_grand_total(ws, block_totals: list[tuple[int, dict[str, float]]]) -> None:
    """Write a grand-total row below all month blocks, summing every block's
    Mt brut / tvq / tps / Expedition totals into one row."""
    max_total_row = max(total_row for total_row, _ in block_totals)
    grand_row = max_total_row + 2

    label_cell = ws.cell(row=grand_row, column=2, value="TOTAL PÉRIODE")
    label_cell.fill = TOTAL_LABEL_FILL
    label_cell.font = Font(bold=True, size=16, color="FFFFFFFF")
    label_cell.alignment = CENTER
    ws.merge_cells(start_row=grand_row, start_column=2, end_row=grand_row, end_column=3)
    ws.row_dimensions[grand_row].height = 25.5

    fills = [TOTAL_AMOUNT_FILL, TOTAL_TAX_FILL, TOTAL_TAX_FILL, TOTAL_AMOUNT_FILL]
    target_cols = [4, 5, 6, 7]  # D, E, F, G: Mt brut, tvq, tps, Expedition
    keys = ["mt_brut", "tvq", "tps", "expedition"]
    for target_col, fill, key in zip(target_cols, fills, keys):
        value = round(sum(totals[key] for _, totals in block_totals), 2)
        cell = ws.cell(row=grand_row, column=target_col, value=value)
        cell.number_format = NUMBER_FORMAT
        cell.fill = fill
        cell.font = Font(bold=True, size=10)
        cell.alignment = CENTER
        cell.border = CELL_BORDER


def build_workbook(orders: list[dict]) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Suivi ventes"
    ws.freeze_panes = "A2"

    day_orders = group_by_day(orders)
    block_totals = []
    for i, (year, month) in enumerate(months_covered(orders)):
        total_row, totals = write_block(
            ws, start_col=2 + i * BLOCK_WIDTH, year=year, month=month, day_orders=day_orders
        )
        block_totals.append((total_row, totals))

    write_grand_total(ws, block_totals)

    return wb


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=None)
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args()

    input_path = args.input or find_latest_input()
    orders = load_orders(input_path)

    output_path = args.output
    if output_path is None:
        output_path = DATA_DIR / f"{input_path.stem}_suivi_ventes.xlsx"

    wb = build_workbook(orders)
    wb.save(output_path)
    print(f"Wrote {output_path} from {len(orders)} orders in {input_path}")


if __name__ == "__main__":
    main()
