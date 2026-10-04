#!/usr/bin/env python3
"""Build the quarterly TPS/TVQ declaration workbook from a Shopify orders JSON dump.

Reproduces the "Section A — Revenus" sheet:

  row 6   TAUX EN VIGUEUR
  row 7   Taux TPS (fédéral) 5,00 %   |   Taux TVQ (provincial) 9,975 %
  row 9   SECTION A — REVENUS (ventes et services facturés)
  row 10  Case | Description | Montant HT (Brut) ($) | TPS perçue ($) | TVQ perçue ($)
  row 11  101     Vente - En ligne (shopify): total sales before tax, and the
                  TPS/TVQ those sales represent at the rates of row 7
  row 12  105/205 Taxes perçues (clients): sales on which tax was charged, and
                  the TPS/TVQ actually collected

E12 and F12 are the "tps" and "tvq" of the "TOTAL PÉRIODE" row of the
"Suivi ventes" workbook built from the same orders (same per-day rounding),
so the two files always agree to the cent.

Usage:
  python scripts/build_tax_declaration_xlsx.py [--input data/orders/2026-T4.json] [--output path.xlsx]

Defaults: --input is the most recent extraction JSON in data/orders/,
--output is data/orders/<same-name>_declaration_tps_tvq.xlsx.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from build_sales_tracking_xlsx import (
    DATA_DIR,
    find_latest_input,
    load_orders,
    period_totals,
    split_taxes,
)

TPS_RATE = 0.05
TVQ_RATE = 0.09975

MONEY_FORMAT = '#,##0.00 "$"'
TPS_RATE_FORMAT = "0.00%"
TVQ_RATE_FORMAT = "0.000%"

TEAL_FILL = PatternFill("solid", fgColor="FF0E6E8C")
SECTION_FILL = PatternFill("solid", fgColor="FF1F4E78")
HEADER_FILL = PatternFill("solid", fgColor="FF2E75B6")
RATE_FILL = PatternFill("solid", fgColor="FFDCE6F1")
CASE_FILL = PatternFill("solid", fgColor="FFDCE6F1")
INPUT_FILL = PatternFill("solid", fgColor="FFFFFF00")
TAX_FILL = PatternFill("solid", fgColor="FFEEF4FB")

THIN = Side(style="thin", color="FFBFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

WHITE_TITLE = Font(name="Arial", bold=True, size=14, color="FFFFFFFF")
WHITE_HEADER = Font(name="Arial", bold=True, size=13, color="FFFFFFFF")
TEXT = Font(name="Arial", size=12)
BLUE_BOLD = Font(name="Arial", bold=True, size=13, color="FF0000FF")
BLUE = Font(name="Arial", size=12, color="FF0000FF")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")


def declaration_amounts(orders: list[dict]) -> dict[str, float]:
    totals = period_totals(orders)
    # Sales on which TPS or TVQ was charged (orders shipped outside Québec /
    # Canada can have no tax), summed with the same per-day rounding.
    taxed = [o for o in orders if any(split_taxes(o))]
    taxed_sales = period_totals(taxed)["mt_brut"] if taxed else 0.0
    return {
        "ventes": totals["mt_brut"],
        "tps_theorique": round(totals["mt_brut"] * TPS_RATE, 2),
        "tvq_theorique": round(totals["mt_brut"] * TVQ_RATE, 2),
        "ventes_taxees": taxed_sales,
        "tps": totals["tps"],
        "tvq": totals["tvq"],
    }


def _cell(ws, ref: str, value=None, font=TEXT, fill=None, align=CENTER, fmt=None, border=True):
    cell = ws[ref]
    if value is not None:
        cell.value = value
    cell.font = font
    cell.alignment = align
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    if border:
        cell.border = BORDER
    return cell


def _banner(ws, row: int, text: str, fill: PatternFill) -> None:
    ws.merge_cells(f"C{row}:F{row}")
    _cell(ws, f"C{row}", text, font=WHITE_TITLE, fill=fill)
    for col in "DEF":
        ws[f"{col}{row}"].fill = fill
        ws[f"{col}{row}"].border = BORDER
    ws.row_dimensions[row].height = 24


def build_declaration(orders: list[dict], period: str = "", warnings: list[str] | None = None) -> Workbook:
    amounts = declaration_amounts(orders)

    wb = Workbook()
    ws = wb.active
    ws.title = "Déclaration TPS-TVQ"
    ws.sheet_view.showGridLines = True

    for col, width in {"A": 4, "B": 12, "C": 30, "D": 21, "E": 29, "F": 29}.items():
        ws.column_dimensions[col].width = width

    _cell(ws, "C2", "Déclaration TPS / TVQ", font=Font(name="Arial", bold=True, size=16),
          align=LEFT, border=False)
    if period:
        _cell(ws, "C3", f"Période : {period}", font=TEXT, align=LEFT, border=False)
    if warnings:
        _cell(ws, "C4", "ATTENTION — données incomplètes : " + " ".join(warnings),
              font=Font(name="Arial", bold=True, size=11, color="FFFF0000"), align=LEFT, border=False)

    _banner(ws, 6, "TAUX EN VIGUEUR", TEAL_FILL)
    _cell(ws, "C7", "Taux TPS (fédéral)", align=LEFT)
    _cell(ws, "D7", TPS_RATE, font=BLUE_BOLD, fill=RATE_FILL, fmt=TPS_RATE_FORMAT)
    _cell(ws, "E7", "Taux TVQ (provincial)", align=LEFT)
    _cell(ws, "F7", TVQ_RATE, font=BLUE_BOLD, fill=RATE_FILL, fmt=TVQ_RATE_FORMAT)

    _banner(ws, 9, "SECTION A — REVENUS (ventes et services facturés)", SECTION_FILL)

    headers = {"B": "Case", "C": "Description", "D": "Montant HT (Brut) ($)",
               "E": "TPS perçue ($)", "F": "TVQ perçue ($)"}
    for col, text in headers.items():
        _cell(ws, f"{col}10", text, font=WHITE_HEADER, fill=HEADER_FILL)
    ws.row_dimensions[10].height = 36

    rows = [
        (11, "101", "Vente -  En ligne (shopify)",
         amounts["ventes"], amounts["tps_theorique"], amounts["tvq_theorique"]),
        (12, "105/205", "Taxes perçues (clients)",
         amounts["ventes_taxees"], amounts["tps"], amounts["tvq"]),
    ]
    for row, case, label, montant, tps, tvq in rows:
        _cell(ws, f"B{row}", case, fill=CASE_FILL if row == 11 else None)
        _cell(ws, f"C{row}", label, align=LEFT)
        _cell(ws, f"D{row}", montant, font=BLUE, fill=INPUT_FILL, fmt=MONEY_FORMAT)
        _cell(ws, f"E{row}", tps, fill=TAX_FILL, fmt=MONEY_FORMAT)
        _cell(ws, f"F{row}", tvq, fill=TAX_FILL, fmt=MONEY_FORMAT)
        ws.row_dimensions[row].height = 18

    _cell(ws, "C14", "Ligne 101 : TPS/TVQ calculées sur le montant HT aux taux en vigueur (indicatif).",
          font=Font(name="Arial", italic=True, size=9, color="FF595959"), align=LEFT, border=False)
    _cell(ws, "C15", "Ligne 105/205 : TPS/TVQ réellement perçues = TOTAL PÉRIODE du fichier Suivi ventes.",
          font=Font(name="Arial", italic=True, size=9, color="FF595959"), align=LEFT, border=False)

    return wb


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=None)
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args()

    input_path = args.input or find_latest_input()
    orders = load_orders(input_path)
    output_path = args.output or DATA_DIR / f"{input_path.stem}_declaration_tps_tvq.xlsx"

    build_declaration(orders, period=input_path.stem).save(output_path)
    print(f"Wrote {output_path} from {len(orders)} orders in {input_path}")


if __name__ == "__main__":
    main()
