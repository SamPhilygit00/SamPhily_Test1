#!/usr/bin/env python3
"""Import a Shopify admin orders export (Commandes -> Exporter -> CSV) into the
monthly files data/orders/mensuel/AAAA-MM.json.

Use it for a month the API can no longer return (older than 60 days). The
export has one row per line item; order-level fields (Subtotal, Shipping,
Tax 1..5 Name/Value, Created at, ...) are read from each order's first row
and converted to the same JSON shape as the Admin API orders, so the quarter
files are built exactly as for the other months.

Orders are merged by id into any existing monthly file (the CSV wins).

Usage:
  python scripts/import_shopify_csv.py orders_export.csv
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

MONTHLY_DIR = Path(__file__).resolve().parent.parent / "data" / "orders" / "mensuel"


def to_iso(value: str) -> str | None:
    """'2026-07-01 20:00:31 -0400' -> '2026-07-01T20:00:31-04:00'."""
    value = (value or "").strip()
    if not value:
        return None
    return datetime.strptime(value, "%Y-%m-%d %H:%M:%S %z").isoformat()


def tax_lines(row: dict) -> list[dict]:
    lines = []
    for i in range(1, 6):
        name = (row.get(f"Tax {i} Name") or "").strip()
        value = (row.get(f"Tax {i} Value") or "").strip()
        if name and value:
            # "GST 5%" -> "GST", "QST 9,975%" -> "QST"
            lines.append({"title": re.split(r"\s+\d", name)[0].strip(), "price": f"{float(value):.2f}"})
    return lines


def row_to_order(row: dict) -> dict:
    created_at = to_iso(row["Created at"])
    return {
        "id": int(row["Id"]),
        "name": row["Name"],
        "created_at": created_at,
        "updated_at": created_at,
        "cancelled_at": to_iso(row.get("Cancelled at")),
        "financial_status": row.get("Financial Status") or None,
        "fulfillment_status": row.get("Fulfillment Status") or None,
        "currency": row.get("Currency") or "CAD",
        "subtotal_price": row.get("Subtotal") or "0.00",
        "total_discounts": row.get("Discount Amount") or "0.00",
        "total_price": row.get("Total") or "0.00",
        "tax_lines": tax_lines(row),
        "total_shipping_price_set": {"shop_money": {"amount": row.get("Shipping") or "0.00"}},
        "shipping_address": {"city": row.get("Shipping City") or None},
        "source_name": row.get("Source") or None,
        "imported_from": "shopify_admin_csv_export",
    }


def read_export(path: Path) -> list[dict]:
    with path.open(encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    first_rows = {}
    for row in rows:
        if row.get("Name") and row.get("Created at"):
            first_rows.setdefault(row["Name"], row)
    return [row_to_order(r) for r in first_rows.values()]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("csv", type=Path)
    args = parser.parse_args()

    by_month: dict[str, list[dict]] = defaultdict(list)
    for order in read_export(args.csv):
        by_month[order["created_at"][:7]].append(order)

    MONTHLY_DIR.mkdir(parents=True, exist_ok=True)
    for month, orders in sorted(by_month.items()):
        path = MONTHLY_DIR / f"{month}.json"
        existing = json.loads(path.read_text(encoding="utf-8")) if path.exists() else []
        merged = {o["id"]: o for o in existing}
        merged.update({o["id"]: o for o in orders})
        result = sorted(merged.values(), key=lambda o: o["created_at"], reverse=True)
        path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
        names = sorted(o["name"] for o in orders)
        print(f"{month} : {len(orders)} commandes importées ({names[0]} à {names[-1]}), "
              f"{len(result)} au total -> {path.name}")


if __name__ == "__main__":
    main()
