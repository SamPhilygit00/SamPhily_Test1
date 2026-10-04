#!/usr/bin/env python3
"""Extract Shopify orders month by month and build the quarterly TPS/TVQ
declaration files.

Shopify's Admin API only returns the last 60 days of orders (without the
read_all_orders scope), so a quarter can't be fetched in one go once it is
over. Instead:

  * EXTRACTION=mensuel (scheduled run, 1st of every month): fetches the month
    that just ended and stores it in data/orders/mensuel/AAAA-MM.json. When
    that month closes a quarter (runs of Jan/Apr/Jul/Oct 1st), the quarter
    files are built from the three stored months and emailed.
  * EXTRACTION=trimestre (default, manual run): builds the quarter files for
    TRIMESTRE right away, from the stored months, fetching from Shopify any
    month not stored yet. A month that is not stored and started more than
    60 days ago can only be fetched partially: the files and the email are
    then flagged INCOMPLET.

Quarter files go to data/orders/, named after the quarter: 2026-T4.json,
2026-T4.csv, 2026-T4_suivi_ventes.xlsx, 2026-T4_declaration_tps_tvq.xlsx.

Required environment variables:
  SHOPIFY_STORE_URL      e.g. "my-store.myshopify.com"

  Authentication - provide EITHER of these:
  SHOPIFY_ACCESS_TOKEN   a static Admin API access token (legacy custom apps)
  OR
  SHOPIFY_CLIENT_ID      Client ID from a Dev Dashboard custom app
  SHOPIFY_CLIENT_SECRET  Client secret from a Dev Dashboard custom app
                         (a fresh access token is requested via the client
                         credentials grant on every run)

  After extraction, the CSV and both xlsx files are emailed via SMTP. Required for that:
  SMTP_USERNAME  the sending account's login (e.g. a Yahoo Mail address)
  SMTP_PASSWORD  an app password for that account

Optional environment variables:
  EXTRACTION            "mensuel" or "trimestre" (default), see above
  TRIMESTRE             quarter to extract: empty (default) = last completed
                        quarter, "actuel" = current quarter to date, or an
                        explicit quarter such as "2026-T3"
  SHOPIFY_TIMEZONE      store time zone used for quarter boundaries,
                        default "America/Toronto"
  SHOPIFY_API_VERSION   default "2024-10"
  SHOPIFY_ORDER_STATUS  default "any" (any|open|closed|cancelled)
  SMTP_HOST      default "smtp.mail.yahoo.com"
  SMTP_PORT      default 465 (implicit TLS)
  EMAIL_FROM     default SMTP_USERNAME
  EMAIL_TO       default "eladdas@yahoo.fr"
"""

from __future__ import annotations

import csv
import json
import os
import re
import smtplib
import sys
import time
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

import requests

from build_sales_tracking_xlsx import build_workbook
from build_tax_declaration_xlsx import build_declaration

DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "orders"
MONTHLY_DIR = DATA_DIR / "mensuel"
# Shopify returns 60 days of orders; keep a safety margin.
HISTORY_DAYS = 59
PAGE_LIMIT = 250
QUARTER_RE = re.compile(r"^\s*(\d{4})\s*-?\s*[TtQq]\s*([1-4])\s*$")
LINK_NEXT_RE = re.compile(r'<([^>]+)>;\s*rel="next"')


def env_or_die(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        print(f"Missing required environment variable: {name}", file=sys.stderr)
        sys.exit(1)
    return value


def get_access_token_via_client_credentials(store_url: str, client_id: str, client_secret: str) -> str:
    url = f"https://{store_url}/admin/oauth/access_token"
    resp = requests.post(
        url,
        json={
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    resp.raise_for_status()
    token = resp.json().get("access_token")
    if not token:
        print("Client credentials grant did not return an access_token", file=sys.stderr)
        sys.exit(1)
    return token


def build_session(access_token: str) -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "X-Shopify-Access-Token": access_token,
            "Content-Type": "application/json",
        }
    )
    return session


def quarter_of(d: date) -> tuple[int, int]:
    return d.year, (d.month - 1) // 3 + 1


def resolve_quarter(spec: str, today: date) -> tuple[int, int]:
    """Turn the TRIMESTRE setting into a (year, quarter) pair."""
    spec = (spec or "").strip()
    if not spec:
        year, q = quarter_of(today)
        return (year, q - 1) if q > 1 else (year - 1, 4)
    if spec.lower() in ("actuel", "current"):
        return quarter_of(today)
    match = QUARTER_RE.match(spec)
    if not match:
        print(f"TRIMESTRE invalide : {spec!r} (attendu : vide, 'actuel' ou 'AAAA-TN', ex. 2026-T3)",
              file=sys.stderr)
        sys.exit(1)
    return int(match.group(1)), int(match.group(2))


def quarter_months(year: int, quarter: int) -> list[tuple[int, int]]:
    return [(year, 3 * quarter - 2 + i) for i in range(3)]


def quarter_bounds(year: int, quarter: int, tz: ZoneInfo) -> tuple[datetime, datetime]:
    """Start (inclusive) and end (exclusive) of the quarter in the store time zone."""
    start = datetime(year, 3 * quarter - 2, 1, tzinfo=tz)
    end = datetime(year + 1, 1, 1, tzinfo=tz) if quarter == 4 else datetime(year, 3 * quarter + 1, 1, tzinfo=tz)
    return start, end


def month_bounds(year: int, month: int, tz: ZoneInfo) -> tuple[datetime, datetime]:
    """Start (inclusive) and end (exclusive) of the month in the store time zone."""
    start = datetime(year, month, 1, tzinfo=tz)
    end = datetime(year + 1, 1, 1, tzinfo=tz) if month == 12 else datetime(year, month + 1, 1, tzinfo=tz)
    return start, end


def previous_month(today: date) -> tuple[int, int]:
    return (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)


def monthly_path(year: int, month: int) -> Path:
    return MONTHLY_DIR / f"{year}-{month:02d}.json"


def save_month(year: int, month: int, orders: list[dict]) -> Path:
    MONTHLY_DIR.mkdir(parents=True, exist_ok=True)
    path = monthly_path(year, month)
    path.write_text(json.dumps(orders, indent=2), encoding="utf-8")
    print(f"Mois {year}-{month:02d} enregistré : {len(orders)} commandes -> {path.name}")
    return path


def collect_quarter(year: int, quarter: int, tz: ZoneInfo, now: datetime, fetch) -> tuple[list[dict], list[str]]:
    """Gather the quarter's orders from the stored months, fetching (and storing)
    any finished month not stored yet. Returns (orders, warnings)."""
    orders: list[dict] = []
    warnings: list[str] = []
    for y, m in quarter_months(year, quarter):
        start, end = month_bounds(y, m, tz)
        label = f"{y}-{m:02d}"
        if start > now:
            continue
        if end > now:
            print(f"Mois {label} en cours : extrait jusqu'à maintenant (non enregistré)")
            orders += fetch(start, now)
        elif monthly_path(y, m).exists():
            print(f"Mois {label} : fichier mensuel enregistré utilisé")
            orders += json.loads(monthly_path(y, m).read_text(encoding="utf-8"))
        elif start < now - timedelta(days=HISTORY_DAYS):
            warnings.append(
                f"Mois {label} INCOMPLET : aucun fichier mensuel enregistré et Shopify ne donne "
                f"que les {HISTORY_DAYS + 1} derniers jours de commandes."
            )
            orders += fetch(start, end)
        else:
            month_orders = fetch(start, end)
            save_month(y, m, month_orders)
            orders += month_orders
    unique = {o["id"]: o for o in orders}
    return sorted(unique.values(), key=lambda o: o["created_at"], reverse=True), warnings


def fetch_all_orders(store_url: str, session: requests.Session, api_version: str,
                      status: str, created_at_min: str, created_at_max: str) -> list[dict]:
    base = f"https://{store_url}/admin/api/{api_version}/orders.json"
    params = {"limit": PAGE_LIMIT, "status": status,
              "created_at_min": created_at_min, "created_at_max": created_at_max}

    orders: list[dict] = []
    url = base
    while url:
        for attempt in range(5):
            resp = session.get(url, params=params if url == base else None, timeout=30)
            if resp.status_code == 429:
                retry_after = float(resp.headers.get("Retry-After", 2))
                time.sleep(retry_after)
                continue
            resp.raise_for_status()
            break
        else:
            resp.raise_for_status()

        payload = resp.json()
        batch = payload.get("orders", [])
        orders.extend(batch)
        print(f"Fetched {len(batch)} orders (total so far: {len(orders)})")

        link_header = resp.headers.get("Link", "")
        match = LINK_NEXT_RE.search(link_header)
        url = match.group(1) if match else None
        params = None

    return orders


def split_taxes(order: dict) -> tuple[str, str]:
    """Split an order's tax_lines into (TPS, TVQ) amounts.

    Shopify tags federal GST/TPS lines with title "GST" and Quebec QST/TVQ
    lines with title "QST" (rate ~0.05 and ~0.09975 respectively).
    """
    tps = 0.0
    tvq = 0.0
    for tax_line in order.get("tax_lines") or []:
        title = (tax_line.get("title") or "").upper()
        price = float(tax_line.get("price") or 0)
        if "GST" in title or "TPS" in title:
            tps += price
        elif "QST" in title or "TVQ" in title:
            tvq += price
    return f"{tps:.2f}", f"{tvq:.2f}"


def shipping_fee(order: dict) -> str:
    amount = (order.get("total_shipping_price_set") or {}).get("shop_money", {}).get("amount")
    return amount if amount is not None else "0.00"


def flatten_order(order: dict) -> dict:
    shipping = order.get("shipping_address") or {}
    tps, tvq = split_taxes(order)
    return {
        "id": order.get("id"),
        "order_number": order.get("name"),
        "created_at": order.get("created_at"),
        "updated_at": order.get("updated_at"),
        "cancelled_at": order.get("cancelled_at"),
        "financial_status": order.get("financial_status"),
        "fulfillment_status": order.get("fulfillment_status"),
        "currency": order.get("currency"),
        "subtotal_price": order.get("subtotal_price"),
        "total_discounts": order.get("total_discounts"),
        "TPS": tps,
        "TVQ": tvq,
        "frais_livraison": shipping_fee(order),
        "total_price": order.get("total_price"),
        "shipping_city": shipping.get("city"),
    }


def write_outputs(orders: list[dict], label: str, months: list[tuple[int, int]],
                  period: str, warnings: list[str] | None = None) -> list[Path]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    json_path = DATA_DIR / f"{label}.json"
    json_path.write_text(json.dumps(orders, indent=2), encoding="utf-8")

    rows = [flatten_order(o) for o in orders]
    csv_path = DATA_DIR / f"{label}.csv"
    fieldnames = list(rows[0].keys()) if rows else [
        "id", "order_number", "created_at", "updated_at", "cancelled_at",
        "financial_status", "fulfillment_status", "currency", "subtotal_price",
        "total_discounts", "TPS", "TVQ", "frais_livraison", "total_price",
        "shipping_city",
    ]
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    attachments = [csv_path]
    if orders:
        xlsx_path = DATA_DIR / f"{label}_suivi_ventes.xlsx"
        build_workbook(orders, months).save(xlsx_path)
        declaration_path = DATA_DIR / f"{label}_declaration_tps_tvq.xlsx"
        build_declaration(orders, period, warnings).save(declaration_path)
        attachments += [xlsx_path, declaration_path]

    names = ", ".join(p.name for p in [json_path, *attachments])
    print(f"Wrote {len(orders)} orders to {names}")
    return attachments


def send_by_email(attachments: list[Path], order_count: int, period: str,
                  warnings: list[str] | None = None) -> None:
    smtp_host = os.environ.get("SMTP_HOST", "smtp.mail.yahoo.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "465"))
    username = env_or_die("SMTP_USERNAME")
    password = env_or_die("SMTP_PASSWORD")
    from_addr = os.environ.get("EMAIL_FROM", username)
    to_addr = os.environ.get("EMAIL_TO", "eladdas@yahoo.fr")

    date_str = attachments[0].stem.split("_")[0]
    msg = EmailMessage()
    flag = " - INCOMPLET" if warnings else ""
    msg["Subject"] = f"Extraction commandes Shopify (TPS/TVQ) - {date_str}{flag}"
    msg["From"] = from_addr
    msg["To"] = to_addr
    body = f"Ci-joint l'extraction des commandes du trimestre {period} ({order_count} commandes).\n"
    if warnings:
        body += "\nATTENTION - données incomplètes :\n" + "".join(f"  - {w}\n" for w in warnings)
    msg.set_content(body)
    for path in attachments:
        maintype, subtype = ("text", "csv") if path.suffix == ".csv" else (
            "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        msg.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)

    with smtplib.SMTP_SSL(smtp_host, smtp_port) as smtp:
        smtp.login(username, password)
        smtp.send_message(msg)

    print(f"Emailed {', '.join(p.name for p in attachments)} to {to_addr}")


def resolve_access_token(store_url: str) -> str:
    access_token = os.environ.get("SHOPIFY_ACCESS_TOKEN")
    if access_token:
        return access_token

    client_id = os.environ.get("SHOPIFY_CLIENT_ID")
    client_secret = os.environ.get("SHOPIFY_CLIENT_SECRET")
    if client_id and client_secret:
        return get_access_token_via_client_credentials(store_url, client_id, client_secret)

    print(
        "Missing credentials: set SHOPIFY_ACCESS_TOKEN, or both "
        "SHOPIFY_CLIENT_ID and SHOPIFY_CLIENT_SECRET",
        file=sys.stderr,
    )
    sys.exit(1)


def build_quarter(year: int, quarter: int, tz: ZoneInfo, now: datetime, fetch) -> None:
    start, end = quarter_bounds(year, quarter, tz)
    label = f"{year}-T{quarter}"
    period = f"{label} ({start:%Y-%m-%d} au {end - timedelta(days=1):%Y-%m-%d})"
    print(f"Trimestre : {period}")
    orders, warnings = collect_quarter(year, quarter, tz, now, fetch)
    for warning in warnings:
        print(f"ATTENTION : {warning}")
    attachments = write_outputs(orders, label, quarter_months(year, quarter), period, warnings)
    send_by_email(attachments, len(orders), period, warnings)


def main() -> None:
    store_url = env_or_die("SHOPIFY_STORE_URL")
    api_version = os.environ.get("SHOPIFY_API_VERSION", "2024-10")
    status = os.environ.get("SHOPIFY_ORDER_STATUS", "any")
    tz = ZoneInfo(os.environ.get("SHOPIFY_TIMEZONE") or "America/Toronto")
    mode = (os.environ.get("EXTRACTION") or "trimestre").strip().lower()
    now = datetime.now(tz)

    if urlparse(f"https://{store_url}").hostname != store_url:
        store_url = urlparse(store_url if "://" in store_url else f"https://{store_url}").hostname or store_url

    session = build_session(resolve_access_token(store_url))

    def fetch(start: datetime, end: datetime) -> list[dict]:
        return fetch_all_orders(store_url, session, api_version, status,
                                start.isoformat(), (end - timedelta(seconds=1)).isoformat())

    if mode == "mensuel":
        year, month = previous_month(now.date())
        save_month(year, month, fetch(*month_bounds(year, month, tz)))
        if month % 3 == 0:
            build_quarter(year, month // 3, tz, now, fetch)
        else:
            print("Pas de fin de trimestre : aucun fichier trimestriel ni courriel ce mois-ci.")
        return

    year, quarter = resolve_quarter(os.environ.get("TRIMESTRE", ""), now.date())
    build_quarter(year, quarter, tz, now, fetch)


if __name__ == "__main__":
    main()
