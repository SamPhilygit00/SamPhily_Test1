#!/usr/bin/env python3
"""Extract orders from a Shopify store via the Admin REST API.

Writes both a full JSON dump and a flattened CSV summary to data/orders/,
plus stable "latest" copies for easy downstream consumption.

Required environment variables:
  SHOPIFY_STORE_URL      e.g. "my-store.myshopify.com"

  Authentication - provide EITHER of these:
  SHOPIFY_ACCESS_TOKEN   a static Admin API access token (legacy custom apps)
  OR
  SHOPIFY_CLIENT_ID      Client ID from a Dev Dashboard custom app
  SHOPIFY_CLIENT_SECRET  Client secret from a Dev Dashboard custom app
                         (a fresh access token is requested via the client
                         credentials grant on every run)

Optional environment variables:
  SHOPIFY_API_VERSION   default "2024-10"
  SHOPIFY_ORDER_STATUS  default "any" (any|open|closed|cancelled)
  SHOPIFY_UPDATED_AT_MIN  ISO 8601 timestamp to only fetch orders updated since
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

import requests

DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "orders"
PAGE_LIMIT = 250
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


def fetch_all_orders(store_url: str, session: requests.Session, api_version: str,
                      status: str, updated_at_min: str | None) -> list[dict]:
    base = f"https://{store_url}/admin/api/{api_version}/orders.json"
    params = {"limit": PAGE_LIMIT, "status": status}
    if updated_at_min:
        params["updated_at_min"] = updated_at_min

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


def write_outputs(orders: list[dict]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    json_path = DATA_DIR / f"orders_{timestamp}.json"
    json_path.write_text(json.dumps(orders, indent=2), encoding="utf-8")

    rows = [flatten_order(o) for o in orders]
    csv_path = DATA_DIR / f"orders_{timestamp}.csv"
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

    latest_json = DATA_DIR / "orders_latest.json"
    latest_csv = DATA_DIR / "orders_latest.csv"
    latest_json.write_text(json_path.read_text(encoding="utf-8"), encoding="utf-8")
    latest_csv.write_text(csv_path.read_text(encoding="utf-8"), encoding="utf-8")

    print(f"Wrote {len(orders)} orders to {json_path.name} / {csv_path.name}")
    print(f"Updated {latest_json.name} / {latest_csv.name}")


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


def main() -> None:
    store_url = env_or_die("SHOPIFY_STORE_URL")
    api_version = os.environ.get("SHOPIFY_API_VERSION", "2024-10")
    status = os.environ.get("SHOPIFY_ORDER_STATUS", "any")
    updated_at_min = os.environ.get("SHOPIFY_UPDATED_AT_MIN") or None

    if urlparse(f"https://{store_url}").hostname != store_url:
        store_url = urlparse(store_url if "://" in store_url else f"https://{store_url}").hostname or store_url

    access_token = resolve_access_token(store_url)
    session = build_session(access_token)
    orders = fetch_all_orders(store_url, session, api_version, status, updated_at_min)
    write_outputs(orders)


if __name__ == "__main__":
    main()
