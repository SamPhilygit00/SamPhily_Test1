# SamPhily_Test1

A small Snake game built with plain HTML, CSS, and JavaScript — no build step required.

## Play

Open `index.html` in a browser, or serve the folder locally:

```
python3 -m http.server 8000
```

Then visit `http://localhost:8000`.

Controls: arrow keys / WASD, or swipe on touch devices.

## Shopify orders extraction

`scripts/shopify_extract_orders.py` pulls the last 90 calendar days of orders
from a Shopify store via the Admin REST API, writes a JSON dump, a flattened
CSV summary, and a formatted "Suivi ventes" Excel workbook to `data/orders/`
named after today's date (`YYYY-MM-DD.json` / `.csv` / `.._suivi_ventes.xlsx`),
and emails the CSV and the xlsx.

The xlsx (built by `scripts/build_sales_tracking_xlsx.py`, which can also be
run standalone) reproduces a monthly-block tracking format: one 6-column
block per calendar month (Date, Commande, Mt brut, tvq, tps, Expedition),
one row per calendar day, a totals row per month, and a grand-total row for
the whole period. Orders on the same calendar day are summed onto that
day's single row.

### 1. Create a Shopify custom app

1. In your Shopify admin: **Settings → Apps and sales channels → Applications
   → Développer des applications**, which now redirects to the **Dev
   Dashboard**.
2. Create an app from the Dev Dashboard ("Démarrer depuis le Dev Dashboard"),
   grant the `read_orders` Admin API scope, and publish a version.
3. Install the app on your store (**Installations → Installer l'appli**).
4. Open **Paramètres de l'appli** and copy the **Client ID** and **Secret**
   (click the eye icon to reveal it). The script exchanges these for a fresh
   Admin API access token on every run via the client credentials grant, so
   there is no static token to copy/rotate.

### 2. Configure GitHub secrets

In this repository: **Settings → Secrets and variables → Actions**, add:

| Secret                  | Value                                    |
| ------------------------ | ----------------------------------------- |
| `SHOPIFY_STORE_URL`      | `your-store.myshopify.com`               |
| `SHOPIFY_CLIENT_ID`      | the Client ID from step 1                |
| `SHOPIFY_CLIENT_SECRET`  | the Client secret from step 1            |

(A static `SHOPIFY_ACCESS_TOKEN` secret also works if you have one from a
legacy custom app — the script prefers it over the client ID/secret pair
when both are set.)

Also add, for emailing the CSV and xlsx via Yahoo Mail SMTP:

| Secret          | Value                                              |
| ---------------- | --------------------------------------------------- |
| `SMTP_USERNAME`  | the sending Yahoo Mail address (e.g. `eladdas@yahoo.fr`) |
| `SMTP_PASSWORD`  | a Yahoo **app password** for that account (Yahoo Account → Security → Generate app password) — not the regular account password |

By default both files are sent from and to `eladdas@yahoo.fr`. Override with
the `EMAIL_FROM` / `EMAIL_TO` environment variables if needed.

### 3. Automated runs

`.github/workflows/shopify-extract.yml` runs the script every Monday at
23:30 UTC and commits any new/updated files under `data/orders/` back to
this repo. You can also trigger it manually from the **Actions** tab ("Run
workflow").

### Run locally

```
pip install -r scripts/requirements.txt
export SHOPIFY_STORE_URL="your-store.myshopify.com"
export SHOPIFY_CLIENT_ID="..."
export SHOPIFY_CLIENT_SECRET="..."
export SMTP_USERNAME="eladdas@yahoo.fr"
export SMTP_PASSWORD="..."
python scripts/shopify_extract_orders.py
```

Optional environment variables: `SHOPIFY_API_VERSION` (default `2024-10`),
`SHOPIFY_ORDER_STATUS` (`any`/`open`/`closed`/`cancelled`, default `any`),
`SMTP_HOST` (default `smtp.mail.yahoo.com`), `SMTP_PORT` (default `465`),
`EMAIL_FROM` (default `SMTP_USERNAME`), `EMAIL_TO` (default
`eladdas@yahoo.fr`).

