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

`scripts/shopify_extract_orders.py` pulls all orders from a Shopify store via
the Admin REST API and writes them to `data/orders/` as both a timestamped
JSON dump, a flattened CSV summary, and `orders_latest.{json,csv}` copies.

### 1. Create a Shopify custom app

1. In your Shopify admin: **Settings → Apps and sales channels → Develop apps**.
2. Create an app, then under **Configuration** grant the `read_orders` Admin
   API access scope.
3. Install the app and copy the generated **Admin API access token**.

### 2. Configure GitHub secrets

In this repository: **Settings → Secrets and variables → Actions**, add:

| Secret                 | Value                                  |
| ----------------------- | --------------------------------------- |
| `SHOPIFY_STORE_URL`     | `your-store.myshopify.com`             |
| `SHOPIFY_ACCESS_TOKEN`  | the Admin API access token from step 1 |

### 3. Automated runs

`.github/workflows/shopify-extract.yml` runs the script daily at 06:00 UTC
and commits any new/updated files under `data/orders/` back to this repo. You
can also trigger it manually from the **Actions** tab ("Run workflow").

### Run locally

```
pip install -r scripts/requirements.txt
export SHOPIFY_STORE_URL="your-store.myshopify.com"
export SHOPIFY_ACCESS_TOKEN="shpat_..."
python scripts/shopify_extract_orders.py
```

Optional environment variables: `SHOPIFY_API_VERSION` (default `2024-10`),
`SHOPIFY_ORDER_STATUS` (`any`/`open`/`closed`/`cancelled`, default `any`),
`SHOPIFY_UPDATED_AT_MIN` (ISO 8601 timestamp to only fetch recently updated
orders).

