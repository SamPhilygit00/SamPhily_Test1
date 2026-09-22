# Récapitulatif — Extraction automatisée des commandes Shopify

Résumé de l'automatisation mise en place (session du 22 septembre 2026).

## Ce que ça fait

Chaque lundi à 23:30 UTC (ou sur déclenchement manuel), le workflow GitHub
Actions **"Shopify orders extraction"** :

1. Récupère toutes les commandes des **90 derniers jours calendrier**
   (`created_at_min`) via l'API Admin de Shopify.
2. Écrit dans `data/orders/`, nommés par la date du jour :
   - `YYYY-MM-DD.json` — dump complet
   - `YYYY-MM-DD.csv` — résumé aplati
   - `YYYY-MM-DD_suivi_ventes.xlsx` — fichier "Suivi ventes" au format modèle
3. Envoie le **CSV et le xlsx** par courriel à **eladdas@yahoo.fr**.
4. Commit et pousse les fichiers générés dans le dépôt (`main`).

## Colonnes du CSV

`id, order_number, created_at, updated_at, cancelled_at, financial_status,
fulfillment_status, currency, subtotal_price, total_discounts, TPS, TVQ,
frais_livraison, total_price, shipping_city`

- **TPS** / **TVQ** : taxes fédérale (GST, 5%) et provinciale (QST, ~9.975%)
  séparées à partir des `tax_lines` de chaque commande.
- **frais_livraison** : `total_shipping_price_set` de la commande.

## Fichier "Suivi ventes" (xlsx)

Généré par `scripts/build_sales_tracking_xlsx.py` (utilisable aussi en
standalone), il reproduit le format du fichier modèle fourni :

- Un bloc de 6 colonnes par mois calendaire (Date, Commande, Mt brut, tvq,
  tps, Expedition), blocs disposés côte à côte de gauche à droite.
- Une ligne par jour calendaire ; si plusieurs commandes tombent le même
  jour, leurs montants sont additionnés sur cette ligne et les numéros de
  commande sont joints par `, `.
- Une ligne de total par mois (calculée directement en Python, pas une
  formule Excel — donc toujours correcte à l'ouverture).
- Une ligne **"TOTAL PÉRIODE"** en bas, sommant les totaux de tous les mois.
- Format numérique forcé au point décimal (`0"."00`), peu importe la
  configuration régionale d'Excel.
- Mise en forme (couleurs, bordures, largeurs de colonnes) reproduite à
  l'identique du modèle fourni.

## Comment c'est authentifié

Une app privée Shopify a été créée depuis le **Dev Dashboard** (Settings →
Applications → Développer des applications → Dev Dashboard), avec le scope
`read_orders`, installée sur la boutique **LBL extensions**
(`4xnusa-is.myshopify.com`). Le script échange le Client ID/Secret de cette
app contre un token d'accès frais à chaque exécution (grant client
credentials) — pas de token statique à faire tourner.

## Secrets GitHub configurés

Dans **Settings → Secrets and variables → Actions** :

| Secret | Contenu |
|---|---|
| `SHOPIFY_STORE_URL` | `4xnusa-is.myshopify.com` |
| `SHOPIFY_CLIENT_ID` | Client ID de l'app Shopify |
| `SHOPIFY_CLIENT_SECRET` | Secret de l'app Shopify |
| `SMTP_USERNAME` | `eladdas@yahoo.fr` |
| `SMTP_PASSWORD` | mot de passe d'application Yahoo Mail |

## Fichiers du projet

- `scripts/shopify_extract_orders.py` — le script d'extraction/email
- `scripts/build_sales_tracking_xlsx.py` — génère le fichier "Suivi ventes"
- `scripts/requirements.txt` — dépendances Python (`requests`, `openpyxl`)
- `.github/workflows/shopify-extract.yml` — planification + déclenchement manuel
- `data/orders/` — fichiers générés (JSON + CSV + xlsx par date d'exécution)
- `README.md` — instructions de configuration détaillées

## Déclencher un run manuellement

1. Dépôt GitHub → onglet **Actions**
2. Workflow **"Shopify orders extraction"** dans la liste de gauche
3. Bouton **"Run workflow"** → confirmer sur la branche `main`

## Historique de cette session

- Création du script d'extraction + workflow (commandes uniquement)
- Passage du planning quotidien à hebdomadaire (lundi 23:30 UTC)
- Migration de l'authentification vers le flux Dev Dashboard (Client ID/Secret)
- Séparation TPS/TVQ, ajout des frais de livraison, retrait des colonnes
  client/tags/pays
- Passage à une fenêtre glissante de 90 jours + noms de fichiers `YYYY-MM-DD`
- Ajout de l'envoi automatique du CSV par courriel (Yahoo Mail SMTP)
- Création du fichier "Suivi ventes" (xlsx) reproduisant le format modèle
  fourni, avec total par mois et total général, intégré à l'automatisation
  hebdomadaire et envoyé par courriel en plus du CSV
