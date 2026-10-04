# Récapitulatif — Extraction automatisée des commandes Shopify

Résumé de l'automatisation mise en place (session du 22 septembre 2026).

## Ce que ça fait

Shopify ne donne accès qu'aux **60 derniers jours** de commandes. Pour qu'un
trimestre soit toujours complet, l'extraction se fait donc **chaque mois** :

- **Le 1er de chaque mois à 05:30 UTC** (premier lancement le **1er novembre
  2026**), le workflow GitHub Actions **"Shopify orders extraction"**
  récupère le mois qui vient de se terminer et l'enregistre dans
  `data/orders/mensuel/AAAA-MM.json`. Ces fichiers restent dans le dépôt,
  même si Shopify change ses règles plus tard.
- **Les 1er janvier, avril, juillet et octobre**, il assemble en plus les 3
  mois du trimestre terminé (ex. le 1er janvier 2027 : octobre, novembre et
  décembre 2026) et :

1. Calcule les commandes du **trimestre** (heure du Québec) à partir des 3
   fichiers mensuels, pour la déclaration TPS/TVQ.
2. Écrit dans `data/orders/`, nommés par le trimestre :
   - `AAAA-TN.json` — dump complet (ex. `2026-T4.json`)
   - `AAAA-TN.csv` — résumé aplati
   - `AAAA-TN_suivi_ventes.xlsx` — fichier "Suivi ventes" au format modèle
   - `AAAA-TN_declaration_tps_tvq.xlsx` — fichier de déclaration TPS/TVQ
3. Envoie le **CSV et les deux xlsx** par courriel à **eladdas@yahoo.fr**.
4. Commit et pousse les fichiers générés dans le dépôt (`main`).

## Fichier "Déclaration TPS/TVQ" (xlsx)

Généré par `scripts/build_tax_declaration_xlsx.py` (utilisable aussi en
standalone), il reproduit la « Section A — Revenus » :

| Ligne | Case | Montant HT (Brut) | TPS perçue (E) | TVQ perçue (F) |
|---|---|---|---|---|
| 11 | 101 — Vente en ligne (Shopify) | total des ventes HT | ventes × 5 % (indicatif) | ventes × 9,975 % (indicatif) |
| 12 | 105/205 — Taxes perçues (clients) | ventes sur lesquelles une taxe a été perçue | **= tps du TOTAL PÉRIODE** | **= tvq du TOTAL PÉRIODE** |

E12 et F12 sont calculés exactement comme la ligne « TOTAL PÉRIODE » du
fichier Suivi ventes : les deux fichiers concordent toujours au cent près.
Les taxes perçues incluent les taxes sur les frais de livraison, d'où un
écart normal avec « montant × taux ».

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
- Une ligne par jour ayant au moins une commande (les jours sans commande
  ne sont pas affichés) ; si plusieurs commandes tombent le même jour, leurs
  montants sont additionnés sur cette ligne et les numéros de commande sont
  joints par `, `.
- Une ligne de total par mois (calculée directement en Python, pas une
  formule Excel — donc toujours correcte à l'ouverture).
- Une ligne **"TOTAL PÉRIODE"** en bas, sommant les totaux de tous les mois.
- Format numérique à deux décimales (`0.00`) : 441 s'affiche `441.00`
  (`441,00` dans un Excel en français).
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

## Données incomplètes

Si un mois du trimestre n'a pas de fichier mensuel et a commencé il y a plus
de 60 jours, il ne peut être récupéré qu'en partie : le courriel porte alors
la mention **INCOMPLET** et la déclaration affiche un avertissement en rouge
(cellule C4). Août et septembre 2026 ont été enregistrés à partir des
extractions précédentes ; juillet 2026 n'est que partiel (depuis le 24 juillet).

## Déclencher un run manuellement

1. Dépôt GitHub → onglet **Actions**
2. Workflow **"Shopify orders extraction"** dans la liste de gauche
3. Bouton **"Run workflow"** → champ **trimestre** : vide = dernier
   trimestre terminé, `actuel` = trimestre en cours, ou ex. `2026-T3` →
   confirmer sur la branche `main`

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
