# Dépenses — documents pour la déclaration TPS/TVQ

Déposez ici vos factures, reçus et relevés de frais, **dans le dossier du trimestre** :

```
depenses/2026-T3/   juillet, août, septembre 2026
depenses/2026-T4/   octobre, novembre, décembre 2026
depenses/2027-T1/   janvier, février, mars 2027   (créer le dossier au besoin)
```

Dès qu'un document est déposé (ou supprimé), GitHub recalcule automatiquement
la déclaration du trimestre et l'envoie par courriel :

- **TPS perçue** des clients (Shopify) − **TPS payée** sur les dépenses
- **TVQ perçue** des clients (Shopify) − **TVQ payée** sur les dépenses
- résultat > 0 : montant **à payer à Revenu Québec** ;
  résultat < 0 : montant **à réclamer** (remboursement).

## Formats acceptés

| Format | Comment c'est lu |
|---|---|
| Excel (`.xlsx`) ou `.csv` | Directement, ligne par ligne. Il faut une colonne **TPS** et une colonne **TVQ** ; les colonnes Date, Fournisseur, Description, N° facture, Montant HT et Total sont utilisées si présentes. |
| PDF (`.pdf`) | Par Claude (IA), qui relève la TPS/TVH et la TVQ imprimées sur chaque facture. |
| Photos (`.jpg`, `.png`, `.heic`…) | Par Claude (IA), comme les PDF. |

Les PDF et photos demandent le secret GitHub `ANTHROPIC_API_KEY`
(Settings → Secrets and variables → Actions). Chaque document n'est lu qu'une
fois : le résultat est gardé dans `data/depenses/cache.json`.

## Vérifier le résultat

Le fichier `data/orders/AAAA-TN_declaration_tps_tvq.xlsx` (aussi envoyé par
courriel) contient :

- **Section B — Dépenses** : total des taxes payées ;
- **Section C — Solde** : le calcul et le message à payer / à réclamer ;
- l'onglet **Dépenses** : une ligne par facture, avec le fichier d'origine et
  les remarques (doublon possible, devise étrangère, date hors trimestre…).

Un document placé directement dans `depenses/` (sans dossier de trimestre) est
réparti selon la date de chaque dépense.
