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

## Plusieurs factures dans un seul PDF

Un PDF peut contenir **plusieurs factures à la suite** (ex. un lot de reçus
scannés) : chacune est extraite séparément, comme si elle avait été déposée
seule. C'est la façon recommandée de déposer un trimestre complet d'un coup.

GitHub limite le téléversement web à 25 Mo et le chat à 30 Mo : si votre PDF
dépasse cette taille, divisez-le en 2-3 fichiers plus petits (par plage de
pages, sans tout fusionner) avec un outil comme
[ilovepdf.com/split_pdf](https://www.ilovepdf.com/split_pdf) — avoir
plusieurs PDF au lieu d'un seul ne change rien au résultat.

## Repère manuscrit pour calculer une taxe

Si vous annotez une facture à la main et que vous ne connaissez pas le
détail TPS/TVQ (ex. un reçu qui n'affiche qu'une taxe combinée ou aucune
taxe du tout), écrivez **« TPS - ? »** et/ou **« TVQ - ? »** à côté du
montant. C'est le seul repère manuscrit pris en compte : il indique que la
taxe doit être **calculée** plutôt que laissée à 0. Toute autre annotation
manuscrite (catégorie, flèches, notes) est ignorée pour les montants.

## Catégories

Chaque dépense est classée automatiquement selon le nom du fournisseur :

| Fournisseur | Catégorie |
|---|---|
| Amazon, MM, tout fournisseur non listé | Matériel de cils |
| DHL | Frais de livraison |
| Scores, Maxi, Dollarama, H&M, Dollar Max Dépôt, Michaels | Frais de maintien divers |
| La Perle Virtuelle | Soutien technique |
| Zahka Inc | Service d'accompagnement en ligne de marque |
| Telus | Frais Téléphonie et Prêt |

La catégorie apparaît dans l'onglet Dépenses, à côté du fournisseur.

## Factures Telus (cellulaire de l'entreprise)

- **Facture mensuelle Telus** : déposer **une seule facture par trimestre**
  (ex. celle de septembre). Le même montant étant payé chaque mois, son
  montant HT, sa TPS, sa TVQ et son total sont **multipliés par 3**. Si la
  facture n'affiche pas de TPS/TVQ, elles sont d'abord calculées à partir
  du total. Si plusieurs factures mensuelles Telus sont déposées dans le
  même trimestre, aucune n'est multipliée (avertissement dans le courriel).
- **Prêt téléphone Telus** (mettre « prêt », « financement » ou
  « appareil » dans le nom du fichier, ex. `PRET TELEPHONE.pdf`) : la taxe
  n'a été payée qu'une fois (juillet 2026), elle est comptée **telle
  quelle, sans × 3**.

## Factures « taxes incluses »

Si une facture n'affiche qu'un **total taxes incluses** (ou « TTC »), sans
montants de TPS/TVQ, ils sont calculés à partir de ce total :

- montant HT = total ÷ 1,14975
- TPS = 5 % du montant HT
- TVQ = le reste (HT + TPS + TVQ = total, au cent près)

La ligne est marquée « taxes calculées » dans l'onglet Dépenses. Pour forcer
ce calcul sur un document qui ne porte pas la mention (ex. un reçu DHL),
mettez **`taxes-incluses`** ou **`TTC`** dans le nom du fichier :
`dhl_taxes-incluses.jpg`. Dans un tableau Excel/CSV nommé ainsi, le calcul
s'applique aux lignes qui ont un total mais pas de TPS/TVQ.

## Factures avec une taxe combinée (non détaillée)

Beaucoup de factures (Amazon « Taxes de vente », Shopify « Taxes estimées »)
affichent **un seul montant de taxe**, sans distinguer TPS et TVQ. Dans ce
cas, le montant est réparti selon le même ratio que les taux en vigueur
(5 % / 9,975 %) et la ligne est marquée « taxe combinée répartie ».

## Vérifier le résultat

Le fichier `data/orders/AAAA-TN_declaration_tps_tvq.xlsx` (aussi envoyé par
courriel) contient :

- **Section B — Dépenses** : total des taxes payées ;
- **Section C — Solde** : le calcul et le message à payer / à réclamer ;
- l'onglet **Dépenses** : une ligne par facture, avec le fichier d'origine et
  les remarques (doublon possible, devise étrangère, date hors trimestre…).

Un document placé directement dans `depenses/` (sans dossier de trimestre) est
réparti selon la date de chaque dépense.
