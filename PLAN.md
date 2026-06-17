# Plan — Module Odoo `mrp_bom_import`

## Objectif

Wizard dans le module MRP permettant d'importer des nomenclatures (BOMs) depuis le fichier Excel
formaté (`Nomenclature_mrp_bom_output.xlsx`). Pour chaque entité référencée, le module vérifie
si elle existe déjà avant de la créer, sans jamais produire de doublons.

**Fichier d'entrée attendu** : le fichier Excel produit par le script de conversion, qui contient
deux feuilles exploitées par le module :
- `Product Import` — liste des composants et produits finis
- `BOM Import` — les 33 nomenclatures avec leurs lignes

---

## 1. Structure du module

```
mrp_bom_import/
├── __manifest__.py
├── __init__.py
├── wizard/
│   ├── __init__.py
│   └── mrp_bom_import_wizard.py
├── views/
│   └── mrp_bom_import_wizard_views.xml
└── security/
    └── ir.model.access.csv
```

---

## 2. `__manifest__.py`

```python
{
    'name': 'MRP BOM Import Wizard',
    'version': '1.0.0',
    'category': 'Manufacturing',
    'summary': 'Wizard pour importer des nomenclatures depuis un fichier Excel',
    'depends': ['mrp', 'product', 'uom'],
    'data': [
        'security/ir.model.access.csv',
        'views/mrp_bom_import_wizard_views.xml',
    ],
    'installable': True,
    'application': False,
}
```

---

## 3. Modèle du wizard — `mrp_bom_import_wizard.py`

### 3.1 Déclaration du modèle

```python
class MrpBomImportWizard(models.TransientModel):
    _name = 'mrp.bom.import.wizard'
    _description = 'Import de nomenclatures depuis Excel'

    file_data = fields.Binary(string='Fichier Excel (.xlsx)', required=True, attachment=False)
    file_name = fields.Char(string='Nom du fichier')
    state      = fields.Selection(
        [('draft', 'Prêt'), ('done', 'Terminé')],
        default='draft'
    )
    result_log = fields.Text(string='Rapport d\'import', readonly=True)
```

### 3.2 Méthode principale `action_import`

Appelée par le bouton « Importer ». Orchestre les étapes dans l'ordre :

```
1. Décoder le fichier Binary → bytes → openpyxl Workbook
2. Appeler _import_products(wb)   → dict {display_name: product.template record}
3. Appeler _import_boms(wb, product_map)
4. Passer state à 'done', écrire le log, retourner l'action pour rester ouvert
```

Encapsuler l'ensemble dans un seul `self.env.cr.savepoint()` pour pouvoir rollback proprement
en cas d'erreur critique. Les erreurs non bloquantes (ligne inconnue, produit introuvable)
sont loggées dans `result_log` sans interrompre l'import.

---

## 4. Étape 1 — Import des produits (`_import_products`)

**Feuille lue** : `Product Import`

**Colonnes attendues (row 1 = headers)** :

| Col | Nom attendu | Usage |
|-----|-------------|-------|
| A | `id` | Ignoré (external ID géré par Odoo natif, pas utilisé ici) |
| B | `name` | Nom du produit |
| C | `default_code` | Référence interne (peut être vide) |
| D | `standard_price` | Prix de revient |
| E | `weight` | Poids |
| F | `uom_id` | Unité de mesure (`Unité(s)` ou `m`) |

### Logique par ligne

```
Pour chaque ligne de données (à partir de la ligne 2) :

1. Lire name, default_code, standard_price, weight, uom_id

2. Rechercher le product.template existant :
   - Si default_code non vide :
       chercher product.template où default_code = valeur (limite 1)
   - Sinon :
       chercher product.template où name = valeur (limite 1)

3. Si trouvé → utiliser l'enregistrement existant (pas de modification)

4. Si non trouvé → créer product.template avec :
       name           = name
       default_code   = default_code (si non vide)
       standard_price = standard_price (si non vide)
       weight         = weight (si non vide)
       uom_id         = _resolve_uom(uom_id_str)
       type           = 'product'   (Article stockable)

5. Construire le display_name local :
       "[default_code] name"  si default_code existe
       "name"                 sinon

6. Ajouter au dictionnaire product_map[display_name] = record
   (Ajouter aussi product_map[name] = record comme clé de fallback)

Retourner product_map
```

### Méthode utilitaire `_resolve_uom(uom_str)`

```
- Si uom_str == 'm'          → chercher uom.uom où name = 'm' (ou 'Mètre')
- Si uom_str == 'Unité(s)'   → chercher uom.uom où name ilike 'unit' ou 'unité'
- Fallback                   → uom.uom ref 'uom.product_uom_unit'
```

---

## 5. Étape 2 — Import des BOMs (`_import_boms`)

**Feuille lue** : `BOM Import`

**Colonnes attendues** :

| Col | Nom attendu | Usage |
|-----|-------------|-------|
| A | `id` | Ignoré |
| B | `product_tmpl_id` | Nom du produit fini (non vide = début d'une nouvelle BOM) |
| C | `code` | Référence de la nomenclature |
| D | `type` | Toujours `'Fabriquer ce produit'` → `'normal'` |
| E | `bom_line_ids/product_id` | Display name du composant |
| F | `bom_line_ids/product_qty` | Quantité |
| G | `note` | Notes supplémentaires (ignorées pour l'import BOM) |

### Logique de lecture

Les lignes sont groupées : quand la colonne B est non vide, c'est le début d'une nouvelle BOM.
Les lignes suivantes (col B vide) sont des lignes de composants de la BOM courante.

```
current_bom = None
current_lines = []

Pour chaque ligne de données :
    Si col B non vide :
        → Sauvegarder la BOM précédente (si elle existe)
        → Démarrer une nouvelle BOM : product_name = col B, code = col C
        → current_lines = []
    
    Si col E non vide :
        → Ajouter {component_display: col E, qty: col F} à current_lines

À la fin de la boucle : sauvegarder la dernière BOM
```

### Méthode `_save_bom(product_name, code, lines, product_map)`

```
1. Résoudre le produit fini :
   - Chercher dans product_map[product_name]
   - Sinon chercher product.template où name = product_name
   - Si toujours introuvable : logger l'erreur, passer à la suivante

2. Vérifier si la BOM existe déjà :
   - Chercher mrp.bom où product_tmpl_id = produit ET code = code (si code non vide)
   - OU mrp.bom où product_tmpl_id = produit (si pas de code)
   - Si trouvée → SKIP (logger "BOM déjà existante, ignorée")

3. Si non trouvée → créer mrp.bom :
       product_tmpl_id = produit fini
       code            = code (si non vide)
       type            = 'normal'

4. Pour chaque ligne dans lines :
       a. Résoudre le composant :
              - Chercher product_map[component_display]         (correspondance exacte)
              - Sinon extraire le default_code de "[ref] name"
                et chercher product.template où default_code = ref
              - Sinon chercher product.template où name ilike nom extrait
              - Si introuvable : logger l'avertissement, skip la ligne

       b. Créer mrp.bom.line :
              bom_id      = bom créée en étape 3
              product_id  = product.product par défaut du template trouvé
                            (template.product_variant_ids[0])
              product_qty = qty (float, défaut 1.0 si vide)
```

---

## 6. Vues — `mrp_bom_import_wizard_views.xml`

### Form view du wizard

```xml
<record id="view_mrp_bom_import_wizard_form" model="ir.ui.view">
  <field name="name">mrp.bom.import.wizard.form</field>
  <field name="model">mrp.bom.import.wizard</field>
  <field name="arch" type="xml">
    <form string="Import de nomenclatures">

      <div class="alert alert-info" attrs="{'invisible': [('state','=','done')]}">
        Importer le fichier Excel de nomenclatures
        (feuilles <b>Product Import</b> et <b>BOM Import</b>).
        Les produits et BOMs existants ne seront pas modifiés.
      </div>

      <group attrs="{'invisible': [('state','=','done')]}">
        <field name="file_data" filename="file_name" widget="binary"/>
        <field name="file_name" invisible="1"/>
      </group>

      <group attrs="{'invisible': [('state','!=','done')]}">
        <field name="result_log" widget="text" readonly="1"
               style="font-family: monospace; font-size: 12px;"/>
      </group>

      <footer>
        <button name="action_import" type="object" string="Importer"
                class="btn-primary"
                attrs="{'invisible': [('state','=','done')]}"/>
        <button string="Fermer" class="btn-secondary" special="cancel"/>
      </footer>

    </form>
  </field>
</record>
```

### Action du wizard

```xml
<record id="action_mrp_bom_import_wizard" model="ir.actions.act_window">
  <field name="name">Importer des nomenclatures</field>
  <field name="res_model">mrp.bom.import.wizard</field>
  <field name="view_mode">form</field>
  <field name="target">new</field>
</record>
```

### Ajout dans le menu MRP

```xml
<menuitem
  id="menu_mrp_bom_import"
  name="Importer des nomenclatures"
  parent="mrp.menu_mrp_bom"
  action="action_mrp_bom_import_wizard"
  sequence="99"/>
```

---

## 7. Sécurité — `ir.model.access.csv`

```csv
id,name,model_id:id,group_id:id,perm_read,perm_write,perm_create,perm_unlink
access_mrp_bom_import_wizard,mrp.bom.import.wizard,model_mrp_bom_import_wizard,mrp.group_mrp_manager,1,1,1,1
```

---

## 8. Gestion des erreurs et log

Le `result_log` doit être lisible par un non-développeur. Format suggéré :

```
✔ Produits importés    : 87 créés, 13 déjà existants
✔ BOMs importées       : 28 créées, 5 déjà existantes
⚠ Avertissements (3) :
   - Ligne 45 : composant "[28999] Produit inconnu" introuvable, ligne ignorée
   - Ligne 132 : BOM "Horus G.E 1200 DALI" déjà existante, ignorée
   - Ligne 201 : unité de mesure "pcs" non reconnue, Unité(s) utilisée par défaut
```

---

## 9. Points d'attention pour l'implémentation

1. **openpyxl** — disponible dans l'environnement Odoo standard (utilisé par les exports natifs).
   Import : `from openpyxl import load_workbook`. Lire avec `data_only=True` pour résoudre
   les formules Excel.

2. **product.product vs product.template** — Les lignes de BOM (`mrp.bom.line`) référencent
   `product_id` (product.product). Récupérer la variante par défaut d'un template via
   `template.product_variant_ids[0]` ou `template.product_variant_id`.

3. **UoM** — La correspondance `Unité(s)` → enregistrement `uom.uom` peut varier selon la
   langue installée. Faire une recherche `ilike` plutôt qu'une égalité stricte, ou utiliser
   les `xml_id` Odoo (`uom.product_uom_unit`, `uom.product_uom_meter`).

4. **Transactions** — Utiliser `with self.env.cr.savepoint():` autour de chaque BOM (pas de
   l'import entier) pour qu'une BOM en erreur n'annule pas toutes les précédentes.

5. **Performances** — Précharger tous les `product.template` par `default_code` et par `name`
   en une seule requête `search` avant la boucle, plutôt que d'interroger la BDD ligne par ligne.

6. **Test** — Tester d'abord avec `raise UserError(str(product_map))` après `_import_products`
   pour valider que les 124 produits sont correctement résolus avant d'attaquer les BOMs.
