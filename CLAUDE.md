# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Environment

This is a custom Odoo 19 addons directory. The full stack runs via Docker Compose from the parent directory (`/home/aranorn/odoo19-dev/`):

```bash
# Start the stack
cd /home/aranorn/odoo19-dev && docker-compose up -d

# Restart Odoo only (after code changes)
docker-compose restart odoo

# Install or upgrade a module
docker-compose exec odoo odoo -d <dbname> -i <module_name> --stop-after-init
docker-compose exec odoo odoo -d <dbname> -u <module_name> --stop-after-init

# Tail Odoo logs
docker-compose logs -f odoo
```

The `./addons/` directory is mounted at `/mnt/extra-addons` inside the container.

## Module conventions

All modules in this repo follow the `twc_` prefix convention (e.g., `twc_mrp_bom_import`). Each module has the standard Odoo structure:

```
twc_<module>/
├── __manifest__.py       # name, version, depends, data list
├── __init__.py
├── wizard/               # TransientModel classes for one-shot operations
├── views/                # XML: form views, actions, menuitems
└── security/
    └── ir.model.access.csv
```

## Current modules

### `twc_mrp_bom_import`

A one-shot import wizard that reads a two-sheet Excel file and creates `product.template` + `mrp.bom` + `mrp.bom.line` records without producing duplicates.

**Entry point:** `Manufacturing → Bills of Materials → Importer des nomenclatures`

**Excel format required:**
- Sheet `Product Import`: columns id, name, default_code, standard_price, weight, uom_id
- Sheet `BOM Import`: columns id, product_tmpl_id, code, type, bom_line_ids/product_id, bom_line_ids/product_qty, note — rows with a non-empty column B start a new BOM; following rows (col B empty) are its component lines

**Key design decisions:**
- Products are pre-loaded in bulk before the loop (`search([])` then dict by `default_code` / `name`) to avoid N+1 queries.
- Each BOM is wrapped in its own `self.env.cr.savepoint()` so a failure on one BOM does not roll back the others.
- `mrp.bom.line.product_id` requires a `product.product` variant — always resolved via `template.product_variant_id`.
- Deduplication: existing products are never modified; existing BOMs (matched by `product_tmpl_id` + `code`) are skipped with a log entry.
- UoM resolution uses Odoo XML IDs (`uom.product_uom_unit`, `uom.product_uom_meter`) with `ilike` fallbacks for locale variation.
- Access restricted to `mrp.group_mrp_manager` via `ir.model.access.csv`.

## Odoo-specific patterns to follow

- TransientModels (`models.TransientModel`) for wizards; persistent models (`models.Model`) for stored data.
- Always use `self.ensure_one()` at the top of button handler methods.
- Return an `ir.actions.act_window` dict from wizard button methods to keep the dialog open after processing.
- Use `self.env.ref('xml_id', raise_if_not_found=False)` when resolving external IDs that may not exist.
- Menu items referencing standard Odoo menus must use their full XML ID (e.g., `mrp.menu_mrp_bom`).
- Use `invisible="state == 'done'"` directly on elements for conditional visibility — `attrs=` was removed in Odoo 17 and will cause a `ParseError` on install.
