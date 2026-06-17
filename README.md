# odoo19-dev — Custom Addons

Custom Odoo 19 modules for the TWC project. All modules follow the `twc_` prefix convention.

## Stack

The full stack runs via Docker Compose from the parent directory:

```bash
# Start
docker-compose up -d

# Restart Odoo after code changes
docker-compose restart odoo

# Install a module
docker-compose exec odoo odoo -d <dbname> -i <module_name> --stop-after-init

# Upgrade a module
docker-compose exec odoo odoo -d <dbname> -u <module_name> --stop-after-init

# Logs
docker-compose logs -f odoo
```

This `addons/` directory is mounted at `/mnt/extra-addons` inside the container.

## Modules

### `twc_mrp_bom_import`

Wizard allowing users to import Bills of Materials from an Excel file directly inside Odoo.

**Entry point:** Manufacturing → Bills of Materials → Importer des nomenclatures

**Excel format:** see `Nomenclature_mrp_bom_output_14.xlsx` as reference, and `GUIDE_IMPORT_BOM.md` at the root of the project for the full user guide.

## Module structure

```
twc_<module>/
├── __manifest__.py
├── __init__.py
├── wizard/          # TransientModel classes
├── views/           # XML views, actions, menu items
└── security/
    └── ir.model.access.csv
```
