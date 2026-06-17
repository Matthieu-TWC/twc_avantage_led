{
    'name': 'TWC MRP BOM Import Wizard',
    'version': '1.0.0',
    'category': 'Manufacturing',
    'summary': 'Wizard pour importer des nomenclatures depuis un fichier Excel',
    'depends': ['mrp', 'product', 'uom'],
    'data': [
        'views/mrp_bom_import_wizard_views.xml',
    ],
    'installable': True,
    'application': False,
}
