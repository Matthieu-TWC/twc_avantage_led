from odoo import api, fields, models


class MailActivityType(models.Model):
    _inherit = 'mail.activity.type'

    subtype_ids = fields.Many2many(
        'mail.activity.type',
        'mail_activity_type_subtype_rel',
        'parent_type_id',
        'subtype_id',
        string="Sous types d'activité",
    )

    parent_type_ids = fields.Many2many(
        'mail.activity.type',
        'mail_activity_type_subtype_rel',
        'subtype_id',
        'parent_type_id',
        string="Type(s) parent",
    )

    is_subtype = fields.Boolean(
        compute='_compute_is_subtype',
        store=True,
    )

    @api.depends('parent_type_ids')
    def _compute_is_subtype(self):
        for rec in self:
            rec.is_subtype = bool(rec.parent_type_ids)
