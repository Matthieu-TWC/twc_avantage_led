from odoo import fields, models, api


class MailActivitySchedule(models.TransientModel):
    _inherit = 'mail.activity.schedule'

    activity_subtype_id = fields.Many2one(
        'mail.activity.type',
        string="Sous type d'activité",
    )

    available_subtype_ids = fields.Many2many(
        'mail.activity.type',
        compute='_compute_subtype_info',
    )

    is_container_type = fields.Boolean(
        compute='_compute_subtype_info',
    )

    @api.depends('activity_type_id')
    def _compute_subtype_info(self):
        for rec in self:
            subtypes = rec.activity_type_id.subtype_ids
            rec.available_subtype_ids = subtypes
            rec.is_container_type = bool(subtypes)

    @api.onchange('activity_subtype_id')
    def _onchange_activity_subtype_id(self):
        subtype = self.activity_subtype_id
        if not subtype:
            return
        self.summary = subtype.summary
        self.note = subtype.default_note
        self.date_deadline = subtype._get_date_deadline()
        self.activity_user_id = subtype.default_user_id or self.env.user
