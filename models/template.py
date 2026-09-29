import base64
import binascii

from odoo import api, fields, models, _
from odoo.exceptions import ValidationError

from ..core.template import read_template


class QuadraticTemplate(models.Model):
    _name = "cq.xlsx.template"
    _description = "Plantilla XLSX de conciliación cuadrática"
    _order = "sequence, id desc"
    _check_company_auto = True

    name = fields.Char("Nombre", required=True)
    active = fields.Boolean(default=True)
    sequence = fields.Integer("Prioridad", default=10, help="Se elige la menor prioridad; entre iguales, la plantilla más reciente.")
    company_id = fields.Many2one("res.company", required=True, default=lambda self: self.env.company, string="Empresa")
    account_id = fields.Many2one("account.account", required=True, check_company=True, string="Cuenta contable bancaria")
    file_data = fields.Binary("Plantilla XLSX", required=True, attachment=True)
    file_name = fields.Char("Nombre del archivo", required=True)

    @api.constrains("file_data", "file_name", "account_id", "company_id")
    def _check_template(self):
        for template in self:
            if template.company_id not in template.account_id.company_ids:
                raise ValidationError(_("La cuenta contable debe estar disponible en la empresa de la plantilla."))
            if not (template.file_name or "").lower().endswith(".xlsx"):
                raise ValidationError(_("Seleccione un archivo .xlsx con el formato de conciliación."))
            if not template.file_data:
                raise ValidationError(_("Adjunte la plantilla XLSX."))
            try:
                read_template(base64.b64decode(template.file_data, validate=True))
            except (ValueError, binascii.Error) as error:
                raise ValidationError(str(error)) from error
