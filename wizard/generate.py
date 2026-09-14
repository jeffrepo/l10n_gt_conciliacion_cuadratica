import base64
from datetime import timedelta
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

from odoo import api, fields, models, _
from odoo.exceptions import AccessError, UserError, ValidationError

from ..core.calculation import MONTHS


class QuadraticGenerateWizard(models.TransientModel):
    _name = "cq.generate.wizard"
    _description = "Generar conciliación cuadrática"
    _check_company_auto = True

    company_id = fields.Many2one("res.company", required=True, default=lambda self: self.env.company, string="Empresa")
    year = fields.Integer("Año", required=True, default=lambda self: (fields.Date.context_today(self).replace(day=1) - timedelta(days=1)).year)
    month = fields.Selection([(str(index + 1), name) for index, name in enumerate(MONTHS)], required=True,
                             default=lambda self: str((fields.Date.context_today(self).replace(day=1) - timedelta(days=1)).month), string="Mes de corte")
    all_journals = fields.Boolean("Todas las cuentas habilitadas")
    journal_ids = fields.Many2many("account.journal", string="Cuentas bancarias", check_company=True)
    state = fields.Selection([("select", "Selección"), ("result", "Resultados")], default="select")
    report_ids = fields.Many2many("cq.report", string="Resultados", readonly=True)
    file_data = fields.Binary("Archivo", readonly=True)
    file_name = fields.Char("Nombre de archivo", readonly=True)

    @api.onchange("company_id")
    def _onchange_company(self):
        self.journal_ids = False
        self.report_ids = False
        self.state = "select"

    @api.constrains("year")
    def _check_year(self):
        if any(not 1900 <= wizard.year <= 9998 for wizard in self):
            raise ValidationError(_("El año debe estar entre 1900 y 9998."))

    def action_calculate(self):
        self.ensure_one()
        self.check_access("write")
        if not self.env.user.has_group("account.group_account_user"):
            raise AccessError(_("Se requiere acceso de Contabilidad para generar este reporte."))
        if self.company_id.id not in self.env.companies.ids:
            raise AccessError(_("La empresa seleccionada no está permitida en la sesión."))
        journals = self.journal_ids
        if self.all_journals:
            journals = self.env["account.journal"].search([
                ("company_id", "=", self.company_id.id), ("type", "=", "bank"), ("cq_enabled", "=", True),
            ])
        if not journals:
            raise UserError(_("Seleccione cuentas bancarias o habilítelas en la configuración de diarios."))
        reports = self.env["cq.report"]
        for journal in journals:
            reports |= self.env["cq.report"]._generate(self.company_id, journal, self.year, int(self.month))
        self.write({"report_ids": [(6, 0, reports.ids)], "state": "result", "file_data": False})
        return {"type": "ir.actions.act_window", "name": _("Conciliación cuadrática"),
                "res_model": self._name, "res_id": self.id, "view_mode": "form", "target": "new"}

    def action_view_reports(self):
        self.ensure_one()
        self.check_access("read")
        action = self.env["ir.actions.actions"]._for_xml_id("l10n_gt_conciliacion_cuadratica.action_cq_reports")
        action["domain"] = [("id", "in", self.report_ids.ids)]
        if len(self.report_ids) == 1:
            action.update({"view_mode": "form", "views": [(False, "form")], "res_id": self.report_ids.id})
        return action

    def action_download(self):
        self.ensure_one()
        self.check_access("write")
        if not self.report_ids:
            raise UserError(_("Calcule los resultados antes de descargar."))
        self.report_ids.check_access("read")
        if len(self.report_ids) == 1:
            return self.report_ids.action_export()
        buffer = BytesIO()
        with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
            for report in self.report_ids:
                archive.writestr("%s_%s" % (report.id, report.file_name), report._xlsx_bytes())
        self.write({"file_data": base64.b64encode(buffer.getvalue()), "file_name": "conciliaciones_%s_%s.zip" % (self.year, self.month)})
        return {"type": "ir.actions.act_url", "url": "/web/content/cq.generate.wizard/%s/file_data/%s?download=true" % (self.id, self.file_name), "target": "self"}
