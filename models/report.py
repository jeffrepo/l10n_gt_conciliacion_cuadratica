import base64

from odoo import api, fields, models, _
from odoo.exceptions import AccessError, UserError

from ..core.export import export_xlsx


# An RPC context cannot recreate this object. Snapshot values can only be
# created by the server extraction path, never supplied as editable balances.
INTERNAL = object()


class SnapshotMixin(models.AbstractModel):
    _name = "cq.snapshot.mixin"
    _description = "Protección de resultados históricos"

    @api.model_create_multi
    def create(self, vals_list):
        if self.env.context.get("cq_internal") is not INTERNAL:
            raise AccessError(_("Genere el resultado desde el asistente de conciliación cuadrática."))
        return super().create(vals_list)

    def write(self, vals):
        if self.env.context.get("cq_internal") is not INTERNAL:
            raise AccessError(_("El resultado histórico es inmutable. Genere una nueva versión para actualizarlo."))
        return super().write(vals)

    def unlink(self):
        raise AccessError(_("Los resultados históricos no se eliminan; genere una nueva versión."))


class QuadraticReport(models.Model):
    _name = "cq.report"
    _inherit = "cq.snapshot.mixin"
    _description = "Conciliación cuadrática por cuenta"
    _order = "create_date desc, id desc"
    _check_company_auto = True

    name = fields.Char("Referencia", required=True)
    company_id = fields.Many2one("res.company", required=True, index=True)
    # Optional for snapshots created before grouping by ledger account. Their
    # original payload, journal and exported bytes must remain untouched.
    account_id = fields.Many2one("account.account", string="Cuenta contable bancaria", check_company=True, index=True)
    journal_ids = fields.Many2many("account.journal", string="Diarios incluidos", check_company=True)
    journal_id = fields.Many2one("account.journal", string="Diario de control / anterior", check_company=True, index=True)
    currency_id = fields.Many2one("res.currency", string="Moneda", required=True)
    year = fields.Integer("Año", required=True)
    month = fields.Integer("Mes de corte", required=True)
    state = fields.Selection([("draft", "Borrador"), ("confirmed", "Cierre conservado")], default="draft", required=True)
    generated_at = fields.Datetime("Generado el", required=True)
    generated_by = fields.Many2one("res.users", required=True, string="Generado por")
    confirmed_at = fields.Datetime("Cierre conservado el")
    confirmed_by = fields.Many2one("res.users", string="Cierre conservado por")
    payload = fields.Json("Resultado", required=True)
    issue_count = fields.Integer("Pendientes de revisión")
    unclassified_count = fields.Integer("Movimientos sin clasificar")
    difference = fields.Monetary("Diferencia al corte")
    month_ids = fields.One2many("cq.report.month", "report_id", string="Resumen mensual")
    issue_ids = fields.One2many("cq.report.issue", "report_id", string="Revisión")
    file_data = fields.Binary("Archivo XLSX", attachment=True)
    file_name = fields.Char("Nombre de archivo")

    def action_confirm(self):
        self.check_access("write")
        for report in self:
            if report.issue_count:
                raise UserError(_("Resuelva los pendientes en Odoo y genere una nueva versión antes de conservar el cierre."))
            if report.state == "confirmed":
                continue
            report.with_context(cq_internal=INTERNAL).write({
                "state": "confirmed", "confirmed_at": fields.Datetime.now(),
                "confirmed_by": self.env.uid, "file_data": False,
            })
        return True

    def _xlsx_bytes(self):
        self.ensure_one()
        self.check_access("read")
        if self.file_data:
            return base64.b64decode(self.file_data)
        return export_xlsx(self.payload, confirmed=self.state == "confirmed")

    def action_export(self):
        self.ensure_one()
        self.check_access("write")
        self.with_context(cq_internal=INTERNAL).write({"file_data": base64.b64encode(self._xlsx_bytes())})
        return {"type": "ir.actions.act_url", "url": "/web/content/cq.report/%s/file_data/%s?download=true" % (self.id, self.file_name), "target": "self"}

    def action_new_version(self):
        self.ensure_one()
        self.check_access("read")
        return {
            "type": "ir.actions.act_window", "res_model": "cq.generate.wizard",
            "view_mode": "form", "target": "new", "name": _("Nueva versión"),
            "context": {"default_company_id": self.company_id.id, "default_year": self.year,
                        "default_month": str(self.month),
                        "default_account_ids": [(6, 0, (self.account_id or self.journal_id.default_account_id).ids)]},
        }

    def action_movements(self):
        self.ensure_one()
        self.check_access("read")
        action = self.env["ir.actions.actions"]._for_xml_id("l10n_gt_conciliacion_cuadratica.action_cq_movements")
        action["domain"] = [("id", "in", [row["source_id"] for row in self.payload["movements"]])]
        return action


class QuadraticMonth(models.Model):
    _name = "cq.report.month"
    _inherit = "cq.snapshot.mixin"
    _description = "Resumen mensual cuadrático"
    _order = "number"

    report_id = fields.Many2one("cq.report", required=True, ondelete="cascade", index=True)
    company_id = fields.Many2one(related="report_id.company_id", store=True)
    currency_id = fields.Many2one(related="report_id.currency_id")
    number = fields.Integer("Número de mes")
    name = fields.Char("Mes")
    cutoff = fields.Date("Fecha de corte")
    bank_available = fields.Boolean("Saldo inicial respaldado")
    control_available = fields.Boolean("Extracto de cierre disponible")
    bank_opening = fields.Monetary("Saldo inicial banco")
    income = fields.Monetary("Ingresos")
    expense = fields.Monetary("Egresos")
    bank_end = fields.Monetary("Saldo calculado banco")
    statement_end = fields.Monetary("Saldo extracto")
    bank_difference = fields.Monetary("Diferencia extracto")
    deposits = fields.Monetary("Depósitos en tránsito")
    checks = fields.Monetary("Cheques en circulación")
    payments = fields.Monetary("Otros pagos pendientes")
    bank_adjusted = fields.Monetary("Banco ajustado")
    ledger_bank = fields.Monetary("Mayor de banco")
    ledger_outstanding = fields.Monetary("Mayor de pendientes")
    ledger_suspense = fields.Monetary("Mayor de transitoria")
    book_balance = fields.Monetary("Saldo según libros")
    book_adjustment = fields.Monetary("Ajustes identificados de libros")
    book_adjusted = fields.Monetary("Libros ajustados")
    difference = fields.Monetary("Diferencia conciliación")


class QuadraticIssue(models.Model):
    _name = "cq.report.issue"
    _inherit = "cq.snapshot.mixin"
    _description = "Pendiente de revisión cuadrática"

    report_id = fields.Many2one("cq.report", required=True, ondelete="cascade", index=True)
    company_id = fields.Many2one(related="report_id.company_id", store=True)
    code = fields.Char("Tipo")
    month = fields.Integer("Mes")
    message = fields.Text("Detalle")
    statement_line_id = fields.Many2one("account.bank.statement.line", string="Movimiento", ondelete="set null")
