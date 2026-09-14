from odoo import api, fields, models, _
from odoo.exceptions import ValidationError


class BankStatement(models.Model):
    _inherit = "account.bank.statement"

    cq_date_end = fields.Date(
        "Fecha de corte del extracto",
        help="Fin del período cubierto por el documento del banco. Puede ser posterior a la última transacción. Si se deja vacío se utiliza la fecha del extracto de Odoo.",
    )

    @api.constrains("cq_date_end", "line_ids", "date")
    def _cq_check_cutoff(self):
        for statement in self:
            dates = statement.line_ids.filtered(lambda line: line.state == "posted").mapped("date")
            if statement.cq_date_end and dates and statement.cq_date_end < max(dates):
                raise ValidationError(_("La fecha de corte no puede ser anterior a las transacciones del extracto."))


class BankStatementLine(models.Model):
    _inherit = "account.bank.statement.line"

    cq_concept_id = fields.Many2one("cq.concept", string="Concepto cuadrático", ondelete="restrict", copy=False)
    cq_allocation_ids = fields.One2many("cq.allocation", "statement_line_id", string="Distribución por conceptos", copy=False)
    cq_note = fields.Char("Observación de clasificación", copy=False)

    @api.constrains("cq_concept_id", "amount", "cq_allocation_ids")
    def _cq_check_classification(self):
        for line in self:
            direction = "in" if line.amount >= 0 else "out"
            if line.cq_concept_id and line.cq_concept_id.direction != direction:
                raise ValidationError(_("El concepto debe corresponder al signo del movimiento."))
            if line.cq_concept_id and line.cq_allocation_ids:
                raise ValidationError(_("Utilice un concepto único o una distribución, no ambos."))
            if line.currency_id.compare_amounts(sum(line.cq_allocation_ids.mapped("amount")), abs(line.amount)) > 0:
                raise ValidationError(_("La distribución no puede superar el importe bancario."))


class QuadraticAllocation(models.Model):
    _name = "cq.allocation"
    _description = "Distribución de movimiento bancario"
    _check_company_auto = True

    statement_line_id = fields.Many2one("account.bank.statement.line", required=True, ondelete="cascade", check_company=True)
    company_id = fields.Many2one(related="statement_line_id.company_id", store=True)
    currency_id = fields.Many2one(related="statement_line_id.currency_id")
    concept_id = fields.Many2one("cq.concept", required=True, ondelete="restrict", string="Concepto")
    amount = fields.Monetary("Importe", required=True)
    note = fields.Char("Descripción")

    @api.constrains("statement_line_id", "concept_id", "amount")
    def _check_amount(self):
        for allocation in self:
            if allocation.amount <= 0:
                raise ValidationError(_("El importe distribuido debe ser mayor que cero."))
            direction = "in" if allocation.statement_line_id.amount >= 0 else "out"
            if allocation.concept_id.direction != direction:
                raise ValidationError(_("El concepto no corresponde al signo del movimiento bancario."))
        self.statement_line_id._cq_check_classification()
