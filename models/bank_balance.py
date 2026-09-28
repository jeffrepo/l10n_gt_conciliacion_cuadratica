from odoo import api, fields, models, _
from odoo.exceptions import ValidationError

from ..core.calculation import MONTHS


class QuadraticBankBalance(models.Model):
    _name = "cq.bank.balance"
    _description = "Saldos mensuales según banco"
    _order = "year desc, month desc, account_id"
    _check_company_auto = True

    name = fields.Char("Referencia del estado de cuenta", required=True)
    company_id = fields.Many2one("res.company", string="Empresa", required=True, default=lambda self: self.env.company)
    account_id = fields.Many2one("account.account", string="Cuenta contable bancaria", required=True, check_company=True, ondelete="restrict")
    available_account_ids = fields.Many2many("account.account", compute="_compute_accounts")
    year = fields.Integer("Año", required=True, default=lambda self: fields.Date.context_today(self).year)
    month = fields.Selection([(str(i + 1), name) for i, name in enumerate(MONTHS)], string="Mes", required=True,
                             default=lambda self: str(fields.Date.context_today(self).month))
    currency_id = fields.Many2one("res.currency", string="Moneda", compute="_compute_currency", store=True)
    opening_balance = fields.Monetary("Saldo inicial según banco", required=True)
    closing_balance = fields.Monetary("Saldo final según banco", required=True)
    note = fields.Text("Observaciones")
    _period_unique = models.Constraint("UNIQUE(company_id, account_id, year, month)", "Ya existe una captura para esa empresa, cuenta, año y mes. Edite la captura existente.")

    def _journals(self):
        self.ensure_one()
        return self.env["account.journal"].with_context(active_test=False).search([
            ("company_id", "=", self.company_id.id), ("type", "=", "bank"), ("default_account_id", "=", self.account_id.id),
        ])

    @api.depends("company_id")
    def _compute_accounts(self):
        for balance in self:
            balance.available_account_ids = self.env["account.journal"].with_context(active_test=False).search([
                ("company_id", "=", balance.company_id.id), ("type", "=", "bank"),
            ]).default_account_id

    @api.depends("company_id", "account_id")
    def _compute_currency(self):
        for balance in self:
            journals = balance._journals() if balance.account_id else self.env["account.journal"]
            balance.currency_id = balance.account_id.currency_id or journals[:1].currency_id or balance.company_id.currency_id

    @api.onchange("company_id")
    def _onchange_company(self):
        self.account_id = False

    @api.constrains("company_id", "account_id", "year", "currency_id")
    def _check_configuration(self):
        for balance in self:
            if not 1900 <= balance.year <= 9998:
                raise ValidationError(_("El año debe estar entre 1900 y 9998."))
            journals = balance._journals()
            currencies = {journal.currency_id.id or balance.company_id.currency_id.id for journal in journals}
            if balance.company_id not in balance.account_id.company_ids or not journals:
                raise ValidationError(_("Seleccione una cuenta bancaria de la empresa indicada."))
            if currencies != {balance.currency_id.id}:
                raise ValidationError(_("La moneda de la cuenta debe coincidir con la de todos sus diarios bancarios."))
