from odoo import api, fields, models, _
from odoo.exceptions import ValidationError


class QuadraticConcept(models.Model):
    _name = "cq.concept"
    _description = "Concepto de conciliación cuadrática"
    _order = "direction, sequence, code"

    name = fields.Char("Concepto", required=True, translate=True)
    code = fields.Char("Código", required=True, index=True)
    report_code = fields.Char("Código en Excel", size=12, help="Código breve de presentación. El código interno identifica el concepto de forma estable.")
    sequence = fields.Integer("Orden", default=10)
    active = fields.Boolean(default=True)
    direction = fields.Selection([("in", "Ingreso"), ("out", "Egreso")], required=True)
    detail = fields.Selection(
        [("none", "Sin desglose"), ("partner", "Por contraparte"), ("bank", "Por cuenta bancaria")],
        default="none", required=True, string="Detalle",
    )
    _code_unique = models.Constraint("UNIQUE(code)", "El código del concepto debe ser único.")


class QuadraticRule(models.Model):
    _name = "cq.rule"
    _description = "Regla de clasificación cuadrática"
    _order = "sequence, id"
    _check_company_auto = True

    name = fields.Char("Nombre", required=True)
    active = fields.Boolean(default=True)
    sequence = fields.Integer("Prioridad", default=10, help="El número menor tiene prioridad. Empates entre conceptos quedan pendientes.")
    company_id = fields.Many2one("res.company", required=True, default=lambda self: self.env.company)
    journal_id = fields.Many2one("account.journal", string="Solo para el diario", check_company=True)
    concept_id = fields.Many2one("cq.concept", required=True, ondelete="restrict", string="Concepto")
    account_ids = fields.Many2many("account.account", string="Cuentas de contrapartida", check_company=True)
    partner_ids = fields.Many2many("res.partner", string="Contrapartes", check_company=True)
    country_scope = fields.Selection(
        [("any", "Cualquier país"), ("local", "País de la compañía"), ("foreign", "Otro país")],
        default="any", required=True, string="Ubicación de la contraparte",
    )
    method_code = fields.Char("Código de método de pago", help="Opcional; por ejemplo check_printing. Coincidencia exacta.")
    label_contains = fields.Char("Descripción contiene")
    fallback = fields.Boolean("Regla general", help="Permite una regla sin filtros. Utilice una prioridad mayor que las reglas específicas.")

    @api.constrains("account_ids", "partner_ids", "country_scope", "method_code", "label_contains", "fallback", "journal_id", "company_id")
    def _check_conditions(self):
        for rule in self:
            if not (rule.account_ids or rule.partner_ids or rule.country_scope != "any" or rule.method_code or rule.label_contains or rule.fallback):
                raise ValidationError(_("Defina una condición o marque explícitamente Regla general."))
            if rule.journal_id and (rule.journal_id.type != "bank" or rule.journal_id.company_id != rule.company_id):
                raise ValidationError(_("El diario debe ser bancario y pertenecer a la compañía de la regla."))
            if any(rule.company_id not in account.company_ids for account in rule.account_ids):
                raise ValidationError(_("Las cuentas de la regla deben estar disponibles en su compañía."))


class AccountJournal(models.Model):
    _inherit = "account.journal"

    cq_enabled = fields.Boolean("Incluir en conciliación cuadrática")
    cq_account_type = fields.Selection(
        [("monetary", "Monetaria"), ("savings", "Ahorro"), ("other", "Otra")],
        default="monetary", string="Tipo de cuenta bancaria",
    )
    cq_extra_outstanding_account_ids = fields.Many2many(
        "account.account", "cq_journal_outstanding_rel", "journal_id", "account_id",
        string="Cuentas pendientes adicionales", check_company=True,
        help="Las cuentas de cobros y pagos pendientes de los métodos de pago se incluyen automáticamente. Añada solo otras cuentas pendientes atribuibles a este banco.",
    )

    def _cq_outstanding_accounts(self):
        self.ensure_one()
        return (
            self.inbound_payment_method_line_ids.payment_account_id
            | self.outbound_payment_method_line_ids.payment_account_id
            | self.cq_extra_outstanding_account_ids
        ) - self.default_account_id - self.suspense_account_id

    @api.constrains("cq_enabled", "type", "cq_extra_outstanding_account_ids", "company_id", "default_account_id", "suspense_account_id")
    def _cq_check_configuration(self):
        for journal in self:
            if journal.cq_enabled and journal.type != "bank":
                raise ValidationError(_("Solo se pueden habilitar diarios bancarios."))
            for account in journal.cq_extra_outstanding_account_ids:
                if journal.company_id not in account.company_ids:
                    raise ValidationError(_("Las cuentas pendientes deben pertenecer a la compañía del diario."))
                if account in (journal.default_account_id | journal.suspense_account_id):
                    raise ValidationError(_("La cuenta bancaria y la cuenta transitoria no son cuentas pendientes adicionales."))
