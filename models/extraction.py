from collections import defaultdict
import base64
from datetime import date
import hashlib

from odoo import Command, fields, models, _
from odoo.exceptions import AccessError, UserError

from ..core.calculation import build_snapshot, choose_rule, month_end
from ..core.export import safe_filename
from .report import INTERNAL


class QuadraticExtraction(models.Model):
    _inherit = "cq.report"

    def _generate(self, company, account, year, month, template=None):
        """Private entry point: every search honors the requesting user's ACLs."""
        if company.id not in self.env.companies.ids:
            raise AccessError(_("Seleccione una compañía permitida."))
        company.ensure_one()
        account.ensure_one()
        # A report has one company, independently of the companies checked in
        # the top bar. Re-browse inputs so unrelated prefetch records cannot
        # pull contacts from other companies into metadata computations.
        model = self.with_context(allowed_company_ids=company.ids).with_company(company)
        company = model.env["res.company"].browse(company.id)
        account = model.env["account.account"].browse(account.id)
        account.check_access("read")
        if company not in account.company_ids:
            raise UserError(_("La cuenta contable debe estar disponible en la compañía indicada."))
        journals = model.env["account.journal"].with_context(active_test=False).search([
            ("company_id", "=", company.id), ("type", "=", "bank"),
            ("default_account_id", "=", account.id),
        ], order="id")
        if not journals.filtered("cq_enabled"):
            raise UserError(_("Habilite la conciliación cuadrática en al menos un diario bancario de esta cuenta."))
        currencies = {bank.currency_id.id or company.currency_id.id for bank in journals}
        if len(currencies) != 1:
            raise UserError(_("Los diarios de la cuenta %s tienen monedas diferentes. Revise su configuración antes de consolidarlos.", account.display_name))
        currency = journals[:1].currency_id or company.currency_id
        if account.currency_id and account.currency_id != currency:
            raise UserError(_("La moneda de la cuenta contable no coincide con la de sus diarios bancarios."))
        generated_at = fields.Datetime.now()
        template_model = model.env["cq.xlsx.template"]
        template = template_model.browse(template.id) if template else template_model.search([
            ("company_id", "=", company.id), ("account_id", "=", account.id),
        ], limit=1)
        if template:
            template.check_access("read")
            if template.company_id != company or template.account_id != account:
                raise UserError(_("La plantilla debe pertenecer a la empresa y a la cuenta contable del reporte."))
        template_data = template.file_data if template else False
        data = model._extract(company, account, journals, currency, year, month)
        data["metadata"].update({"generated_at": str(generated_at), "generated_by": self.env.user.display_name})
        if template:
            data["metadata"].update({"template_name": template.name, "template_sha256": hashlib.sha256(base64.b64decode(template_data)).hexdigest()})
        payload = build_snapshot(data)
        month_values = []
        month_model = self.env["cq.report.month"]
        for result in payload["months"]:
            vals = {key: value for key, value in result.items() if key in month_model._fields and key not in ("control", "concept_totals")}
            vals.update({"bank_available": result["bank_end"] is not None, "control_available": bool(result["control"])})
            month_values.append(Command.create(vals))
        issues = [Command.create({
            "code": issue["code"], "month": issue.get("month", 0),
            "message": issue["message"], "statement_line_id": issue.get("source_id"),
        }) for issue in payload["issues"]]
        filename = safe_filename("%s_%s_%s_%s_%04d_%02d" % (company.name, account.code, account.name, currency.name, year, month)) + ".xlsx"
        report = model.with_context(cq_internal=INTERNAL).create({
            "name": "%s / %04d-%02d / %s" % (account.display_name, year, month, generated_at),
            "company_id": company.id, "account_id": account.id,
            "journal_ids": [Command.set(journals.ids)],
            "journal_id": data["metadata"]["control_journal_id"],
            "currency_id": currency.id,
            "year": year, "month": month, "generated_at": generated_at,
            "generated_by": self.env.uid, "payload": payload,
            "month_ids": month_values, "issue_ids": issues,
            "issue_count": len(payload["issues"]),
            "unclassified_count": len({issue["source_id"] for issue in payload["issues"] if issue["code"] == "classification"}),
            "difference": payload["months"][-1]["difference"], "file_name": filename,
            "template_id": template.id, "template_name": template.file_name if template else False,
            "template_data": template_data,
        })
        return report.with_env(self.env)

    def _extract(self, company, account, journals, currency, year, month):
        start, cutoff = date(year, 1, 1), month_end(year, month)
        issues = []
        partner_cache, bank_cache, restricted_sources = {}, {}, set()

        def contact_info(partner, document, source_id=False):
            info = self._partner_info(partner, partner_cache)
            if info["restricted"] and ("partner", partner.id) not in restricted_sources:
                restricted_sources.add(("partner", partner.id))
                issues.append({
                    "code": "partner_access", "source_id": source_id,
                    "message": _("El documento %s tiene un contacto restringido en la empresa del reporte. Se conserva el importe y se omiten los datos del contacto. Revise la compañía y los permisos de la contraparte en el registro original.", document),
                })
            return info

        def bank_info(bank, document, source_id=False):
            info = self._counterparty_bank_info(bank, partner_cache, bank_cache)
            if info["restricted"] and ("bank", bank.id) not in restricted_sources:
                restricted_sources.add(("bank", bank.id))
                issues.append({
                    "code": "partner_bank_access", "source_id": source_id,
                    "message": _("El documento %s tiene datos bancarios de contraparte restringidos en la empresa del reporte. Se conserva el importe y se omiten esos datos. Revise el registro original.", document),
                })
            return info

        bank_lines = self.env["account.bank.statement.line"].search([
            ("company_id", "=", company.id), ("journal_id", "in", journals.ids),
            ("state", "=", "posted"), ("date", "<=", cutoff),
        ], order="internal_index, id")
        all_banks = self.env["account.journal"].with_context(active_test=False).search([("company_id", "=", company.id), ("type", "=", "bank")])
        owners = defaultdict(set)
        for bank in all_banks:
            for ledger_account in bank.default_account_id | bank.suspense_account_id | bank._cq_outstanding_accounts():
                owners[ledger_account.id].add(bank.default_account_id.id)
        payment_methods = journals.inbound_payment_method_line_ids | journals.outbound_payment_method_line_ids
        if payment_methods.filtered(lambda method: method.payment_account_id == account):
            issues.append({"code": "direct_bank", "message": "Hay métodos que contabilizan directamente en la cuenta bancaria. Esta versión requiere cuentas separadas de cobros/pagos pendientes para identificar partidas en tránsito."})
        bank_accounts = journals.bank_account_id
        if not bank_accounts:
            issues.append({"code": "bank_account", "message": "Los diarios de esta cuenta contable no tienen una cuenta bancaria vinculada."})
        elif len(bank_accounts) > 1:
            issues.append({"code": "bank_identity", "message": "La cuenta contable está vinculada a varias cuentas bancarias físicas. Revise esta configuración antes de conservar el cierre."})
        identified_account = bank_accounts if len(bank_accounts) == 1 else self.env["res.partner.bank"]
        if identified_account and not identified_account.bank_id:
            issues.append({"code": "bank_name", "message": "La cuenta bancaria vinculada no tiene un banco identificado. Configure el banco en esa cuenta; el nombre de la cuenta contable no identifica la entidad bancaria."})
        outstanding = self.env["account.account"]
        for bank in journals:
            outstanding |= bank._cq_outstanding_accounts()
            if not bank.suspense_account_id:
                issues.append({"code": "suspense_configuration", "message": "El diario %s no tiene cuenta transitoria bancaria." % bank.display_name})
        suspense_accounts = journals.suspense_account_id - account
        if outstanding & suspense_accounts or account in journals.suspense_account_id:
            issues.append({"code": "account_roles", "message": "Una cuenta se usa con funciones incompatibles de banco, pendientes o transitoria entre los diarios incluidos. Revise su configuración."})
        outstanding -= suspense_accounts | account
        for pending_account in outstanding:
            if not pending_account.reconcile:
                issues.append({"code": "account_configuration", "message": "La cuenta pendiente %s debe permitir conciliación." % pending_account.display_name})
        ledger_accounts = account | suspense_accounts | outstanding
        ledger_lines = self.env["account.move.line"].search([
            ("company_id", "=", company.id), ("account_id", "in", ledger_accounts.ids),
            ("parent_state", "=", "posted"), ("date", "<=", cutoff),
        ], order="date, id")
        ledger = []
        owner_cache = {}
        for line in ledger_lines:
            # The bank ledger itself belongs to the selected account, including
            # opening/manual entries from general journals. Attribute only the
            # shared clearing accounts through their source bank account.
            if line.account_id != account:
                owner = self._line_owner(line, owners[line.account_id.id], owner_cache)
                if owner is None:
                    issues.append({"code": "ownership", "message": "No se puede atribuir el apunte %s (%s) a una única cuenta contable bancaria." % (line.move_id.name, line.account_id.display_name)})
                    continue
                if owner != account.id:
                    continue
            is_company_currency = currency == company.currency_id
            if not is_company_currency and line.currency_id != currency:
                # Never invent a historical bank-currency amount using today's
                # exchange rate, including when a shared account is in GTQ.
                if not company.currency_id.is_zero(line.balance):
                    issues.append({"code": "currency", "message": "El apunte %s no conserva el importe en %s. Revise la moneda de sus cuentas pendientes." % (line.move_id.name, currency.name)})
                continue
            role = "bank" if line.account_id == account else ("outstanding" if line.account_id in outstanding else "suspense")
            matches = []
            for partial in line.matched_debit_ids | line.matched_credit_ids:
                debit = partial.debit_move_id == line
                other = partial.credit_move_id if debit else partial.debit_move_id
                if other.move_id.state != "posted":
                    continue
                part_amount = partial.amount if is_company_currency else (partial.debit_amount_currency if debit else partial.credit_amount_currency)
                matches.append({"date": str(partial.max_date), "delta": -part_amount if debit else part_amount,
                                "company_delta": -partial.amount if debit else partial.amount})
            ledger.append({
                "source_id": line.id, "move_id": line.move_id.id, "date": str(line.date),
                "document": line.move_id.name, "description": line.name or "",
                "partner": contact_info(line.partner_id, line.move_id.name, line.statement_line_id.id)["name"],
                "account": line.account_id.display_name,
                "journal": line.journal_id.display_name,
                "method_code": line.payment_id.payment_method_code or "",
                "method": line.payment_id.payment_method_line_id.name or "",
                "amount": line.balance if is_company_currency else line.amount_currency,
                "company_amount": line.balance, "role": role, "matches": matches,
            })

        # Use an independently stored statement opening as the anchor. The
        # bank roll-forward never resets to a later ending balance to hide gaps.
        control_journal = journals.filtered("cq_statement_source")
        source_journals = bank_lines.journal_id
        if not control_journal and len(source_journals) == 1:
            control_journal = source_journals
        if len(control_journal) > 1 or (not control_journal and len(source_journals) > 1):
            issues.append({"code": "statement_source", "message": "Hay movimientos bancarios en varios diarios. Marque 'Usar extractos como control de la cuenta' en el diario que contiene el estado de cuenta completo. No se suman saldos de extractos de distintos diarios."})
            control_journal = self.env["account.journal"]
        statements = bank_lines.filtered(lambda line: line.journal_id == control_journal).statement_id.sorted("first_line_index")
        opening = None
        if statements:
            before = statements.filtered(lambda stmt: stmt.first_line_index[:8] < start.strftime("%Y%m%d"))
            anchor = before[-1:] or statements[:1]
            opening = anchor.balance_start
            for line in bank_lines:
                if line.internal_index >= anchor.first_line_index and line.date < start:
                    opening += line.amount
                elif line.internal_index < anchor.first_line_index and line.date >= start:
                    opening -= line.amount
        controls = {}
        for number in range(1, month + 1):
            ending = month_end(year, number)
            candidates = statements.filtered(lambda stmt: (stmt.cq_date_end or stmt.date) == ending)
            if candidates:
                control = candidates[-1:]
                controls[str(number)] = {
                    "source_id": control.id, "name": control.name or "", "date": str(ending),
                    "balance": control.balance_end_real, "complete": control.is_complete,
                }
        for statement in statements:
            if statement.date and statement.date >= start and not statement.is_complete:
                issues.append({"code": "statement_incomplete", "message": "El extracto %s contiene una diferencia entre sus líneas y su saldo final." % statement.display_name})

        concepts = self.env["cq.concept"].with_context(active_test=False).search([], order="sequence, code")
        rules = self.env["cq.rule"].search([
            ("company_id", "=", company.id), "|", ("journal_id", "=", False), ("journal_id", "in", journals.ids),
        ])
        rule_data = [{
            "id": rule.id, "sequence": rule.sequence, "concept_code": rule.concept_id.code, "journal_id": rule.journal_id.id,
            "direction": rule.concept_id.direction, "partner_ids": rule.partner_ids.ids,
            "account_ids": rule.account_ids.ids, "country_scope": rule.country_scope,
            "method_code": rule.method_code or "", "label_contains": rule.label_contains or "",
        } for rule in rules if rule.concept_id.active]
        movements = []
        for line in bank_lines.filtered(lambda item: item.date >= start):
            linked_moves = self._linked_moves(line, cutoff)
            payments = linked_moves.origin_payment_id
            partner = line.partner_id or (payments.partner_id if len(payments.partner_id) == 1 else self.env["res.partner"])
            partner_data = contact_info(partner, line.move_id.name, line.id)
            partner_country_id = partner_data["country_id"]
            country_scope = "unknown"
            if partner_country_id and company.country_id:
                country_scope = "local" if partner_country_id == company.country_id.id else "foreign"
            counterpart_accounts = (line.move_id | linked_moves).line_ids.account_id - ledger_accounts
            description = line.payment_ref or line.move_id.ref or ""
            allocations = []
            if line.cq_allocation_ids:
                allocations = [{"code": item.concept_id.code, "amount": item.amount, "origin": "manual", "note": item.note or ""} for item in line.cq_allocation_ids]
            elif line.cq_concept_id and not currency.is_zero(line.amount):
                allocations = [{"code": line.cq_concept_id.code, "amount": abs(line.amount), "origin": "manual"}]
            else:
                rule, ambiguous = choose_rule(
                    [rule for rule in rule_data if not rule["journal_id"] or rule["journal_id"] == line.journal_id.id],
                    "in" if line.amount >= 0 else "out", partner_data["partner_id"],
                    country_scope, counterpart_accounts.ids, payments.mapped("payment_method_code"), description,
                )
                if ambiguous:
                    issues.append({"code": "rule_ambiguity", "source_id": line.id,
                                   "message": "Movimiento %s: dos reglas con la misma prioridad proponen conceptos distintos." % line.move_id.name})
                if rule and not currency.is_zero(line.amount):
                    allocations = [{"code": rule["concept_code"], "amount": abs(line.amount), "origin": "rule", "rule_id": rule["id"]}]
                elif not ambiguous and line.amount > 0:
                    code = self._customer_receipt_code(line.move_id, account, outstanding, company, partner_data, cutoff)
                    if code in {concept.code for concept in concepts if concept.active}:
                        allocations = [{"code": code, "amount": line.amount, "origin": "reconciliation"}]
            partner_banks = []
            if partner_data["commercial_id"]:
                for bank in payments.partner_bank_id:
                    details = bank_info(bank, line.move_id.name, line.id)
                    if details["commercial_id"] == partner_data["commercial_id"]:
                        partner_banks.append(details)
            identified_bank = partner_banks[0] if len(partner_banks) == 1 else {}
            company_amount = sum(line.move_id.line_ids.filtered(lambda aml: aml.account_id == account).mapped("balance"))
            movements.append({
                "source_id": line.id, "move_id": line.move_id.id, "date": str(line.date),
                "accounting_date": str(line.move_id.date),
                "document": line.move_id.name or "", "reference": line.move_id.ref or "",
                "journal": line.journal_id.display_name,
                "partner": partner_data["name"] or line.partner_name or "",
                "description": description, "amount": line.amount, "company_amount": company_amount,
                "accounting_dates": ", ".join(sorted({str(move.date) for move in linked_moves})),
                "linked_documents": ", ".join(sorted(set(linked_moves.mapped("name")))),
                "method": ", ".join(sorted(set(payments.payment_method_line_id.mapped("name")))) or line.transaction_type or "",
                "counterparty_account": identified_bank.get("number") or line.account_number or "",
                "counterparty_bank": identified_bank.get("name", ""),
                "note": line.cq_note or "", "allocations": allocations,
            })
        accounting_movements = []
        ledger_by_id = {row["source_id"]: row for row in ledger if row["role"] == "bank"}
        movements_by_move = {row["move_id"]: row for row in movements}
        active_codes = {concept.code for concept in concepts if concept.active}
        for line in ledger_lines.filtered(lambda aml: aml.account_id == account and aml.date >= start):
            if line.id not in ledger_by_id:
                continue  # Currency conversion was not backed by this entry.
            row = dict(ledger_by_id[line.id])
            payment = line.move_id.origin_payment_id
            counterpart = line.move_id.line_ids.filtered(lambda aml: aml.account_id != account)
            partners = counterpart.partner_id
            partner = line.partner_id or payment.partner_id or (partners if len(partners) == 1 else self.env["res.partner"])
            details = contact_info(partner, line.move_id.name)
            scope = "unknown"
            if details["country_id"] and company.country_id:
                scope = "local" if details["country_id"] == company.country_id.id else "foreign"
            bank_movement = movements_by_move.get(line.move_id.id)
            code, origin = "", "pending"
            if bank_movement:
                # Reuse an already classified transaction only if it represents
                # this exact ledger amount and has a single full concept.
                allocations = bank_movement["allocations"]
                codes = {item["code"] for item in allocations}
                if len(codes) == 1 and currency.compare_amounts(sum(item["amount"] for item in allocations), abs(row["amount"])) == 0:
                    code, origin = allocations[0]["code"], allocations[0]["origin"]
            else:
                rule, ambiguous = choose_rule(
                    [rule for rule in rule_data if not rule["journal_id"] or rule["journal_id"] == line.journal_id.id],
                    "in" if row["amount"] >= 0 else "out", details["partner_id"], scope,
                    counterpart.account_id.ids, payment.mapped("payment_method_code"), line.name or line.move_id.ref or "",
                )
                if rule:
                    code, origin = rule["concept_code"], "rule"
                elif not ambiguous and row["amount"] > 0:
                    automatic = self._customer_receipt_code(line.move_id, account, outstanding, company, details, cutoff)
                    if automatic in active_codes:
                        code, origin = automatic, "reconciliation"
            row.update({"partner": details["name"], "reference": line.move_id.ref or "",
                        "concept_code": code, "classification_origin": origin,
                        "statement_line_id": line.statement_line_id.id,
                        "counterpart_accounts": ", ".join(counterpart.account_id.mapped("display_name"))})
            accounting_movements.append(row)
        if not movements and accounting_movements:
            issues.append({"code": "bank_transactions_missing", "message": _(
                "No se encontraron transacciones bancarias publicadas entre %s y %s en los diarios incluidos. Sí hay %s apuntes publicados en la cuenta contable: consulte Mayor bancario Odoo y Resumen contable Odoo. Pagos, cobros y asientos no son extractos; sus saldos no sustituyen el saldo según banco.", start, cutoff, len(accounting_movements))})
        account_types = set(journals.mapped("cq_account_type"))
        account_type = dict(journals._fields["cq_account_type"].selection).get(next(iter(account_types)), "") if len(account_types) == 1 else ""
        return {
            "year": year, "month": month, "rounding": str(currency.rounding),
            "company_rounding": str(company.currency_id.rounding),
            "metadata": {
                "company": company.name, "vat": company.vat or "", "country": company.country_id.name or "",
                "bank": identified_account.bank_id.name or "",
                "account_number": identified_account.acc_number or "", "account_type": account_type,
                "journal": ", ".join(journals.mapped("display_name")), "ledger_account": account.display_name,
                "ledger_account_name": account.name, "ledger_account_code": account.code,
                "account_id": account.id, "journal_ids": journals.ids,
                "control_journal_id": control_journal.id,
                "control_journal": control_journal.display_name or "",
                "currency": currency.name, "company_currency": company.currency_id.name,
                "base_url": self.env["ir.config_parameter"].sudo().get_param("web.base.url", ""),
            },
            "bank_opening": opening, "controls": controls, "ledger": ledger,
            "concepts": [{"code": item.code, "report_code": item.report_code or "", "name": item.name, "direction": item.direction,
                          "sequence": item.sequence, "detail": item.detail} for item in concepts],
            "movements": movements, "accounting_movements": accounting_movements, "issues": issues,
        }

    def _customer_receipt_code(self, move, bank_account, outstanding, company, partner_data, cutoff):
        """Only pure customer receipts; explicit rules always take precedence.

        A clearing line must be fully matched at the cutoff to posted customer
        receipts. Mixed fees, refunds, partial matches, unknown countries and
        restricted/different counterparties stay for explicit classification.
        """
        if not partner_data["commercial_id"] or not partner_data["country_id"] or not company.country_id:
            return None
        cache = {}

        def receivable_lines(lines):
            return bool(lines) and all(
                aml.account_id.account_type == "asset_receivable" and aml.balance < 0
                and self._partner_info(aml.partner_id, cache)["commercial_id"] == partner_data["commercial_id"]
                for aml in lines
            )

        counterpart = move.line_ids.filtered(lambda aml: aml.account_id != bank_account and not company.currency_id.is_zero(aml.balance))
        if not counterpart:
            return None
        for aml in counterpart:
            if receivable_lines(aml):
                continue
            if aml.account_id not in outstanding or aml.balance >= 0:
                return None
            partials = aml.matched_debit_ids.filtered(lambda partial: partial.max_date <= cutoff and partial.debit_move_id.move_id.state == "posted")
            if company.currency_id.compare_amounts(sum(partials.mapped("amount")), -aml.balance):
                return None
            for receipt in partials.debit_move_id.move_id:
                payment = receipt.origin_payment_id
                if not payment or payment.payment_type != "inbound" or payment.partner_type != "customer":
                    return None
                other = receipt.line_ids.filtered(lambda item: item.account_id != aml.account_id and not company.currency_id.is_zero(item.balance))
                if not receivable_lines(other):
                    return None
        return "IN_CUSTOMERS_LOCAL" if partner_data["country_id"] == company.country_id.id else "IN_CUSTOMERS_FOREIGN"

    def _partner_info(self, partner, cache):
        """Read optional contact data with the report user's real permissions.

        IDs on accessible entries do not grant access to their related records.
        Guard even cached fields, and isolate prefetch before reading a contact
        or its commercial entity. Never widen companies or use sudo here.
        """
        empty = {"partner_id": False, "name": "", "commercial_id": False,
                 "country_id": False, "restricted": False}
        if not partner:
            return empty
        if partner.id not in cache:
            partner = self.env["res.partner"].browse(partner.id)
            try:
                partner.check_access("read")
                commercial = self.env["res.partner"].browse(partner.commercial_partner_id.id)
                if commercial:
                    commercial.check_access("read")
                cache[partner.id] = {
                    "partner_id": partner.id, "name": partner.display_name or "",
                    "commercial_id": commercial.id, "country_id": commercial.country_id.id,
                    "restricted": False,
                }
            except AccessError:
                cache[partner.id] = {**empty, "name": _("Contacto restringido"), "restricted": True}
        return cache[partner.id]

    def _counterparty_bank_info(self, bank, partner_cache, cache):
        if bank.id not in cache:
            bank = self.env["res.partner.bank"].browse(bank.id)
            restricted = {"commercial_id": False, "number": "", "name": "", "restricted": True}
            try:
                bank.check_access("read")
                partner = self._partner_info(bank.partner_id, partner_cache)
                if partner["restricted"]:
                    cache[bank.id] = restricted
                else:
                    cache[bank.id] = {"commercial_id": partner["commercial_id"],
                                      "number": bank.acc_number or "", "name": bank.bank_id.name or "",
                                      "restricted": False}
            except AccessError:
                cache[bank.id] = restricted
        return cache[bank.id]

    def _line_owner(self, line, account_owners, cache):
        """Return the bank ledger account owning a shared clearing entry.

        Unlinked exchange/manual entries may inherit an unambiguous owner
        from their reconciliation component. Never allocate by partner alone.
        """
        if line.id in cache:
            return cache[line.id]

        def direct_owner(item):
            if item.statement_line_id:
                return item.statement_line_id.journal_id.default_account_id.id
            if item.payment_id:
                return item.payment_id.journal_id.default_account_id.id
            if item.journal_id.type == "bank":
                return item.journal_id.default_account_id.id
            return None

        owner = direct_owner(line)
        if owner:
            cache[line.id] = owner
            return owner
        if len(account_owners) == 1:
            return next(iter(account_owners))
        visited, found = set(), set()
        queue = line
        while queue:
            item, queue = queue[:1], queue[1:]
            if item.id in visited:
                continue
            visited.add(item.id)
            owner = direct_owner(item)
            if owner:
                found.add(owner)
                continue
            links = item.matched_debit_ids | item.matched_credit_ids
            queue |= (links.debit_move_id | links.credit_move_id).filtered(lambda other: other.id not in visited)
        owner = next(iter(found)) if len(found) == 1 else None
        cache[line.id] = owner
        return owner

    def _linked_moves(self, statement_line, cutoff):
        initial = statement_line.move_id.line_ids
        partials = (initial.matched_debit_ids | initial.matched_credit_ids).filtered(lambda partial: partial.max_date <= cutoff)
        moves = (partials.debit_move_id | partials.credit_move_id).move_id - statement_line.move_id
        payment_lines = moves.filtered("origin_payment_id").line_ids
        invoice_partials = (payment_lines.matched_debit_ids | payment_lines.matched_credit_ids).filtered(lambda partial: partial.max_date <= cutoff)
        invoices = (invoice_partials.debit_move_id | invoice_partials.credit_move_id).move_id.filtered(lambda move: move.is_invoice(include_receipts=True))
        return (moves | invoices).filtered(lambda move: move.state == "posted" and move.date <= cutoff)
