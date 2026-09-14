from collections import defaultdict
from datetime import date

from odoo import Command, fields, models, _
from odoo.exceptions import AccessError, UserError

from ..core.calculation import build_snapshot, choose_rule, month_end
from ..core.export import safe_filename
from .report import INTERNAL


class QuadraticExtraction(models.Model):
    _inherit = "cq.report"

    def _generate(self, company, account, year, month):
        """Private entry point: every search honors the requesting user's ACLs."""
        if company.id not in self.env.companies.ids:
            raise AccessError(_("Seleccione una compañía permitida."))
        account.ensure_one()
        account.check_access("read")
        if company not in account.company_ids:
            raise UserError(_("La cuenta contable debe estar disponible en la compañía indicada."))
        model = self.with_company(company)
        account = account.with_company(company)
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
        data = model._extract(company, account, journals, currency, year, month)
        data["metadata"].update({"generated_at": str(generated_at), "generated_by": self.env.user.display_name})
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
        })
        return report.with_env(self.env)

    def _extract(self, company, account, journals, currency, year, month):
        start, cutoff = date(year, 1, 1), month_end(year, month)
        issues = []
        bank_lines = self.env["account.bank.statement.line"].search([
            ("company_id", "=", company.id), ("journal_id", "in", journals.ids),
            ("state", "=", "posted"), ("date", "<=", cutoff),
        ], order="date, id")
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
                "partner": line.partner_id.display_name or "", "account": line.account_id.display_name,
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
            anchor_line = anchor.line_ids.sorted(lambda line: (line.date, line.id))[:1]
            anchor_key = (anchor_line.date, anchor_line.id)
            for line in bank_lines:
                if (line.date, line.id) >= anchor_key and line.date < start:
                    opening += line.amount
                elif (line.date, line.id) < anchor_key and line.date >= start:
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
            partner_country = partner.commercial_partner_id.country_id
            country_scope = "unknown"
            if partner_country and company.country_id:
                country_scope = "local" if partner_country == company.country_id else "foreign"
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
                    "in" if line.amount >= 0 else "out", partner.id,
                    country_scope, counterpart_accounts.ids, payments.mapped("payment_method_code"), description,
                )
                if ambiguous:
                    issues.append({"code": "rule_ambiguity", "source_id": line.id,
                                   "message": "Movimiento %s: dos reglas con la misma prioridad proponen conceptos distintos." % line.move_id.name})
                if rule and not currency.is_zero(line.amount):
                    allocations = [{"code": rule["concept_code"], "amount": abs(line.amount), "origin": "rule", "rule_id": rule["id"]}]
            partner_banks = payments.partner_bank_id.filtered(lambda bank: bank.partner_id.commercial_partner_id == partner.commercial_partner_id) if partner else self.env["res.partner.bank"]
            identified_bank = partner_banks if len(partner_banks) == 1 else self.env["res.partner.bank"]
            company_amount = sum(line.move_id.line_ids.filtered(lambda aml: aml.account_id == account).mapped("balance"))
            movements.append({
                "source_id": line.id, "move_id": line.move_id.id, "date": str(line.date),
                "document": line.move_id.name or "", "reference": line.move_id.ref or "",
                "journal": line.journal_id.display_name,
                "partner": partner.display_name or line.partner_name or "",
                "description": description, "amount": line.amount, "company_amount": company_amount,
                "accounting_dates": ", ".join(sorted({str(move.date) for move in linked_moves})),
                "linked_documents": ", ".join(sorted(set(linked_moves.mapped("name")))),
                "method": ", ".join(sorted(set(payments.payment_method_line_id.mapped("name")))) or line.transaction_type or "",
                "counterparty_account": identified_bank.acc_number or line.account_number or "",
                "counterparty_bank": identified_bank.bank_id.name or "",
                "note": line.cq_note or "", "allocations": allocations,
            })
        account_types = set(journals.mapped("cq_account_type"))
        account_type = dict(journals._fields["cq_account_type"].selection).get(next(iter(account_types)), "") if len(account_types) == 1 else ""
        return {
            "year": year, "month": month, "rounding": str(currency.rounding),
            "company_rounding": str(company.currency_id.rounding),
            "metadata": {
                "company": company.name, "vat": company.vat or "", "country": company.country_id.name or "",
                "bank": identified_account.bank_id.name or account.name,
                "account_number": identified_account.acc_number or "", "account_type": account_type,
                "journal": ", ".join(journals.mapped("display_name")), "ledger_account": account.display_name,
                "account_id": account.id, "journal_ids": journals.ids,
                "control_journal_id": control_journal.id,
                "control_journal": control_journal.display_name or "",
                "currency": currency.name, "company_currency": company.currency_id.name,
                "base_url": self.env["ir.config_parameter"].sudo().get_param("web.base.url", ""),
            },
            "bank_opening": opening, "controls": controls, "ledger": ledger,
            "concepts": [{"code": item.code, "report_code": item.report_code or "", "name": item.name, "direction": item.direction,
                          "sequence": item.sequence, "detail": item.detail} for item in concepts],
            "movements": movements, "issues": issues,
        }

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
