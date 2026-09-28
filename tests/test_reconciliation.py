import base64
from copy import deepcopy
from io import BytesIO

from openpyxl import load_workbook
from psycopg2 import IntegrityError

from odoo import Command
from odoo.addons.account.tests.common import AccountTestInvoicingCommon
from odoo.addons.l10n_gt_conciliacion_cuadratica.models.report import INTERNAL
from odoo.exceptions import AccessError, UserError, ValidationError
from odoo.tests import Form, tagged
from odoo.tests.common import new_test_user
from odoo.tools import mute_logger


@tagged("post_install", "-at_install")
class TestQuadraticReconciliation(AccountTestInvoicingCommon):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.company = cls.company_data["company"]
        cls.incoming = cls.env["account.account"].create({
            "name": "CQ pending receipts", "code": "CQ1001", "account_type": "asset_current",
            "reconcile": True, "company_ids": [Command.set(cls.company.ids)],
        })
        cls.outgoing = cls.env["account.account"].create({
            "name": "CQ pending payments", "code": "CQ1002", "account_type": "asset_current",
            "reconcile": True, "company_ids": [Command.set(cls.company.ids)],
        })
        cls.bank = cls.env["account.journal"].create({
            "name": "CQ test bank", "code": "CQBK", "type": "bank", "company_id": cls.company.id,
            "cq_enabled": True,
        })
        cls.bank.inbound_payment_method_line_ids.payment_account_id = cls.incoming
        cls.bank.outbound_payment_method_line_ids.payment_account_id = cls.outgoing
        cls.bank.bank_account_id = cls.env["res.partner.bank"].create({
            "acc_number": "TEST-CQ-0001", "partner_id": cls.company.partner_id.id,
            "company_id": cls.company.id,
            "bank_id": cls.env["res.bank"].create({"name": "Banco sintético CQ"}).id,
        })
        cls.in_concept = cls.env.ref("l10n_gt_conciliacion_cuadratica.concept_in_customers_local")
        cls.out_concept = cls.env.ref("l10n_gt_conciliacion_cuadratica.concept_out_suppliers")

    def _statement(self, when, amount, balance_start, balance_end, cutoff, classified=True, journal=None):
        journal = journal or self.bank
        line = self.env["account.bank.statement.line"].create({
            "journal_id": journal.id, "date": when, "payment_ref": "CQ transaction",
            "amount": amount, "partner_id": self.partner_a.id,
            "cq_concept_id": (self.in_concept if amount >= 0 else self.out_concept).id if classified else False,
        })
        if line.move_id.state == "draft":
            line.move_id.action_post()
        statement = self.env["account.bank.statement"].create({
            "name": "CQ %s" % cutoff, "line_ids": [Command.link(line.id)],
            "balance_start": balance_start, "balance_end_real": balance_end, "cq_date_end": cutoff,
        })
        return statement, line

    def _report(self, month=1):
        return self.env["cq.report"]._generate(self.company, self.bank.default_account_id, 2024, month)

    def _related_journal(self, code, **values):
        journal = self.env["account.journal"].create({
            "name": "CQ %s" % code, "code": code, "type": "bank", "company_id": self.company.id,
            "default_account_id": self.bank.default_account_id.id,
            "suspense_account_id": self.bank.suspense_account_id.id,
            **values,
        })
        journal.inbound_payment_method_line_ids.payment_account_id = self.incoming
        journal.outbound_payment_method_line_ids.payment_account_id = self.outgoing
        return journal

    def test_real_statement_report_export_and_immutable_close(self):
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        report = self._report()
        self.assertEqual(report.issue_count, 0, report.payload["issues"])
        self.assertEqual(report.month_ids.bank_end, 100)
        self.assertEqual(report.month_ids.book_balance, 0)
        self.assertEqual(report.month_ids.book_adjustment, 100)
        self.assertEqual(report.month_ids.difference, 0)
        self.assertTrue(report._xlsx_bytes().startswith(b"PK"))
        report.action_export()
        self.assertTrue(report.file_data)
        report.action_confirm()
        self.assertEqual(report.state, "confirmed")
        with self.assertRaises(AccessError):
            report.write({"payload": {"forged": True}})
        with self.assertRaises(AccessError):
            report.with_context(cq_internal=True).write({"issue_count": 0})
        with self.assertRaises(AccessError):
            report.month_ids.write({"income": 999})

    def test_classification_changes_require_new_snapshot(self):
        _, line = self._statement("2024-01-29", 100, 0, 100, "2024-01-31", classified=False)
        old = self._report()
        self.assertEqual(old.unclassified_count, 1)
        with self.assertRaises(UserError):
            old.action_confirm()
        line.cq_concept_id = self.in_concept
        new = self._report()
        self.assertEqual(new.unclassified_count, 0)
        self.assertEqual(old.unclassified_count, 1)
        self.assertNotEqual(old.id, new.id)

    def test_rules_and_split_constraints(self):
        _, line = self._statement("2024-01-29", 100, 0, 100, "2024-01-31", classified=False)
        self.env["cq.rule"].create({
            "name": "Known customer", "company_id": self.company.id, "journal_id": self.bank.id,
            "concept_id": self.in_concept.id, "partner_ids": [Command.set(self.partner_a.ids)],
        })
        self.assertEqual(self._report().unclassified_count, 0)
        allocation = self.env["cq.allocation"].create({
            "statement_line_id": line.id, "concept_id": self.in_concept.id, "amount": 60,
        })
        self.assertEqual(self._report().unclassified_count, 1)
        with self.assertRaises(ValidationError), self.cr.savepoint():
            allocation.amount = 101
        with self.assertRaises(ValidationError), self.cr.savepoint():
            allocation.concept_id = self.out_concept

    def test_partial_payment_after_cutoff(self):
        self._statement("2024-01-10", 100, 0, 100, "2024-01-31")
        payment_journal = self._related_journal("CQPM")
        payment = self.env["account.payment"].create({
            "journal_id": payment_journal.id, "date": "2024-01-20", "amount": 80,
            "payment_type": "outbound", "partner_type": "supplier", "partner_id": self.partner_a.id,
            "payment_method_line_id": payment_journal.outbound_payment_method_line_ids[0].id,
        })
        payment.action_post()
        _, bank_line = self._statement("2024-02-05", -30, 100, 70, "2024-02-29")
        counterpart = bank_line.move_id.line_ids.filtered(lambda aml: aml.account_id == self.bank.suspense_account_id)
        counterpart.account_id = self.outgoing
        payment_line = payment.move_id.line_ids.filtered(lambda aml: aml.account_id == self.outgoing)
        (payment_line | counterpart).reconcile()
        january = self._report()
        february = self._report(2)
        self.assertEqual(january.month_ids.payments, 80)
        feb = february.month_ids.filtered(lambda item: item.number == 2)
        self.assertEqual(feb.payments, 50)
        self.assertEqual(feb.difference, 0)
        self.assertEqual(february.issue_count, 0, february.payload["issues"])

    def test_missing_coverage_cannot_be_confirmed(self):
        statement, _ = self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        statement.cq_date_end = False
        report = self._report()
        self.assertFalse(report.month_ids.control_available)
        self.assertIn("coverage", {issue["code"] for issue in report.payload["issues"]})
        with self.assertRaises(UserError):
            report.action_confirm()

    def test_foreign_bank_currency_is_preserved(self):
        foreign = self.setup_other_currency("EUR", rates=[("1900-01-01", 0.5)])
        self.bank.currency_id = foreign
        self.incoming.currency_id = foreign
        self.outgoing.currency_id = foreign
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        report = self._report()
        self.assertEqual(report.currency_id, foreign)
        self.assertEqual(report.month_ids.income, 100)
        self.assertEqual(report.month_ids.bank_end, 100)
        self.assertEqual(report.month_ids.difference, 0)

    def test_shared_suspense_does_not_mix_banks(self):
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        other = self.env["account.journal"].create({
            "name": "CQ other bank", "code": "CQB2", "type": "bank", "company_id": self.company.id,
            "suspense_account_id": self.bank.suspense_account_id.id,
        })
        self.env["account.bank.statement.line"].create({
            "journal_id": other.id, "date": "2024-01-20", "payment_ref": "Other bank", "amount": 900,
        })
        report = self._report()
        self.assertEqual(report.month_ids.ledger_suspense, -100)
        self.assertEqual(report.month_ids.income, 100)

    def test_explicit_origin_precedes_current_account_mapping(self):
        other = self.env["account.journal"].create({
            "name": "CQ source bank", "code": "CQS2", "type": "bank", "company_id": self.company.id,
        })
        source = self.env["account.bank.statement.line"].create({
            "journal_id": other.id, "date": "2024-01-20", "payment_ref": "Original bank", "amount": 25,
        })
        line = source.move_id.line_ids[:1]
        owner = self.env["cq.report"]._line_owner(line, {self.bank.default_account_id.id}, {})
        self.assertEqual(owner, other.default_account_id.id)

    def test_wizard_view_and_company_boundary(self):
        with Form(self.env["cq.generate.wizard"]) as form:
            form.company_id = self.company
            form.year = 2024
            form.month = "1"
            form.account_ids.add(self.bank.default_account_id)
        wizard = form.save()
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        wizard.action_calculate()
        self.assertEqual(len(wizard.report_ids), 1)
        self.assertEqual(wizard.state, "result")
        other_company = self.env["res.company"].create({"name": "CQ isolated company"})
        with self.assertRaises(AccessError):
            self.env["cq.report"].with_context(allowed_company_ids=self.company.ids)._generate(other_company, self.bank.default_account_id, 2024, 1)

    def test_three_journals_one_account_and_one_opening_balance(self):
        checks = self._related_journal("CQCH")
        transfers = self._related_journal("CQTR", cq_statement_source=True)
        self._statement("2024-01-10", 100, 0, 100, "2024-01-10")
        self._statement("2024-01-20", -20, 100, 80, "2024-01-20", journal=checks)
        self._statement("2024-01-25", -30, 80, 50, "2024-01-31", journal=transfers)
        # Archived / non-enabled journals remain part of the account's history.
        checks.active = False
        wizard = self.env["cq.generate.wizard"].create({
            "company_id": self.company.id, "year": 2024, "month": "1", "all_accounts": True,
        })
        wizard.action_calculate()
        self.assertEqual(len(wizard.report_ids), 1)
        report = wizard.report_ids
        self.assertEqual(report.account_id, self.bank.default_account_id)
        self.assertEqual(set(report.journal_ids.ids), set((self.bank | checks | transfers).ids))
        self.assertEqual(report.journal_id, transfers)
        self.assertEqual(report.month_ids.bank_opening, 0)
        self.assertEqual(report.month_ids.income, 100)
        self.assertEqual(report.month_ids.expense, 50)
        self.assertEqual(report.month_ids.bank_end, 50)
        self.assertEqual(report.month_ids.ledger_bank, 50)
        self.assertEqual(report.month_ids.ledger_suspense, -50)
        self.assertEqual(len(report.payload["movements"]), 3)
        self.assertEqual(report.issue_count, 0, report.payload["issues"])
        from io import BytesIO
        from openpyxl import load_workbook
        workbook = load_workbook(BytesIO(report._xlsx_bytes()), data_only=True)
        self.assertEqual(workbook["Data"].max_row, 11)
        self.assertEqual(workbook["Data"]["W11"].value, transfers.display_name)
        self.assertTrue(any(transfers.display_name in str(cell.value or "") for row in workbook["Banco"] for cell in row))

    def test_payments_from_related_journals_and_other_bank_stay_separate(self):
        checks = self._related_journal("CQCH")
        transfers = self._related_journal("CQTR")
        self._statement("2024-01-10", 100, 0, 100, "2024-01-31")
        for journal, amount in [(checks, 20), (transfers, 30)]:
            payment = self.env["account.payment"].create({
                "journal_id": journal.id, "date": "2024-01-20", "amount": amount,
                "payment_type": "outbound", "partner_type": "supplier", "partner_id": self.partner_a.id,
                "payment_method_line_id": journal.outbound_payment_method_line_ids[0].id,
            })
            payment.action_post()
        other = self.env["account.journal"].create({
            "name": "Another ledger bank", "code": "CQB2", "type": "bank", "company_id": self.company.id,
            "cq_enabled": True, "suspense_account_id": self.bank.suspense_account_id.id,
        })
        other.outbound_payment_method_line_ids.payment_account_id = self.outgoing
        payment = self.env["account.payment"].create({
            "journal_id": other.id, "date": "2024-01-20", "amount": 900,
            "payment_type": "outbound", "partner_type": "supplier", "partner_id": self.partner_a.id,
            "payment_method_line_id": other.outbound_payment_method_line_ids[0].id,
        })
        payment.action_post()
        wizard = self.env["cq.generate.wizard"].create({
            "company_id": self.company.id, "year": 2024, "month": "1", "all_accounts": True,
        })
        wizard.action_calculate()
        self.assertEqual(len(wizard.report_ids), 2)
        report = wizard.report_ids.filtered(lambda item: item.account_id == self.bank.default_account_id)
        self.assertEqual(report.month_ids.payments, 50)
        self.assertEqual(report.month_ids.ledger_outstanding, -50)
        self.assertEqual(report.month_ids.bank_adjusted, 50)
        self.assertEqual(report.issue_count, 0, report.payload["issues"])
        other_report = wizard.report_ids - report
        self.assertEqual(other_report.month_ids.payments, 900)

    def test_multiple_statement_sources_require_explicit_control(self):
        other = self._related_journal("CQTR")
        self._statement("2024-01-10", 100, 0, 100, "2024-01-10")
        self._statement("2024-01-25", -30, 100, 70, "2024-01-31", journal=other)
        report = self._report()
        self.assertIn("statement_source", report.issue_ids.mapped("code"))
        self.assertIsNone(report.payload["months"][0]["bank_opening"])
        self.assertFalse(report.month_ids.control_available)
        self.assertEqual(report.month_ids.expense, 30)
        other.cq_statement_source = True
        self.assertEqual(self._report().issue_count, 0)
        with self.assertRaises(ValidationError), self.cr.savepoint():
            self.bank.cq_statement_source = True

    def test_rules_remain_scoped_to_source_journal(self):
        other = self._related_journal("CQTR", cq_statement_source=True)
        _, receipt = self._statement("2024-01-10", 100, 0, 100, "2024-01-10", classified=False)
        _, transfer = self._statement("2024-01-25", 30, 100, 130, "2024-01-31", classified=False, journal=other)
        other_concept = self.env.ref("l10n_gt_conciliacion_cuadratica.concept_in_interest")
        for journal, concept in [(self.bank, self.in_concept), (other, other_concept)]:
            self.env["cq.rule"].create({
                "name": journal.name, "company_id": self.company.id, "journal_id": journal.id,
                "concept_id": concept.id, "fallback": True,
            })
        report = self._report()
        self.assertEqual(report.issue_count, 0, report.payload["issues"])
        codes = {row["source_id"]: row["allocations"][0]["code"] for row in report.payload["movements"]}
        self.assertEqual(codes[receipt.id], self.in_concept.code)
        self.assertEqual(codes[transfer.id], other_concept.code)

    def test_bank_ledger_includes_general_journal_entries(self):
        self._statement("2024-01-10", 100, 0, 100, "2024-01-31")
        move = self.env["account.move"].create({
            "journal_id": self.company_data["default_journal_misc"].id, "date": "2024-01-20",
            "line_ids": [
                Command.create({"account_id": self.bank.default_account_id.id, "debit": 10}),
                Command.create({"account_id": self.company_data["default_account_revenue"].id, "credit": 10}),
            ],
        })
        move.action_post()
        report = self._report()
        self.assertEqual(report.month_ids.ledger_bank, 110)
        self.assertEqual(report.month_ids.difference, -10)
        self.assertIn("book_difference", report.issue_ids.mapped("code"))

    def _customer_payment(self, amount=100, when="2024-01-15", direct=False):
        self.company.partner_id.country_id = self.env.ref("base.gt")
        self.partner_a.country_id = self.env.ref("base.gt")
        method = self.bank.inbound_payment_method_line_ids[:1]
        if direct:
            method.payment_account_id = self.bank.default_account_id
        payment = self.env["account.payment"].create({
            "journal_id": self.bank.id, "date": when, "amount": amount,
            "payment_type": "inbound", "partner_type": "customer", "partner_id": self.partner_a.id,
            "payment_method_line_id": method.id,
        })
        payment.action_post()
        return payment

    def test_direct_payments_and_manual_entries_without_statements(self):
        payment = self._customer_payment(direct=True)
        self._related_journal("CQONE")
        self._related_journal("CQTWO")
        opening = self.env["account.move"].create({
            "journal_id": self.company_data["default_journal_misc"].id, "date": "2023-12-31",
            "line_ids": [
                Command.create({"account_id": self.bank.default_account_id.id, "debit": 500}),
                Command.create({"account_id": self.company_data["default_account_revenue"].id, "credit": 500}),
            ],
        })
        opening.action_post()
        # Draft and future entries never enter the snapshot.
        draft = opening.copy({"date": "2024-01-10"})
        future = opening.copy({"date": "2024-03-10"})
        future.action_post()
        report = self._report(2)
        self.assertEqual(report.payload["movements"], [])
        self.assertEqual(len(report.payload["accounting_movements"]), 1)
        row = report.payload["accounting_movements"][0]
        self.assertEqual(row["move_id"], payment.move_id.id)
        self.assertEqual(row["amount"], 100)
        self.assertEqual(row["concept_code"], "IN_CUSTOMERS_LOCAL")
        january, february = report.payload["accounting_months"]
        self.assertEqual((january["opening"], january["income"], january["closing"]), (500, 100, 600))
        self.assertEqual((february["opening"], february["income"], february["closing"]), (600, 0, 600))
        self.assertTrue(report.has_accounting_summary)
        self.assertFalse(report.has_bank_data)
        self.assertFalse(report.has_bank_movements)
        jan_ui, feb_ui = report.month_ids.sorted("number")
        self.assertTrue(jan_ui.accounting_available)
        self.assertEqual((jan_ui.accounting_opening, jan_ui.accounting_income, jan_ui.accounting_expense, jan_ui.accounting_closing), (500, 100, 0, 600))
        self.assertEqual((feb_ui.accounting_opening, feb_ui.accounting_income, feb_ui.accounting_closing), (600, 0, 600))
        self.assertEqual(jan_ui.accounting_customers_local, 100)
        self.assertFalse(jan_ui.bank_available)
        self.assertTrue(all(month["bank_opening"] is None for month in report.payload["months"]))
        self.assertIn("bank_transactions_missing", report.issue_ids.mapped("code"))
        lines = self.env["account.move.line"].search(report.action_accounting_lines()["domain"])
        self.assertEqual(lines, payment.move_id.line_ids.filtered(lambda aml: aml.account_id == self.bank.default_account_id))
        self.assertNotIn(draft.id, [row["move_id"] for row in report.payload["accounting_movements"]])
        workbook = load_workbook(BytesIO(report._xlsx_bytes()), data_only=True)
        self.assertEqual(workbook["Resumen contable Odoo"]["B6"].value, 100)
        self.assertEqual(workbook["Mayor bancario Odoo"]["I6"].value, 600)
        self.assertEqual(workbook["Banco"]["G19"].value, "n.d.")
        self.assertEqual(workbook.active.title, "Resumen contable Odoo")
        with self.assertRaises(UserError):
            report.action_confirm()

    def test_ui_summary_reads_conserved_snapshot_not_current_books(self):
        self._customer_payment(direct=True)
        original = self._report()
        payload = deepcopy(original.payload)
        original.action_export()
        exported = original.file_data
        self._customer_payment(amount=25)
        self.env.invalidate_all()
        self.assertEqual(original.month_ids.accounting_income, 100)
        self.assertEqual(original.month_ids.accounting_closing, 100)
        self.assertEqual(original.payload, payload)
        self.assertEqual(original.file_data, exported)
        self.assertEqual(self._report().month_ids.accounting_income, 125)

    def test_old_snapshot_without_accounting_summary_is_marked_unavailable(self):
        self._customer_payment(direct=True)
        report = self._report()
        payload = deepcopy(report.payload)
        del payload["accounting_months"]
        del payload["accounting_movements"]
        report.with_context(cq_internal=INTERNAL).write({"payload": payload})
        self.assertFalse(report.has_accounting_summary)
        self.assertFalse(report.month_ids.accounting_available)
        self.assertFalse(report.has_bank_data)
        self.assertEqual(report.payload, payload)

    def test_ui_bank_sources_distinguish_real_zero_from_missing_opening(self):
        _, line = self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        backed = self._report()
        self.assertTrue(backed.has_bank_data)
        self.assertTrue(backed.has_bank_movements)
        self.assertTrue(backed.month_ids.bank_available)
        self.assertTrue(backed.month_ids.control_available)
        self.assertEqual(backed.month_ids.bank_opening, 0)
        line.statement_id = False
        missing = self._report()
        self.assertTrue(missing.has_bank_data)  # Transactions still show 100.
        self.assertFalse(missing.month_ids.bank_available)
        self.assertFalse(missing.month_ids.control_available)
        self.assertEqual(missing.month_ids.income, 100)
        self.assertNotEqual(missing.month_ids.bank_source_status, backed.month_ids.bank_source_status)

    def test_customer_country_and_related_rule_precede_automatic_classification(self):
        self._customer_payment(direct=True)
        self.partner_a.country_id = self.env.ref("base.us")
        self.assertEqual(self._report().payload["accounting_movements"][0]["concept_code"], "IN_CUSTOMERS_FOREIGN")
        self.env["cq.rule"].create({
            "name": "Related customer", "company_id": self.company.id,
            "concept_id": self.env.ref("l10n_gt_conciliacion_cuadratica.concept_in_related").id,
            "partner_ids": [Command.set(self.partner_a.ids)],
        })
        row = self._report().payload["accounting_movements"][0]
        self.assertEqual((row["concept_code"], row["classification_origin"]), ("IN_RELATED", "rule"))

    def test_unknown_country_and_conflicting_rules_are_not_auto_classified(self):
        self._customer_payment(direct=True)
        self.partner_a.country_id = False
        self.assertEqual(self._report().payload["accounting_movements"][0]["concept_code"], "")
        self.partner_a.country_id = self.env.ref("base.gt")
        for code in ("concept_in_related", "concept_in_customers_local"):
            self.env["cq.rule"].create({
                "name": code, "company_id": self.company.id,
                "concept_id": self.env.ref("l10n_gt_conciliacion_cuadratica." + code).id,
                "partner_ids": [Command.set(self.partner_a.ids)], "sequence": 10,
            })
        self.assertEqual(self._report().payload["accounting_movements"][0]["concept_code"], "")

    def test_customer_bank_receipt_needs_complete_reconciliation_at_cutoff(self):
        self._customer_payment(amount=60)
        _, source = self._statement("2024-01-20", 100, 0, 100, "2024-01-31", classified=False)
        counterpart = source.move_id.line_ids.filtered(lambda aml: aml.account_id == self.bank.suspense_account_id)
        counterpart.account_id = self.incoming
        other = self.env["account.move.line"].search([("account_id", "=", self.incoming.id), ("debit", ">", 0), ("company_id", "=", self.company.id)])
        (counterpart | other).reconcile()
        self.assertEqual(self._report().unclassified_count, 1)
        later_payment = self._customer_payment(amount=40, when="2024-02-03")
        later_line = later_payment.move_id.line_ids.filtered(lambda aml: aml.account_id == self.incoming)
        (counterpart | later_line).reconcile()
        self.assertEqual(self._report().unclassified_count, 1)
        february = self._report(2)
        self.assertEqual(february.unclassified_count, 0)
        allocation = february.payload["movements"][0]["allocations"][0]
        self.assertEqual((allocation["code"], allocation["origin"]), ("IN_CUSTOMERS_LOCAL", "reconciliation"))
        self.assertEqual(february.payload["accounting_months"][0]["income"], 100)

    def test_mixed_receipt_does_not_classify_whole_amount_as_customer(self):
        self.company.partner_id.country_id = self.env.ref("base.gt")
        self.partner_a.country_id = self.env.ref("base.gt")
        move = self.env["account.move"].create({
            "journal_id": self.company_data["default_journal_misc"].id, "date": "2024-01-20",
            "line_ids": [
                Command.create({"account_id": self.bank.default_account_id.id, "debit": 100, "partner_id": self.partner_a.id}),
                Command.create({"account_id": self.partner_a.property_account_receivable_id.id, "credit": 80, "partner_id": self.partner_a.id}),
                Command.create({"account_id": self.company_data["default_account_revenue"].id, "credit": 20}),
            ],
        })
        move.action_post()
        row = self._report().payload["accounting_movements"][0]
        self.assertEqual(row["amount"], 100)
        self.assertEqual(row["concept_code"], "")

    def test_opening_respects_odoo_statement_sequence(self):
        self._statement("2024-01-10", 100, 20, 120, "2024-01-31")
        # A later-created line can precede the opening anchor on the same date
        # according to Odoo's statement sequence. ID order would add 20 twice.
        self.env["account.bank.statement.line"].create({
            "journal_id": self.bank.id, "date": "2024-01-10", "sequence": 100,
            "payment_ref": "Earlier bank transaction", "amount": 20,
            "cq_concept_id": self.in_concept.id,
        })
        report = self._report()
        self.assertEqual(report.month_ids.bank_opening, 0)
        self.assertEqual(report.month_ids.bank_end, 120)
        self.assertEqual(report.month_ids.bank_difference, 0)

    def test_group_rejects_mixed_currencies(self):
        foreign = self.setup_other_currency("EUR", rates=[("1900-01-01", 0.5)])
        self._related_journal("CQFX", currency_id=foreign.id)
        with self.assertRaises(UserError):
            self._report()

    def test_old_journal_snapshot_stays_readable(self):
        from ..models.report import INTERNAL
        self._statement("2024-01-10", 100, 0, 100, "2024-01-31")
        old = self._report()
        old.action_export()
        saved = old._xlsx_bytes()
        # Existing v1.0 records have only journal_id; update must not require
        # account_id, rewrite their payload, or regenerate their exported file.
        old.with_context(cq_internal=INTERNAL).write({"account_id": False, "journal_ids": [Command.clear()]})
        self.assertEqual(old._xlsx_bytes(), saved)
        action = old.action_new_version()
        self.assertEqual(action["context"]["default_account_ids"], [(6, 0, self.bank.default_account_id.ids)])

    def test_non_bank_and_wrong_allocation_are_rejected(self):
        with self.assertRaises(ValidationError), self.cr.savepoint():
            self.company_data["default_journal_misc"].cq_enabled = True
        _, line = self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        with self.assertRaises(ValidationError), self.cr.savepoint():
            line.cq_concept_id = self.out_concept

    def _restrict_existing_contact(self):
        other = self.env["res.company"].sudo().create({"name": "CQ contact isolation"})
        # Contacts can be restricted after historical entries were posted.
        self.partner_a.sudo().with_company(other).write({
            "company_id": other.id, "name": "Private CQ counterpart",
        })
        self.env.flush_all()
        self.env.invalidate_all()
        return other

    def test_restricted_contact_keeps_amounts_without_reading_foreign_data(self):
        _, source = self._statement("2024-01-29", 100, 0, 100, "2024-01-31", classified=False)
        self._restrict_existing_contact()
        model = self.env["cq.report"].with_user(self.env.user).with_context(allowed_company_ids=self.company.ids)
        foreign_partner = model.env["res.partner"].browse(self.partner_a.id)
        self.assertFalse(model.env.su)
        self.assertFalse(foreign_partner.has_access("read"))
        with self.assertRaises(AccessError):
            foreign_partner.check_access("read")
        self.assertTrue(source.with_env(model.env).has_access("read"))
        self.env["cq.rule"].create({
            "name": "Foreign customers only", "company_id": self.company.id,
            "concept_id": self.in_concept.id, "country_scope": "foreign",
        })
        report = model._generate(self.company, self.bank.default_account_id, 2024, 1)
        self.assertEqual(report.company_id, self.company)
        self.assertEqual(report.month_ids.income, 100)
        self.assertEqual(report.month_ids.ledger_bank, 100)
        self.assertEqual(report.month_ids.difference, 0)
        self.assertIn("partner_access", report.issue_ids.mapped("code"))
        self.assertEqual(len(report.issue_ids.filtered(lambda issue: issue.code == "partner_access")), 1)
        self.assertEqual(report.unclassified_count, 1)
        self.assertEqual(report.payload["movements"][0]["partner"], "Contacto restringido")
        self.assertNotIn("Private CQ counterpart", str(report.payload))
        self.assertEqual(model.env.companies, self.company)
        self.assertTrue(report._xlsx_bytes().startswith(b"PK"))
        with self.assertRaises(UserError):
            report.action_confirm()

    def test_report_company_is_independent_of_other_active_companies(self):
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        other = self._restrict_existing_contact()
        user = self.env.user
        user.sudo().write({"company_ids": [Command.link(other.id)]})
        model = self.env["cq.report"].with_user(user).with_context(allowed_company_ids=[other.id, self.company.id])
        # The contact is readable in the UI session, but not in a report scoped
        # to another company. The result must not depend on the top-bar choices.
        self.assertTrue(model.env["res.partner"].browse(self.partner_a.id).has_access("read"))
        wizard = model.env["cq.generate.wizard"].create({
            "company_id": self.company.id, "year": 2024, "month": "1",
            "account_ids": [Command.set(self.bank.default_account_id.ids)],
        })
        wizard.action_calculate()
        report = wizard.report_ids
        self.assertEqual(report.company_id, self.company)
        self.assertEqual(report.payload["metadata"]["company"], self.company.name)
        self.assertEqual(report.month_ids.income, 100)
        self.assertIn("partner_access", report.issue_ids.mapped("code"))
        self.assertNotIn("Private CQ counterpart", str(report.payload))
        self.assertEqual(model.env.company, other)
        self.assertEqual(set(model.env.companies.ids), {other.id, self.company.id})

    def test_metadata_does_not_prefetch_other_company_contacts(self):
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        other = self.env["res.company"].sudo().create({"name": "CQ unrelated metadata"})
        other.partner_id.sudo().company_id = other
        self.env.flush_all()
        self.env.invalidate_all()
        model = self.env["cq.report"].with_user(self.env.user).with_context(allowed_company_ids=self.company.ids)
        companies = model.env["res.company"].browse([other.id, self.company.id])
        report = model._generate(companies[1], self.bank.default_account_id, 2024, 1)
        self.assertEqual(report.issue_count, 0, report.payload["issues"])
        self.assertEqual(report.payload["metadata"]["company"], self.company.name)

    def test_restricted_commercial_contact_and_bank_details_are_not_exposed(self):
        parent = self.env["res.partner"].create({"name": "Hidden CQ parent", "is_company": True})
        child = self.env["res.partner"].create({"name": "Visible CQ child", "parent_id": parent.id})
        bank = self.env["res.partner.bank"].create({
            "acc_number": "PRIVATE-CQ-9988", "partner_id": child.id,
        })
        self.env["ir.rule"].sudo().create({
            "name": "CQ hidden commercial entity", "model_id": self.env.ref("base.model_res_partner").id,
            "domain_force": "[('id', '!=', %d)]" % parent.id,
        })
        model = self.env["cq.report"].with_user(self.env.user).with_context(allowed_company_ids=self.company.ids)
        self.assertTrue(child.with_env(model.env).has_access("read"))
        self.assertFalse(parent.with_env(model.env).has_access("read"))
        partners, banks = {}, {}
        info = model._partner_info(child, partners)
        self.assertTrue(info["restricted"])
        self.assertFalse(info["country_id"])
        self.assertFalse(info["partner_id"])
        bank_info = model._counterparty_bank_info(bank, partners, banks)
        self.assertTrue(bank_info["restricted"])
        self.assertFalse(bank_info["number"])
        self.assertFalse(bank_info["commercial_id"])

    def test_contact_prefetch_is_isolated_from_restricted_siblings(self):
        other = self.env["res.company"].sudo().create({"name": "CQ unrelated contact"})
        private = self.env["res.partner"].sudo().with_company(other).create({
            "name": "Private CQ unrelated", "company_id": other.id,
        })
        self.env.flush_all()
        self.env.invalidate_all()
        model = self.env["cq.report"].with_user(self.env.user).with_context(allowed_company_ids=self.company.ids)
        partners = model.env["res.partner"].browse([self.partner_a.id, private.id])
        info = model._partner_info(partners[0], {})
        self.assertFalse(info["restricted"])
        self.assertEqual(info["partner_id"], self.partner_a.id)
        self.assertEqual(info["name"], self.partner_a.display_name)

    def _template(self, note="Texto manual conservado"):
        from ..core.template import read_template
        raw = self._report()._xlsx_bytes()
        book = load_workbook(BytesIO(raw))
        end = read_template(raw)["form_end"]
        for merged in list(book["Banco"].merged_cells.ranges):
            if merged.min_row > end:
                book["Banco"].unmerge_cells(str(merged))
        book["Banco"].delete_rows(end + 1, book["Banco"].max_row)
        book.create_sheet("Notas manuales")["A1"] = note
        output = BytesIO()
        book.save(output)
        return self.env["cq.xlsx.template"].create({
            "name": "CQ plantilla", "company_id": self.company.id,
            "account_id": self.bank.default_account_id.id, "file_name": "formato.xlsx",
            "file_data": base64.b64encode(output.getvalue()),
        })

    def test_template_is_selected_by_account_and_copied_into_snapshot(self):
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        template = self._template()
        report = self._report()
        self.assertEqual(report.template_id, template)
        original_bytes = report.template_data
        book = load_workbook(BytesIO(base64.b64decode(template.file_data)))
        book["Notas manuales"]["A1"] = "Plantilla modificada después del cálculo"
        output = BytesIO()
        book.save(output)
        template.file_data = base64.b64encode(output.getvalue())
        self.assertEqual(report.template_data, original_bytes)
        exported = load_workbook(BytesIO(report._xlsx_bytes()), data_only=True)
        self.assertEqual(exported["Notas manuales"]["A1"].value, "Texto manual conservado")
        self.assertEqual(exported["Banco"]["G19"].value, 0)
        self.assertEqual(exported["Data"]["K9"].value, 100)
        self.assertEqual(exported["Banco"]["D12"].value, self.bank.default_account_id.code)
        updated = load_workbook(BytesIO(self._report()._xlsx_bytes()), data_only=True)
        self.assertEqual(updated["Notas manuales"]["A1"].value, "Plantilla modificada después del cálculo")

    def test_wizard_uses_explicit_template_and_rejects_other_account(self):
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        template = self._template()
        wizard = self.env["cq.generate.wizard"].create({
            "company_id": self.company.id, "year": 2024, "month": "1",
            "account_ids": [Command.set(self.bank.default_account_id.ids)], "template_id": template.id,
        })
        wizard.action_calculate()
        self.assertEqual(wizard.report_ids.template_id, template)
        another = self.env["account.journal"].create({"name": "CQ another bank", "code": "CQALT", "type": "bank", "company_id": self.company.id})
        template.account_id = another.default_account_id
        with self.assertRaises(UserError):
            wizard.action_calculate()
        with self.assertRaises(UserError):
            self.env["cq.report"]._generate(self.company, self.bank.default_account_id, 2024, 1, template=template)

    def test_invalid_template_is_rejected_before_generation(self):
        with self.assertRaises(ValidationError), self.cr.savepoint():
            self.env["cq.xlsx.template"].create({
                "name": "Invalid", "company_id": self.company.id,
                "account_id": self.bank.default_account_id.id, "file_name": "formato.xlsx",
                "file_data": base64.b64encode(b"not an xlsx"),
            })

    def test_missing_bank_name_is_not_replaced_by_ledger_account(self):
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        self.bank.bank_account_id.bank_id = False
        report = self._report()
        self.assertEqual(report.payload["metadata"]["bank"], "")
        self.assertIn("bank_name", report.issue_ids.mapped("code"))
        self.assertEqual(report.month_ids.bank_end, 100)
        self.assertEqual(report.payload["movements"][0]["accounting_date"], "2024-01-29")

    def _bank_capture(self, opening=0, closing=100, month="1", **values):
        return self.env["cq.bank.balance"].create({
            "name": "Estado de prueba", "company_id": self.company.id,
            "account_id": self.bank.default_account_id.id, "year": 2024, "month": month,
            "opening_balance": opening, "closing_balance": closing, **values,
        })

    def test_manual_balances_form_and_immutable_report_without_bank_transactions(self):
        self._customer_payment(direct=True)
        with Form(self.env["cq.bank.balance"]) as form:
            form.company_id = self.company
            form.account_id = self.bank.default_account_id
            form.year, form.month = 2024, "1"
            form.name = "Estado de prueba enero"
            form.opening_balance, form.closing_balance = 0, 100
        capture = form.save()
        self._bank_capture(month="2", closing=999)  # Beyond cutoff, never used.
        report = self._report()
        month = report.month_ids
        self.assertTrue(report.has_bank_data)
        self.assertFalse(report.has_bank_movements)
        self.assertFalse(month.bank_available)
        self.assertTrue(month.reported_opening_available)
        self.assertTrue(month.control_available)
        self.assertEqual((month.display_bank_opening, month.statement_end, month.difference), (0, 100, 0))
        self.assertTrue(month.difference_available)
        self.assertEqual(month.accounting_income, 100)
        self.assertEqual(report.payload["movements"], [])
        manual = report.payload["months"][0]["manual_balance"]
        self.assertEqual(manual["source_id"], capture.id)
        self.assertIn("bank_flow_missing", report.issue_ids.mapped("code"))
        with self.assertRaises(UserError):
            report.action_confirm()
        exported = report._xlsx_bytes()
        old_payload = deepcopy(report.payload)
        capture.closing_balance = 105
        updated = self._report()
        self.assertEqual(updated.month_ids.statement_end, 105)
        self.assertEqual(updated.month_ids.difference, 5)
        self.assertEqual(report.payload, old_payload)
        self.assertEqual(report._xlsx_bytes(), exported)
        self.assertEqual(report.month_ids.statement_end, 100)

    def test_manual_balances_do_not_replace_native_statements(self):
        self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        self._bank_capture(opening=10, closing=110)
        report = self._report()
        self.assertEqual((report.month_ids.display_bank_opening, report.month_ids.statement_end), (0, 100))
        self.assertTrue({"manual_bank_opening_conflict", "manual_bank_closing_conflict"} <= set(report.issue_ids.mapped("code")))
        self.assertEqual(report.payload["months"][0]["manual_balance"]["closing"], 110)

    def test_manual_capture_anchors_bank_transactions_without_statement(self):
        _, line = self._statement("2024-01-29", 100, 0, 100, "2024-01-31")
        line.statement_id = False
        self._bank_capture()
        report = self._report()
        self.assertEqual(report.month_ids.bank_end, 100)
        self.assertEqual(report.issue_count, 0, report.payload["issues"])
        self.assertEqual(report.payload["months"][0]["control"]["source_model"], "cq.bank.balance")
        report.action_confirm()
        self.assertEqual(report.state, "confirmed")

    def test_manual_zero_and_foreign_currency_are_preserved(self):
        foreign = self.setup_other_currency("EUR", rates=[("1900-01-01", 0.5)])
        self.bank.currency_id = foreign
        capture = self._bank_capture(closing=0)
        self.assertEqual(capture.currency_id, foreign)
        report = self._report()
        self.assertEqual(report.currency_id, foreign)
        self.assertTrue(report.month_ids.reported_opening_available)
        self.assertTrue(report.month_ids.control_available)
        self.assertEqual(report.month_ids.statement_end, 0)
        self.assertIsNone(report.payload["months"][0]["bank_end"])
        self.bank.currency_id = self.company.currency_id
        changed = self._report()
        self.assertIn("manual_bank_currency", changed.issue_ids.mapped("code"))
        self.assertFalse(changed.month_ids.control_available)

    def test_manual_capture_rejects_duplicates_wrong_account_and_year(self):
        self._bank_capture()
        with self.assertRaises((ValidationError, IntegrityError)), mute_logger("odoo.sql_db"), self.cr.savepoint():
            self._bank_capture()
        with self.assertRaises(ValidationError), self.cr.savepoint():
            self._bank_capture(year=1899)
        with self.assertRaises(ValidationError), self.cr.savepoint():
            self._bank_capture(account_id=self.incoming.id)

    def test_manual_capture_permissions_and_company_isolation(self):
        user = new_test_user(self.env, login="cq_capture_user", groups="account.group_account_user",
                             company_id=self.company.id, company_ids=[Command.set(self.company.ids)])
        captures = self.env["cq.bank.balance"].with_user(user).with_context(allowed_company_ids=self.company.ids)
        capture = captures.create({
            "name": "Estado capturado por contador", "company_id": self.company.id,
            "account_id": self.bank.default_account_id.id, "year": 2024, "month": "1",
            "opening_balance": 0, "closing_balance": 100,
        })
        capture.closing_balance = 101
        self.assertEqual(capture.closing_balance, 101)
        with self.assertRaises(AccessError):
            capture.unlink()
        other = self.env["res.company"].sudo().create({"name": "CQ capture isolation"})
        bank = self.env["account.journal"].sudo().with_company(other).create({
            "name": "Other capture bank", "code": "CQCAP", "type": "bank", "company_id": other.id,
        })
        private = self.env["cq.bank.balance"].sudo().with_company(other).create({
            "name": "Private bank statement", "company_id": other.id,
            "account_id": bank.default_account_id.id, "year": 2024, "month": "1",
            "opening_balance": 900, "closing_balance": 999,
        })
        self.assertNotIn(private.id, captures.search([]).ids)
        with self.assertRaises(AccessError):
            captures.browse(private.id).read(["opening_balance"])
        self.assertEqual(self._report().month_ids.statement_end, 101)
