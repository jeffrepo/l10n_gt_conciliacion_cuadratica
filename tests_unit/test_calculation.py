import copy
from datetime import date
from decimal import Decimal
import unittest

from core.calculation import build_snapshot, choose_rule, money, month_end, residual_at


def movement(ident, when, amount, code="IN_CUSTOMERS_LOCAL", allocations=None):
    return {
        "source_id": ident, "move_id": ident, "date": when, "amount": amount,
        "company_amount": amount, "document": "BANK/%s" % ident,
        "description": "Movimiento de prueba", "reference": "", "partner": "Cliente de prueba",
        "method": "Manual", "accounting_dates": when, "linked_documents": "",
        "counterparty_account": "", "counterparty_bank": "", "note": "",
        "allocations": allocations if allocations is not None else [{"code": code, "amount": abs(amount), "origin": "manual"}],
    }


def ledger(ident, when, amount, role="bank", matches=None, method="", company_amount=None):
    return {
        "source_id": ident, "move_id": ident, "date": when, "amount": amount,
        "company_amount": amount if company_amount is None else company_amount,
        "role": role, "matches": matches or [], "method_code": method,
        "method": method, "document": "MOVE/%s" % ident, "description": "Prueba",
        "partner": "Contraparte", "account": "Cuenta de prueba",
    }


def fixture():
    return {
        "year": 2024, "month": 1, "bank_opening": 1000, "rounding": "0.01",
        "company_rounding": "0.01",
        "metadata": {"company": "Empresa de prueba", "vat": "TEST", "country": "Guatemala",
                     "bank": "Banco de prueba", "account_number": "0000123", "account_type": "Monetaria",
                     "journal": "Banco", "ledger_account": "101 Banco", "currency": "GTQ",
                     "company_currency": "GTQ", "generated_at": "2024-02-01 12:00:00",
                     "generated_by": "Pruebas", "base_url": "https://odoo.example.test"},
        "concepts": [
            {"code": "IN_CUSTOMERS_LOCAL", "report_code": "I01", "name": "Clientes locales", "direction": "in", "sequence": 10, "detail": "none"},
            {"code": "OUT_SUPPLIERS", "report_code": "E01", "name": "Proveedores", "direction": "out", "sequence": 20, "detail": "none"},
            {"code": "OUT_FEES", "report_code": "E02", "name": "Comisiones", "direction": "out", "sequence": 30, "detail": "none"},
        ],
        "movements": [movement(1, "2024-01-05", 400), movement(2, "2024-01-20", -205, "OUT_SUPPLIERS")],
        "ledger": [ledger(10, "2023-12-31", 1000), ledger(11, "2024-01-05", 400), ledger(12, "2024-01-20", -205)],
        "controls": {"1": {"name": "ENERO", "source_id": 1, "date": "2024-01-31", "balance": 1195, "complete": True}},
    }


class TestCalculation(unittest.TestCase):
    def test_independent_bank_and_book_balances(self):
        result = build_snapshot(fixture())
        january = result["months"][0]
        self.assertEqual((january["income"], january["expense"], january["bank_end"]), (400, 205, 1195))
        self.assertEqual(january["book_adjusted"], 1195)
        self.assertEqual(january["difference"], 0)
        self.assertEqual(result["issues"], [])

    def test_book_difference_cannot_be_hidden_by_bank_balance(self):
        data = fixture()
        data["ledger"].append(ledger(20, "2024-01-25", -5))
        result = build_snapshot(data)
        self.assertEqual(result["months"][0]["difference"], 5)
        self.assertIn("book_difference", {item["code"] for item in result["issues"]})

    def test_statement_difference_is_independent(self):
        data = fixture()
        data["controls"]["1"]["balance"] = 1200
        result = build_snapshot(data)
        self.assertEqual(result["months"][0]["bank_difference"], -5)
        self.assertIn("bank_difference", {item["code"] for item in result["issues"]})

    def test_partial_check_cleared_in_february_remains_in_january(self):
        data = fixture()
        data["month"] = 2
        data["ledger"].extend([
            ledger(20, "2024-01-25", -100, "outstanding", [{"date": "2024-02-05", "delta": 40, "company_delta": 40}], "check_printing"),
            ledger(21, "2024-02-05", 40, "outstanding", [{"date": "2024-02-05", "delta": -40, "company_delta": -40}]),
            ledger(22, "2024-02-05", -40),
        ])
        data["movements"].append(movement(3, "2024-02-05", -40, "OUT_SUPPLIERS"))
        data["controls"]["2"] = {"name": "FEBRERO", "date": "2024-02-29", "source_id": 2, "balance": 1155, "complete": True}
        result = build_snapshot(data)
        january, february = result["months"]
        self.assertEqual((january["checks"], february["checks"]), (100, 60))
        self.assertEqual((january["book_adjusted"], february["book_adjusted"]), (1095, 1095))
        self.assertEqual(february["bank_opening"], january["bank_end"])
        self.assertEqual(result["issues"], [])

    def test_foreign_residual_keeps_both_historical_amounts(self):
        line = ledger(1, "2024-01-10", -100, "outstanding",
                      [{"date": "2024-02-10", "delta": 40, "company_delta": 312}], company_amount=-780)
        self.assertEqual(residual_at(line, "2024-01-31"), Decimal("-100"))
        self.assertEqual(residual_at(line, "2024-02-29"), Decimal("-60"))
        self.assertEqual(residual_at(line, "2024-02-29", company=True), Decimal("-468"))

    def test_suspense_explains_unrecorded_bank_income(self):
        data = fixture()
        data["ledger"].append(ledger(30, "2024-01-05", -400, "suspense"))
        january = build_snapshot(data)["months"][0]
        self.assertEqual(january["book_balance"], 795)
        self.assertEqual(january["book_adjustment"], 400)
        self.assertEqual(january["difference"], 0)

    def test_pending_deposit_and_other_payment_are_separate(self):
        data = fixture()
        data["ledger"] += [ledger(20, "2024-01-20", 50, "outstanding"), ledger(21, "2024-01-22", -30, "outstanding")]
        result = build_snapshot(data)
        self.assertEqual((result["months"][0]["deposits"], result["months"][0]["payments"]), (50, 30))
        self.assertEqual(result["months"][0]["difference"], 0)

    def test_missing_balances_stay_missing(self):
        data = fixture()
        data["bank_opening"] = None
        data["controls"] = {}
        result = build_snapshot(data)
        self.assertIsNone(result["months"][0]["bank_end"])
        self.assertIsNone(result["months"][0]["difference"])
        self.assertEqual({item["code"] for item in result["issues"]}, {"opening", "coverage"})

    def test_real_zero_is_not_missing(self):
        data = fixture()
        data.update(bank_opening=0, movements=[], ledger=[])
        data["controls"]["1"]["balance"] = 0
        result = build_snapshot(data)
        self.assertEqual(result["months"][0]["bank_end"], 0)
        self.assertEqual(result["issues"], [])

    def test_split_and_unclassified_amounts_are_counted_once(self):
        data = fixture()
        data["movements"][1]["allocations"] = [
            {"code": "OUT_SUPPLIERS", "amount": 180, "origin": "manual"},
            {"code": "OUT_FEES", "amount": 5, "origin": "manual"},
        ]
        result = build_snapshot(data)
        totals = result["months"][0]["concept_totals"]
        self.assertEqual((totals["OUT_SUPPLIERS"], totals["OUT_FEES"], totals["unclassified_out"]), (180, 5, 20))
        self.assertEqual(result["months"][0]["expense"], 205)
        self.assertEqual(len([i for i in result["issues"] if i["code"] == "classification"]), 1)

    def test_overallocation_and_wrong_direction_rejected(self):
        for allocation in [
            {"code": "OUT_SUPPLIERS", "amount": 206},
            {"code": "IN_CUSTOMERS_LOCAL", "amount": 205},
        ]:
            data = fixture()
            data["movements"][1]["allocations"] = [allocation]
            with self.assertRaises(ValueError):
                build_snapshot(data)

    def test_void_reversal_nets_without_dropping_audit_trail(self):
        data = fixture()
        data["movements"] += [movement(3, "2024-01-26", -10, "OUT_FEES"), movement(4, "2024-01-27", 10)]
        data["ledger"] += [ledger(23, "2024-01-26", -10), ledger(24, "2024-01-27", 10)]
        result = build_snapshot(data)
        self.assertEqual(len(result["movements"]), 4)
        self.assertEqual(result["months"][0]["bank_end"], 1195)

    def test_snapshot_does_not_mutate_inputs(self):
        data = fixture()
        original = copy.deepcopy(data)
        build_snapshot(data)
        self.assertEqual(data, original)

    def test_rounding_and_leap_year(self):
        self.assertEqual(money("0.025", "0.05"), Decimal("0.05"))
        self.assertEqual(month_end(2024, 2), date(2024, 2, 29))


class TestRules(unittest.TestCase):
    def match(self, rules, **kwargs):
        values = dict(direction="out", partner_id=12, country_scope="local", account_ids=[7], method_codes=["check_printing"], label="Comisión ACH")
        values.update(kwargs)
        return choose_rule(rules, **values)

    def rule(self, ident=1, sequence=10, concept="OUT_FEES", **kwargs):
        return {"id": ident, "sequence": sequence, "concept_code": concept, "direction": "out", **kwargs}

    def test_priority_not_first_record_wins(self):
        rule, ambiguity = self.match([self.rule(sequence=50), self.rule(2, 5, "OUT_SUPPLIERS", account_ids=[7])])
        self.assertEqual(rule["concept_code"], "OUT_SUPPLIERS")
        self.assertFalse(ambiguity)

    def test_equal_priority_conflict_needs_review(self):
        self.assertEqual(self.match([self.rule(), self.rule(2, concept="OUT_SUPPLIERS")]), (None, True))

    def test_all_conditions_must_match(self):
        rule = self.rule(account_ids=[7], partner_ids=[12], method_code="check_printing", label_contains="COMISIÓN")
        self.assertIsNotNone(self.match([rule])[0])
        self.assertIsNone(self.match([rule], partner_id=13)[0])

    def test_unknown_country_is_not_foreign(self):
        self.assertIsNone(self.match([self.rule(country_scope="foreign")], country_scope="unknown")[0])


if __name__ == "__main__":
    unittest.main()
