"""Monthly sources for the XLSX form, separate from bank reconciliation.

Use one source per month. Ledger entries fill the form when bank transactions
are absent, but never establish a bank opening or an independently checked close.
"""

from .calculation import decimal, money


def form_snapshot(snapshot):
    accounting = {month["number"]: month for month in snapshot.get("accounting_months", [])}
    months, movements = [], []
    for original in snapshot["months"]:
        month = dict(original)
        number = month["number"]
        bank_rows = [row for row in snapshot["movements"] if int(row["date"][5:7]) == number]
        control = month.get("control") or {}
        bank_control = control.get("complete") and control.get("source_model") != "cq.bank.balance"
        use_books = number in accounting and not bank_rows and not bank_control
        month["flow_source"] = "accounting" if use_books else "bank"
        month["flows_available"] = bool(use_books or bank_rows or bank_control or month["bank_end"] is not None)
        opening = month.get("reported_bank_opening", month["bank_opening"])
        month["reported_bank_opening"] = opening
        if use_books:
            for key in ("income", "expense", "concept_totals"):
                month[key] = accounting[number][key]
            month["bank_end"] = None if opening is None else float(money(
                decimal(opening) + decimal(month["income"]) - decimal(month["expense"]), snapshot["rounding"]))
            # Reconciliation uses the separately informed bank closing, never
            # the ledger-derived A - B even when that result is now available.
            closing = month["statement_end"]
            month["bank_adjusted"] = None if closing is None else float(money(
                decimal(closing) + decimal(month["deposits"]) - decimal(month["checks"]) - decimal(month["payments"]), snapshot["rounding"]))
            month["difference"] = None if month["bank_adjusted"] is None else float(money(
                decimal(month["bank_adjusted"]) - decimal(month["book_adjusted"]), snapshot["rounding"]))
            month["bank_difference"] = None if month["bank_end"] is None or closing is None else float(money(
                decimal(month["bank_end"]) - decimal(closing), snapshot["rounding"]))
            running = None if opening is None else decimal(opening)
            for row in snapshot.get("accounting_movements", []):
                if int(row["date"][5:7]) != number:
                    continue
                if running is not None:
                    running += decimal(row["amount"])
                movements.append({
                    **row, "flow_source": "accounting", "source_model": "account.move.line",
                    "accounting_date": row["date"], "accounting_dates": row["date"],
                    "linked_documents": row["document"],
                    "counterparty_bank": row.get("counterparty_bank", ""),
                    "counterparty_account": row.get("counterparty_account", ""),
                    "note": "Apunte contable; fecha bancaria no disponible.",
                    "running_balance": None if running is None else float(money(running, snapshot["rounding"])),
                    "allocations": [{"code": row.get("concept_code", ""), "amount": abs(row["amount"]),
                                     "origin": row.get("classification_origin", "pending")}],
                })
        else:
            movements.extend({**row, "flow_source": "bank", "source_model": "account.bank.statement.line"} for row in bank_rows)
        months.append(month)
    return {**snapshot, "months": months, "movements": movements}
