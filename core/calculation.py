"""Cash-book reconciliation using signed amounts in the bank currency.

Bank controls come from statements, never from the calculated bank balance.
Ledger balances include bank, outstanding and suspense accounts. Historical
residuals use the dates of BOTH matched entries, not today's residual field.
"""

import calendar
from datetime import date
from decimal import Decimal, ROUND_HALF_UP


MONTHS = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)


def decimal(value):
    return Decimal(str(value or 0))


def money(value, rounding="0.01"):
    step = decimal(rounding)
    return (decimal(value) / step).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * step


def month_end(year, month):
    return date(year, month, calendar.monthrange(year, month)[1])


def residual_at(line, cutoff, company=False):
    field = "company_amount" if company else "amount"
    delta = "company_delta" if company else "delta"
    return decimal(line[field]) + sum(
        (decimal(match[delta]) for match in line.get("matches", [])
         if match["date"] <= cutoff), Decimal("0"),
    )


def choose_rule(rules, direction, partner_id, country_scope, account_ids,
                method_codes, label):
    """Conditions within a rule are AND; selected accounts/partners are OR.

    Different concepts at the same winning priority are an ambiguity, not an
    arbitrary choice. Explicit source allocations take precedence over rules.
    """
    matches = []
    for rule in rules:
        if rule["direction"] != direction:
            continue
        if rule.get("partner_ids") and partner_id not in rule["partner_ids"]:
            continue
        if rule.get("account_ids") and not set(account_ids).intersection(rule["account_ids"]):
            continue
        if rule.get("country_scope", "any") not in ("any", country_scope):
            continue
        if rule.get("method_code") and rule["method_code"] not in method_codes:
            continue
        if rule.get("label_contains") and rule["label_contains"].casefold() not in label.casefold():
            continue
        matches.append(rule)
    if not matches:
        return None, False
    priority = min(rule["sequence"] for rule in matches)
    winners = [rule for rule in matches if rule["sequence"] == priority]
    if len({rule["concept_code"] for rule in winners}) != 1:
        return None, True
    return sorted(winners, key=lambda rule: rule["id"])[0], False


def build_snapshot(data):
    """Return a JSON-safe snapshot; input dates are ISO strings.

    No missing bank opening/control is converted into a successful zero.
    All movements remain visible, including unclassified and partially
    classified movements. Amounts allocated to concepts are counted once.
    """
    year, last_month = int(data["year"]), int(data["month"])
    if not 1900 <= year <= 9998 or not 1 <= last_month <= 12:
        raise ValueError("Año o mes fuera de rango.")
    data = _manual_opening_anchor(data)
    rounding = data.get("rounding", "0.01")
    start = date(year, 1, 1).isoformat()
    end = month_end(year, last_month).isoformat()
    issues = list(data.get("issues", []))
    concepts = {item["code"]: item for item in data["concepts"]}
    movements = []
    running = None if data.get("bank_opening") is None else decimal(data["bank_opening"])
    if running is None:
        issues.append({"code": "opening", "message": "Falta un extracto que respalde el saldo inicial bancario."})
    for source in data["movements"]:
        if not start <= source["date"] <= end:
            continue
        row = dict(source)
        amount = decimal(row["amount"])
        if running is not None:
            running += amount
        row["running_balance"] = None if running is None else float(money(running, rounding))
        direction = "in" if amount >= 0 else "out"
        allocations = []
        allocated = Decimal("0")
        for allocation in row.get("allocations", []):
            code = allocation["code"]
            value = decimal(allocation["amount"])
            if value <= 0 or code not in concepts or concepts[code]["direction"] != direction:
                raise ValueError("Distribución inválida del movimiento %s." % row["source_id"])
            allocated += value
            allocations.append({**allocation, "amount": float(value)})
        if money(allocated - abs(amount), rounding) > 0:
            raise ValueError("La distribución excede el movimiento %s." % row["source_id"])
        unallocated = money(abs(amount) - allocated, rounding)
        if unallocated:
            allocations.append({"code": "", "amount": float(unallocated), "origin": "pending"})
            issues.append({
                "code": "classification", "source_id": row["source_id"],
                "message": "Movimiento %s: %s sin clasificar." % (row["document"], unallocated),
            })
        row["allocations"] = allocations
        row["amount"] = float(amount)
        movements.append(row)

    months, pending = [], []
    bank_opening = None if data.get("bank_opening") is None else decimal(data["bank_opening"])
    for number in range(1, last_month + 1):
        cutoff = month_end(year, number).isoformat()
        month_start = date(year, number, 1).isoformat()
        rows = [row for row in movements if month_start <= row["date"] <= cutoff]
        income = sum((decimal(row["amount"]) for row in rows if row["amount"] > 0), Decimal("0"))
        expense = -sum((decimal(row["amount"]) for row in rows if row["amount"] < 0), Decimal("0"))
        bank_end = None if bank_opening is None else bank_opening + income - expense
        balances = {role: Decimal("0") for role in ("bank", "outstanding", "suspense")}
        deposits = checks = payments = suspense = Decimal("0")
        month_pending = []
        for line in data["ledger"]:
            if line["date"] > cutoff:
                continue
            role = line["role"]
            balances[role] += decimal(line["amount"])
            if role == "bank":
                continue
            residual = money(residual_at(line, cutoff), rounding)
            if not residual:
                continue
            if role == "suspense":
                kind = "suspense"
                suspense += residual
            elif residual > 0:
                kind = "deposit"
                deposits += residual
            elif line.get("method_code") == "check_printing":
                kind = "check"
                checks -= residual
            else:
                kind = "payment"
                payments -= residual
            month_pending.append({
                **{key: value for key, value in line.items() if key not in ("matches", "role")},
                "month": number, "cutoff": cutoff, "kind": kind,
                "amount": float(residual),
                "company_amount": float(money(residual_at(line, cutoff, company=True), data.get("company_rounding", "0.01"))),
            })
        pending.extend(month_pending)
        outstanding = deposits - checks - payments
        if money(outstanding - balances["outstanding"], rounding) or money(suspense - balances["suspense"], rounding):
            issues.append({"code": "clearing", "month": number,
                           "message": "%s: las partidas pendientes no explican los saldos de las cuentas transitorias." % MONTHS[number - 1]})
        book_balance = sum(balances.values(), Decimal("0"))
        book_adjustment = -suspense
        book_adjusted = book_balance + book_adjustment
        bank_adjusted = None if bank_end is None else bank_end + outstanding
        difference = None if bank_adjusted is None else bank_adjusted - book_adjusted
        control = data.get("controls", {}).get(str(number))
        control_value = decimal(control["balance"]) if control else None
        control_difference = None if control_value is None or bank_end is None else bank_end - control_value
        if not control:
            issues.append({"code": "coverage", "month": number,
                           "message": "%s: falta un extracto con fecha de corte %s." % (MONTHS[number - 1], cutoff)})
        elif not control.get("complete", False):
            issues.append({"code": "statement", "month": number,
                           "message": "%s: el extracto de control no está completo." % MONTHS[number - 1]})
        if control_difference is not None and money(control_difference, rounding):
            issues.append({"code": "bank_difference", "month": number,
                           "message": "%s: los movimientos no coinciden con el saldo del extracto." % MONTHS[number - 1]})
        if difference is not None and money(difference, rounding):
            issues.append({"code": "book_difference", "month": number,
                           "message": "%s: existe una diferencia entre banco ajustado y libros ajustados." % MONTHS[number - 1]})
        totals = {}
        for row in rows:
            for allocation in row["allocations"]:
                key = allocation["code"] or ("unclassified_in" if row["amount"] >= 0 else "unclassified_out")
                totals[key] = totals.get(key, Decimal("0")) + decimal(allocation["amount"])
        values = {
            "book_opening": sum((decimal(line["amount"]) for line in data["ledger"] if line["date"] < month_start), Decimal("0")),
            "bank_opening": bank_opening, "income": income, "expense": expense,
            "bank_end": bank_end, "statement_end": control_value,
            "bank_difference": control_difference, "deposits": deposits,
            "checks": checks, "payments": payments, "bank_adjusted": bank_adjusted,
            "ledger_bank": balances["bank"], "ledger_outstanding": balances["outstanding"],
            "ledger_suspense": balances["suspense"], "book_balance": book_balance,
            "book_adjustment": book_adjustment, "book_adjusted": book_adjusted,
            "difference": difference,
        }
        months.append({
            "number": number, "name": MONTHS[number - 1], "cutoff": cutoff,
            "control": control,
            **{key: None if value is None else float(money(value, rounding)) for key, value in values.items()},
            "concept_totals": {key: float(money(value, rounding)) for key, value in totals.items()},
        })
        bank_opening = bank_end
    result = {
        "schema_version": 1, "year": year, "month": last_month,
        "metadata": data["metadata"], "concepts": data["concepts"],
        "rounding": str(rounding), "months": months, "movements": movements,
        "pending": pending, "issues": issues,
    }
    if "accounting_movements" in data:
        # An independent ledger view, never added to the bank transactions.
        # It includes posted payments and manual entries even without imports.
        accounting = [dict(row) for row in data["accounting_movements"] if start <= row["date"] <= end]
        balance = sum((decimal(row["amount"]) for row in data["ledger"] if row["role"] == "bank" and row["date"] < start), Decimal("0"))
        accounting_months = []
        for month in months:
            opening = balance
            rows = [row for row in accounting if int(row["date"][5:7]) == month["number"]]
            totals = {}
            for row in rows:
                balance += decimal(row["amount"])
                row["running_balance"] = float(money(balance, rounding))
                code = row.get("concept_code") or ("unclassified_in" if row["amount"] >= 0 else "unclassified_out")
                totals[code] = totals.get(code, Decimal("0")) + abs(decimal(row["amount"]))
            accounting_months.append({
                "number": month["number"], "name": month["name"], "opening": float(money(opening, rounding)),
                "income": float(money(sum((max(decimal(row["amount"]), 0) for row in rows), Decimal("0")), rounding)),
                "expense": float(money(sum((max(-decimal(row["amount"]), 0) for row in rows), Decimal("0")), rounding)),
                "closing": float(money(balance, rounding)),
                "concept_totals": {code: float(money(amount, rounding)) for code, amount in totals.items()},
            })
        result.update(accounting_movements=accounting, accounting_months=accounting_months)
    _apply_manual_bank_balances(data, result)
    return result


def _manual_opening_anchor(data):
    """A recorded balance can anchor existing bank transactions, never books."""
    captures = data.get("manual_bank_balances", {})
    start = "%04d-01-01" % int(data["year"])
    end = month_end(int(data["year"]), int(data["month"])).isoformat()
    transactions = [row for row in data["movements"] if start <= row["date"] <= end]
    if data.get("bank_opening") is not None or not transactions or not captures:
        return data
    number = min(int(number) for number in captures)
    anchor_date = "%04d-%02d-01" % (int(data["year"]), number)
    preceding = sum((decimal(row["amount"]) for row in transactions if row["date"] < anchor_date), Decimal("0"))
    return {**data, "bank_opening": float(decimal(captures[str(number)]["opening"]) - preceding)}


def _apply_manual_bank_balances(data, result):
    captures = data.get("manual_bank_balances", {})
    if not captures:
        return
    rounding = result["rounding"]
    issues = result["issues"]
    previous_close, previous_manual = None, False

    def issue(code, month, message):
        if not any(item["code"] == code and item.get("month") == month["number"] for item in issues):
            issues.append({"code": code, "month": month["number"], "message": "%s: %s" % (month["name"], message)})

    for month in result["months"]:
        capture = captures.get(str(month["number"]))
        month["reported_bank_opening"] = month["bank_opening"]
        month["manual_opening_used"] = False
        month["manual_balance"] = dict(capture) if capture else None
        if capture:
            opening, closing = decimal(capture["opening"]), decimal(capture["closing"])
            if previous_close is not None and money(opening - decimal(previous_close), rounding):
                issue("manual_bank_continuity", month, "la apertura capturada no coincide con el saldo final informado del mes anterior.")
            if month["bank_opening"] is None:
                month["reported_bank_opening"] = float(money(opening, rounding))
                month["manual_opening_used"] = True
            elif money(decimal(month["bank_opening"]) - opening, rounding):
                issue("manual_bank_opening_conflict", month, "la apertura capturada difiere de la apertura reconstruida con movimientos bancarios. Se conserva el cálculo y se debe revisar la diferencia.")
            if month["control"]:
                if money(decimal(month["statement_end"]) - closing, rounding):
                    issue("manual_bank_closing_conflict", month, "el saldo final capturado difiere del extracto de Odoo. Se conserva el saldo del extracto.")
            else:
                month["statement_end"] = float(money(closing, rounding))
                month["control"] = {"source_id": capture["source_id"], "source_model": "cq.bank.balance",
                                    "name": "Captura manual: " + capture["name"], "date": month["cutoff"],
                                    "balance": month["statement_end"], "complete": True}
                issues[:] = [item for item in issues if not (item["code"] == "coverage" and item.get("month") == month["number"])]
                if month["bank_end"] is not None:
                    month["bank_difference"] = float(money(decimal(month["bank_end"]) - closing, rounding))
                    if month["bank_difference"]:
                        issue("bank_difference", month, "los movimientos bancarios no coinciden con el saldo final capturado.")
            if month["bank_end"] is None:
                # A declared closing is an independent bank balance. It does
                # not prove that deposits/withdrawals or running totals are 0.
                final = decimal(month["statement_end"])
                adjusted = final + decimal(month["deposits"]) - decimal(month["checks"]) - decimal(month["payments"])
                month["bank_adjusted"] = float(money(adjusted, rounding))
                month["difference"] = float(money(adjusted - decimal(month["book_adjusted"]), rounding))
                issue("bank_flow_missing", month, "se dispone de saldos bancarios capturados, pero faltan transacciones bancarias para reconstruir los depósitos y egresos del período. El saldo capturado no acredita esos movimientos.")
                if month["difference"]:
                    issue("book_difference", month, "existe una diferencia entre banco ajustado y libros ajustados.")
        elif month["reported_bank_opening"] is None and previous_manual and previous_close is not None:
            month["reported_bank_opening"] = previous_close
            month["manual_opening_used"] = True
        previous_close = month["statement_end"]
        previous_manual = (month["control"] or {}).get("source_model") == "cq.bank.balance"
    if result["months"][0]["reported_bank_opening"] is not None:
        issues[:] = [item for item in issues if item["code"] != "opening"]
