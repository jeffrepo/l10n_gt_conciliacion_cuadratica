"""XLSX writer for immutable snapshots; no Odoo dependency."""

from collections import defaultdict
from datetime import date
from io import BytesIO
import re
import unicodedata

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

from .calculation import MONTHS, decimal, money


def safe_filename(value):
    ascii_value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9_-]+", "_", ascii_value).strip("_")[:150] or "conciliacion"


def export_xlsx(snapshot, confirmed=False):
    buffer = BytesIO()
    book = xlsxwriter.Workbook(buffer, {"in_memory": True, "strings_to_formulas": False, "strings_to_urls": False})
    book.set_properties({"title": "Conciliación cuadrática", "company": snapshot["metadata"]["company"]})
    fmt = {
        "title": book.add_format({"bold": True, "font_size": 14}),
        "head": book.add_format({"bold": True, "bg_color": "#183A50", "font_color": "#FFFFFF", "text_wrap": True}),
        "section": book.add_format({"bold": True, "bg_color": "#E8EEF2"}),
        "total": book.add_format({"bold": True, "top": 1, "text_wrap": True, "valign": "vcenter", "num_format": '#,##0.00;[Red](#,##0.00);"-"'}),
        "money": book.add_format({"num_format": '#,##0.00;[Red](#,##0.00);"-"', "valign": "vcenter"}),
        "date": book.add_format({"num_format": "dd/mm/yyyy"}),
        "text": book.add_format({"num_format": "@"}),
        "label": book.add_format({"text_wrap": True, "valign": "vcenter"}),
        "detail": book.add_format({"indent": 1, "font_color": "#555555", "text_wrap": True}),
        "warn": book.add_format({"bg_color": "#FFF2CC", "text_wrap": True}),
        "error": book.add_format({"bg_color": "#FCE4D6"}),
    }
    main = book.add_worksheet("Conciliación")
    details = book.add_worksheet("Movimientos")
    pending = book.add_worksheet("Partidas conciliatorias")
    for sheet in (main, details, pending):
        sheet.hide_gridlines(2)
        sheet.set_default_row(18)
        sheet.set_landscape()
        sheet.set_paper(8)
        sheet.set_footer("&CPágina &P de &N")
    main.fit_to_pages(1, 0)
    details.repeat_columns(0, 2)
    pending.repeat_columns(0, 3)
    _summary(main, snapshot, confirmed, fmt)
    _movements(details, snapshot, fmt)
    _pending(pending, snapshot, fmt)
    book.close()
    return buffer.getvalue()


def _summary(sheet, snapshot, confirmed, fmt):
    meta, months = snapshot["metadata"], snapshot["months"]
    sheet.set_column("A:A", 12)
    sheet.set_column("B:B", 65)
    sheet.set_column("C:O", 16)
    sheet.merge_range("A2:O2", "CONCILIACIÓN CUADRÁTICA MENSUAL POR CUENTA BANCARIA", fmt["title"])
    sheet.merge_range("A3:O3", "%s · NIT: %s · %s" % (meta["company"], meta["vat"], meta["country"]))
    sheet.merge_range("A4:O4", "%s · Cuenta: %s · %s" % (meta["bank"], meta["account_number"], meta["account_type"]))
    sheet.merge_range("A5:O5", "Cuenta contable: %s · Moneda del reporte: %s · Moneda contable: %s" % (meta["ledger_account"], meta["currency"], meta["company_currency"]))
    sheet.merge_range("A6:O6", "Año %s · Corte: %s · Generado: %s · %s" % (snapshot["year"], months[-1]["cutoff"], meta["generated_at"], meta["generated_by"]))
    state = "CIERRE CONSERVADO" if confirmed else "BORRADOR"
    sheet.merge_range("A7:O7", "%s · %s pendientes de revisión" % (state, len(snapshot["issues"])), fmt["warn"] if snapshot["issues"] else fmt["section"])
    sheet.merge_range("A8:O8", "Flujos: acumulados. Saldos: cierre del período. Saldo inicial anual: enero. n.d.: falta respaldo bancario.")
    sheet.write_row(9, 0, ["Código", "Concepto", *MONTHS, "Acum. / saldo al corte"], fmt["head"])
    sheet.set_row(9, 32)
    sheet.freeze_panes(10, 2)
    sheet.repeat_rows(0, 9)
    rows = {}
    row = 10

    def metric(key, label, annual="last", formula=None, total=False):
        nonlocal row
        rows[key] = row
        sheet.write_string(row, 1, label, fmt["total"] if total else fmt["label"])
        sheet.set_row(row, 18 * max(1, (len(label) + 57) // 58))
        for idx, result in enumerate(months):
            col = idx + 2
            value = result[key]
            if value is None:
                sheet.write_string(row, col, "n.d.")
            elif formula:
                sheet.write_formula(row, col, formula(col, result), fmt["total"] if total else fmt["money"], value)
            else:
                sheet.write_number(row, col, value, fmt["total"] if total else fmt["money"])
        values = [item[key] for item in months]
        selected = values[0] if annual == "first" else values[-1]
        if annual == "sum":
            selected = sum(values) if all(value is not None for value in values) else None
            expression = "=SUM(%s:%s)" % (xl_rowcol_to_cell(row, 2), xl_rowcol_to_cell(row, len(months) + 1))
        else:
            expression = "=" + xl_rowcol_to_cell(row, 2 if annual == "first" else len(months) + 1)
        if selected is not None:
            sheet.write_formula(row, 14, expression, fmt["total"] if total else fmt["money"], selected)
        else:
            sheet.write_string(row, 14, "n.d.")
        row += 1

    def cell(key, col):
        return xl_rowcol_to_cell(rows[key], col)

    def section(label):
        nonlocal row
        sheet.merge_range(row, 0, row, 14, label, fmt["section"])
        row += 1

    def concept_rows(direction):
        nonlocal row
        parent_rows = []
        concepts = [item for item in snapshot["concepts"] if item["direction"] == direction]
        concepts.append({"code": "unclassified_" + direction, "name": "Pendiente de clasificar", "detail": "none"})
        for concept in concepts:
            parent = row
            parent_rows.append(parent)
            if concept.get("report_code"):
                sheet.write_string(parent, 0, concept["report_code"], fmt["text"])
            sheet.write_string(parent, 1, concept["name"], fmt["label"])
            sheet.set_row(parent, 18 * max(1, (len(concept["name"]) + 57) // 58))
            for idx, result in enumerate(months):
                value = result["concept_totals"].get(concept["code"], 0)
                sheet.write_number(parent, idx + 2, value, fmt["money"])
            total = sum(result["concept_totals"].get(concept["code"], 0) for result in months)
            sheet.write_formula(parent, 14, "=SUM(%s:%s)" % (xl_rowcol_to_cell(parent, 2), xl_rowcol_to_cell(parent, len(months) + 1)), fmt["money"], total)
            row += 1
            if concept["detail"] != "none":
                groups = defaultdict(lambda: defaultdict(float))
                for movement in snapshot["movements"]:
                    for allocation in movement["allocations"]:
                        if allocation["code"] != concept["code"]:
                            continue
                        label = movement["partner"] or "Contraparte no identificada"
                        if concept["detail"] == "bank":
                            label = " · ".join(filter(None, [label, movement["counterparty_bank"], movement["counterparty_account"]]))
                            if not movement["counterparty_account"]:
                                label += " · Cuenta no identificada"
                        groups[label][int(movement["date"][5:7])] += allocation["amount"]
                grouped = sorted(groups.items(), key=lambda item: (-sum(item[1].values()), item[0]))
                if concept["code"] == "OUT_DIVIDENDS" and len(grouped) > 10:
                    other = defaultdict(float)
                    for _, totals in grouped[10:]:
                        for number, value in totals.items():
                            other[number] += value
                    grouped = grouped[:10] + [("Otros socios", other)]
                for label, totals in grouped:
                    sheet.write_string(row, 1, label, fmt["detail"])
                    sheet.set_row(row, 18 * max(1, (len(label) + 55) // 56))
                    for idx in range(len(months)):
                        sheet.write_number(row, idx + 2, totals.get(idx + 1, 0), fmt["money"])
                    sheet.write_formula(row, 14, "=SUM(%s:%s)" % (xl_rowcol_to_cell(row, 2), xl_rowcol_to_cell(row, len(months) + 1)), fmt["money"], sum(totals.values()))
                    row += 1
        return parent_rows

    metric("bank_opening", "Saldo inicial según banco", annual="first")
    section("INGRESOS BANCARIOS")
    income_rows = concept_rows("in")
    metric("income", "Total ingresos del período", annual="sum", total=True,
           formula=lambda col, _: "=SUM(%s)" % ",".join(xl_rowcol_to_cell(item, col) for item in income_rows))
    section("EGRESOS BANCARIOS")
    expense_rows = concept_rows("out")
    metric("expense", "Total egresos del período", annual="sum", total=True,
           formula=lambda col, _: "=SUM(%s)" % ",".join(xl_rowcol_to_cell(item, col) for item in expense_rows))
    metric("bank_end", "Saldo final bancario calculado", total=True,
           formula=lambda col, _: "=%s+%s-%s" % (cell("bank_opening", col), cell("income", col), cell("expense", col)))
    # Carry forward the previous closing balance; never sum monthly balances.
    for idx in range(1, len(months)):
        if months[idx]["bank_opening"] is not None:
            sheet.write_formula(rows["bank_opening"], idx + 2, "=" + cell("bank_end", idx + 1), fmt["money"], months[idx]["bank_opening"])
    metric("statement_end", "Saldo final del extracto de control")
    metric("bank_difference", "Diferencia contra extracto", total=True,
           formula=lambda col, _: "=%s-%s" % (cell("bank_end", col), cell("statement_end", col)))
    section("PARTIDAS CONCILIATORIAS AL CIERRE")
    metric("deposits", "(+) Depósitos en tránsito")
    metric("checks", "(-) Cheques en circulación")
    metric("payments", "(-) Otros pagos pendientes de cargo")
    metric("bank_adjusted", "Saldo bancario ajustado", total=True,
           formula=lambda col, _: "=%s+%s-%s-%s" % tuple(cell(key, col) for key in ("bank_end", "deposits", "checks", "payments")))
    section("SALDO SEGÚN LIBROS")
    metric("ledger_bank", "Mayor de la cuenta bancaria")
    metric("ledger_outstanding", "Mayor de las cuentas pendientes atribuibles a este banco")
    metric("ledger_suspense", "Mayor de la cuenta transitoria atribuible a este banco")
    metric("book_balance", "Saldo según libros (banco + pendientes + transitoria)", total=True,
           formula=lambda col, _: "=%s+%s+%s" % tuple(cell(key, col) for key in ("ledger_bank", "ledger_outstanding", "ledger_suspense")))
    metric("book_adjustment", "Ajustes identificados de libros (opuesto de transitoria pendiente)")
    metric("book_adjusted", "Saldo de libros ajustado", total=True,
           formula=lambda col, _: "=%s+%s" % (cell("book_balance", col), cell("book_adjustment", col)))
    metric("difference", "Diferencia banco ajustado menos libros ajustados", total=True,
           formula=lambda col, _: "=%s-%s" % (cell("bank_adjusted", col), cell("book_adjusted", col)))
    for key in ("bank_difference", "difference"):
        sheet.conditional_format(rows[key], 2, rows[key], 14, {"type": "cell", "criteria": "not between", "minimum": -float(snapshot["rounding"]) / 2, "maximum": float(snapshot["rounding"]) / 2, "format": fmt["error"]})
    row += 1
    section("EXTRACTOS DE CONTROL")
    for month in months:
        control = month["control"]
        sheet.write_string(row, 1, "%s: %s" % (month["name"], control["name"] if control else "Sin extracto al cierre"))
        sheet.write_string(row, 2, month["cutoff"])
        row += 1
    if snapshot["issues"]:
        section("PENDIENTES DE REVISIÓN")
        for issue in snapshot["issues"]:
            sheet.merge_range(row, 1, row, 14, issue["message"], fmt["warn"])
            sheet.set_row(row, 32)
            row += 1
    sheet.print_area(0, 0, row, 14)


def _link(sheet, row, col, snapshot, model, source_id, label="Abrir en Odoo"):
    base = snapshot["metadata"].get("base_url", "").rstrip("/")
    if base.startswith(("https://", "http://")):
        sheet.write_url(row, col, "%s/web#id=%s&model=%s&view_type=form" % (base, source_id, model), string=label)


def _movements(sheet, snapshot, fmt):
    headers = ["Fecha banco", "Fechas documentos vinculados", "Documento Odoo", "Referencia", "Contraparte", "Concepto / descripción", "Medio de pago", "Código", "Clasificación", "Ingreso", "Egreso", "Saldo bancario", "Importe contable", "Moneda contable", "Documentos vinculados", "Banco contraparte", "Cuenta contraparte", "Observación", "Origen clasificación", "Odoo"]
    sheet.write_row(0, 0, headers, fmt["head"])
    sheet.set_row(0, 34)
    sheet.set_column(0, 1, 18)
    sheet.set_column(2, 4, 28)
    sheet.set_column(5, 5, 55)
    sheet.set_column(6, 8, 30)
    sheet.set_column(9, 13, 18)
    sheet.set_column(14, 18, 35)
    sheet.set_column(19, 19, 18)
    sheet.freeze_panes(1, 0)
    sheet.repeat_rows(0)
    concepts = {item["code"]: item["name"] for item in snapshot["concepts"]}
    report_codes = {item["code"]: item.get("report_code") or item["code"] for item in snapshot["concepts"]}
    row = 1
    for movement in snapshot["movements"]:
        allocations = movement["allocations"] or [{"code": "", "amount": 0, "origin": "pending"}]
        company_remaining = decimal(movement["company_amount"])
        for idx, allocation in enumerate(allocations):
            last = idx == len(allocations) - 1
            company_value = company_remaining if last else money(decimal(movement["company_amount"]) * decimal(allocation["amount"]) / abs(decimal(movement["amount"])))
            company_remaining -= company_value
            code = allocation["code"]
            origin = {"manual": "Manual", "rule": "Regla", "pending": "Pendiente"}.get(allocation["origin"], "")
            sheet.write_datetime(row, 0, date.fromisoformat(movement["date"]), fmt["date"])
            values = {
                1: movement["accounting_dates"], 2: movement["document"], 3: movement["reference"],
                4: movement["partner"], 5: movement["description"], 6: movement["method"],
                7: report_codes.get(code, ""), 8: concepts.get(code, "Pendiente de clasificar"),
                13: snapshot["metadata"]["company_currency"], 14: movement["linked_documents"],
                15: movement["counterparty_bank"], 16: movement["counterparty_account"],
                17: allocation.get("note") or movement["note"], 18: origin,
            }
            for col, value in values.items():
                if value:
                    sheet.write_string(row, col, value, fmt["text"])
                else:
                    sheet.write_blank(row, col, None, fmt["text"])
            sheet.write_number(row, 9, allocation["amount"] if movement["amount"] >= 0 else 0, fmt["money"])
            sheet.write_number(row, 10, allocation["amount"] if movement["amount"] < 0 else 0, fmt["money"])
            if last and movement["running_balance"] is not None:
                sheet.write_number(row, 11, movement["running_balance"], fmt["money"])
            sheet.write_number(row, 12, float(company_value), fmt["money"])
            _link(sheet, row, 19, snapshot, "account.bank.statement.line", movement["source_id"])
            row += 1
    sheet.autofilter(0, 0, max(row - 1, 1), len(headers) - 1)


def _pending(sheet, snapshot, fmt):
    headers = ["Corte", "Fecha contable", "Partida", "Documento", "Contraparte", "Descripción", "Cuenta contable", "Medio de pago", "Pendiente moneda banco", "Moneda banco", "Pendiente moneda compañía", "Moneda compañía", "Odoo"]
    labels = {"deposit": "Depósito en tránsito", "check": "Cheque en circulación", "payment": "Otro pago pendiente", "suspense": "Movimiento en transitoria"}
    sheet.write_row(0, 0, headers, fmt["head"])
    sheet.set_row(0, 36)
    sheet.set_column(0, 1, 14)
    sheet.set_column(2, 4, 30)
    sheet.set_column(5, 6, 50)
    sheet.set_column(7, 11, 24)
    sheet.set_column(12, 12, 18)
    sheet.freeze_panes(1, 0)
    sheet.repeat_rows(0)
    for row, item in enumerate(snapshot["pending"], 1):
        sheet.write_datetime(row, 0, date.fromisoformat(item["cutoff"]), fmt["date"])
        sheet.write_datetime(row, 1, date.fromisoformat(item["date"]), fmt["date"])
        for col, value in {2: labels[item["kind"]], 3: item["document"], 4: item["partner"], 5: item["description"], 6: item["account"], 7: item["method"], 9: snapshot["metadata"]["currency"], 11: snapshot["metadata"]["company_currency"]}.items():
            sheet.write_string(row, col, value or "", fmt["text"])
        sheet.write_number(row, 8, item["amount"], fmt["money"])
        sheet.write_number(row, 10, item["company_amount"], fmt["money"])
        _link(sheet, row, 12, snapshot, "account.move", item["move_id"])
    sheet.autofilter(0, 0, max(len(snapshot["pending"]), 1), len(headers) - 1)
