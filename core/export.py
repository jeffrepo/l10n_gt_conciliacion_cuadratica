"""XLSX writer for immutable snapshots; no Odoo dependency."""

from datetime import date
from io import BytesIO
import re
import unicodedata

import xlsxwriter

from .calculation import MONTHS, decimal, money
from .reference import reference_formats, summary


def safe_filename(value):
    ascii_value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9_-]+", "_", ascii_value).strip("_")[:150] or "conciliacion"


def export_xlsx(snapshot, confirmed=False, template=None):
    profile = None
    if template:
        from .template import read_template
        profile = read_template(template)
    buffer = BytesIO()
    book = xlsxwriter.Workbook(buffer, {"in_memory": True, "strings_to_formulas": False, "strings_to_urls": False})
    book.set_properties({"title": "Conciliación cuadrática", "company": snapshot["metadata"]["company"]})
    fmt = reference_formats(book)
    accounting_first = "accounting_movements" in snapshot and not snapshot["movements"]
    if accounting_first:
        _accounting(book, snapshot, fmt)
    main = book.add_worksheet("Banco")
    details = book.add_worksheet("Data")
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
    layout = summary(main, snapshot, confirmed, fmt, profile)
    _movements(details, snapshot, fmt)
    _pending(pending, snapshot, fmt)
    if "accounting_movements" in snapshot and not accounting_first:
        _accounting(book, snapshot, fmt)
    book.close()
    result = buffer.getvalue()
    if template:
        from .template import fill_template
        result = fill_template(profile, result, layout, snapshot)
    return result


def _link(sheet, row, col, snapshot, model, source_id, label="Abrir en Odoo"):
    base = snapshot["metadata"].get("base_url", "").rstrip("/")
    if base.startswith(("https://", "http://")):
        sheet.write_url(row, col, "%s/web#id=%s&model=%s&view_type=form" % (base, source_id, model), string=label)


def _movements(sheet, snapshot, fmt):
    meta = snapshot["metadata"]
    sheet.write_string("D1", "EMPRESA: " + meta["company"], fmt["label"])
    sheet.write_string("D2", "MOVIMIENTO BANCARIO AÑO %s" % snapshot["year"], fmt["label"])
    sheet.write_string("D3", meta["bank"] or "Banco no identificado", fmt["label"])
    headers = ["Mes de cobro", "Código", "Movimiento", "Fecha Contabilidad", "DOCTO #", "NOMBRE", "Concepto", "REF.", "INGRESO", "EGRESO", "SALDO", "Fecha banco", "Fechas documentos vinculados", "Clasificación", "Importe contable", "Moneda contable", "Documentos vinculados", "Banco contraparte", "Cuenta contraparte", "Observación", "Origen clasificación", "Odoo", "Diario de origen"]
    widths = [9.57, 8.14, 10.71, 14, 20, 30.14, 53.14, 11.43, 14.43, 16.29, 17.29]
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    sheet.set_column(11, 22, 24)
    sheet.write_row(6, 0, headers, fmt["data_head"])
    sheet.set_row(6, 30)
    sheet.write_string("G8", "Saldo inicial", fmt["label"])
    opening = snapshot["months"][0]["bank_opening"]
    if opening is None:
        sheet.write_string("K8", "n.d.", fmt["warn"])
    else:
        sheet.write_number("K8", opening, fmt["money"])
    sheet.freeze_panes(8, 0)
    sheet.repeat_rows(0, 6)
    concepts = {item["code"]: item["name"] for item in snapshot["concepts"]}
    report_codes = {item["code"]: item.get("report_code") or item["code"] for item in snapshot["concepts"]}
    row = 8
    for movement in snapshot["movements"]:
        allocations = movement["allocations"] or [{"code": "", "amount": 0, "origin": "pending"}]
        company_remaining = decimal(movement["company_amount"])
        for idx, allocation in enumerate(allocations):
            last = idx == len(allocations) - 1
            company_value = company_remaining if last else money(decimal(movement["company_amount"]) * decimal(allocation["amount"]) / abs(decimal(movement["amount"])))
            company_remaining -= company_value
            code = allocation["code"]
            origin = _origin(allocation["origin"])
            sheet.write_datetime(row, 3, date.fromisoformat(movement.get("accounting_date") or movement["date"]), fmt["date"])
            sheet.write_datetime(row, 11, date.fromisoformat(movement["date"]), fmt["date"])
            values = {
                0: MONTHS[int(movement["date"][5:7]) - 1].upper(), 1: report_codes.get(code, ""),
                2: movement["method"], 4: movement["document"], 5: movement["partner"],
                6: movement["description"], 7: movement["reference"],
                12: movement["accounting_dates"], 13: concepts.get(code, "Pendiente de clasificar"),
                15: meta["company_currency"], 16: movement["linked_documents"],
                17: movement["counterparty_bank"], 18: movement["counterparty_account"],
                19: allocation.get("note") or movement["note"], 20: origin,
                22: movement.get("journal", ""),
            }
            for col, value in values.items():
                sheet.write_string(row, col, value or "", fmt["text"])
            sheet.write_number(row, 8, allocation["amount"] if movement["amount"] >= 0 else 0, fmt["money"])
            sheet.write_number(row, 9, allocation["amount"] if movement["amount"] < 0 else 0, fmt["money"])
            if last:
                if movement["running_balance"] is None:
                    sheet.write_string(row, 10, "n.d.", fmt["warn"])
                else:
                    sheet.write_number(row, 10, movement["running_balance"], fmt["money"])
            sheet.write_number(row, 14, float(company_value), fmt["money"])
            _link(sheet, row, 21, snapshot, "account.bank.statement.line", movement["source_id"])
            row += 1
    sheet.autofilter(6, 0, max(row - 1, 7), len(headers) - 1)


def _origin(origin):
    return {"manual": "Manual", "rule": "Regla", "pending": "Pendiente",
            "reconciliation": "Contrapartida de cliente"}.get(origin, "")


def _accounting(book, snapshot, fmt):
    """Separate evidence from the GL; never represent it as a bank statement."""
    summary = book.add_worksheet("Resumen contable Odoo")
    detail = book.add_worksheet("Mayor bancario Odoo")
    meta = snapshot["metadata"]
    for sheet in (summary, detail):
        sheet.hide_gridlines(2)
        sheet.set_landscape()
        sheet.fit_to_pages(1, 0)
        sheet.merge_range("A1:M1", "MOVIMIENTOS SEGÚN LIBROS · %s · %s" % (meta["company"], meta["ledger_account"]), fmt["head"])
        sheet.merge_range("A2:M2", "Fuente: apuntes publicados de la cuenta contable bancaria, incluidos pagos, cobros y asientos. No acredita saldos según banco.", fmt["warn"])
        sheet.set_row(0, 30)
        sheet.set_row(1, 32)
    months = snapshot["accounting_months"]
    summary.set_column(0, 0, 63)
    summary.set_column(1, 13, 16)
    summary.write_row(3, 0, ["Concepto (%s)" % meta["currency"]] + list(MONTHS) + ["TOTAL / SALDO AL CORTE"], fmt["head"])
    summary.set_row(3, 32)
    entries = [("opening", "SALDO INICIAL SEGÚN LIBROS (CUENTA BANCARIA)"),
               ("income", "ENTRADAS / DÉBITOS CONTABLES"), ("expense", "SALIDAS / CRÉDITOS CONTABLES"),
               ("closing", "SALDO FINAL SEGÚN LIBROS (CUENTA BANCARIA)")]
    for idx, (key, label) in enumerate(entries, 4):
        summary.write_string(idx, 0, label, fmt["concept"])
        for col, month in enumerate(months, 1):
            summary.write_number(idx, col, month[key], fmt["money"])
        total = months[0][key] if key == "opening" else months[-1][key]
        if key in ("income", "expense"):
            total = float(sum((decimal(month[key]) for month in months), decimal(0)))
        summary.write_number(idx, 13, total, fmt["total"])
    row = 9
    for direction, label in (("in", "DETALLE DE ENTRADAS CONTABLES"), ("out", "DETALLE DE SALIDAS CONTABLES")):
        summary.merge_range(row, 0, row, 13, label, fmt["head"])
        row += 1
        concepts = [(item["code"], item["name"]) for item in snapshot["concepts"] if item["direction"] == direction]
        concepts.append(("unclassified_" + direction, "Pendiente de clasificar / distribución en Data bancaria"))
        for code, name in concepts:
            summary.write_string(row, 0, name, fmt["concept"])
            summary.set_row(row, 30 if len(name) > 58 else 18)
            total = decimal(0)
            for col, month in enumerate(months, 1):
                amount = month["concept_totals"].get(code, 0)
                total += decimal(amount)
                summary.write_number(row, col, amount, fmt["money"])
            summary.write_number(row, 13, float(total), fmt["total"])
            row += 1
    summary.freeze_panes(4, 1)
    summary.print_area(0, 0, row - 1, 13)
    summary.repeat_rows(0, 3)
    headers = ["Fecha contable", "Documento", "Diario", "Contraparte", "Descripción", "Referencia",
               "Entrada / débito", "Salida / crédito", "Saldo según libros", "Moneda", "Concepto", "Origen clasificación",
               "Contrapartidas", "Transacción bancaria", "Importe moneda empresa", "Moneda empresa", "Odoo"]
    detail.write_row(3, 0, headers, fmt["head"])
    detail.set_row(3, 32)
    detail.set_column(0, 0, 14)
    detail.set_column(1, 5, 30)
    detail.set_column(6, 9, 18)
    detail.set_column(10, 13, 32)
    detail.set_column(14, 16, 24)
    detail.write_string(4, 4, "Saldo inicial según libros", fmt["concept"])
    detail.write_number(4, 8, months[0]["opening"], fmt["money"])
    names = {item["code"]: item["name"] for item in snapshot["concepts"]}
    for row, item in enumerate(snapshot["accounting_movements"], 5):
        detail.write_datetime(row, 0, date.fromisoformat(item["date"]), fmt["date"])
        values = {1: item["document"], 2: item["journal"], 3: item["partner"], 4: item["description"],
                  5: item["reference"], 9: meta["currency"], 10: names.get(item["concept_code"], "Pendiente de clasificar"),
                  11: _origin(item["classification_origin"]), 12: item["counterpart_accounts"],
                  13: "Sí" if item["statement_line_id"] else "Sin transacción bancaria vinculada", 15: meta["company_currency"]}
        for col, value in values.items():
            detail.write_string(row, col, value or "", fmt["text"])
        for col, value in {6: max(item["amount"], 0), 7: max(-item["amount"], 0),
                           8: item["running_balance"], 14: item["company_amount"]}.items():
            detail.write_number(row, col, value, fmt["money"])
        _link(detail, row, 16, snapshot, "account.move.line", item["source_id"])
    detail.freeze_panes(5, 0)
    detail.repeat_rows(0, 3)
    detail.autofilter(3, 0, max(4, len(snapshot["accounting_movements"]) + 4), 16)


def _pending(sheet, snapshot, fmt):
    headers = ["Corte", "Fecha contable", "Partida", "Documento", "Contraparte", "Descripción", "Cuenta contable", "Medio de pago", "Pendiente moneda banco", "Moneda banco", "Pendiente moneda compañía", "Moneda compañía", "Odoo", "Diario de origen"]
    labels = {"deposit": "Depósito en tránsito", "check": "Cheque en circulación", "payment": "Otro pago pendiente", "suspense": "Movimiento en transitoria"}
    sheet.write_row(0, 0, headers, fmt["head"])
    sheet.set_row(0, 36)
    sheet.set_column(0, 1, 14)
    sheet.set_column(2, 4, 30)
    sheet.set_column(5, 6, 50)
    sheet.set_column(7, 11, 24)
    sheet.set_column(12, 12, 18)
    sheet.set_column(13, 13, 35)
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
        if item.get("journal"):
            sheet.write_string(row, 13, item["journal"], fmt["text"])
    sheet.autofilter(0, 0, max(len(snapshot["pending"]), 1), len(headers) - 1)
