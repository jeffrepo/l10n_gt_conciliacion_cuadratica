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
            origin = {"manual": "Manual", "rule": "Regla", "pending": "Pendiente"}.get(allocation["origin"], "")
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
