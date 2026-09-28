"""Presentation of the supplied monthly form, without its sample bank data.

Each period has a detail column followed by a subtotal column. Detail rows may
grow, so formulas use the row map rather than the sample workbook's addresses.
"""

from collections import defaultdict

from xlsxwriter.utility import xl_rowcol_to_cell

from .calculation import MONTHS, decimal


def reference_formats(book):
    common = {"font_name": "Arial", "font_size": 9, "valign": "vcenter"}
    amount = '#,##0.00;-#,##0.00;"-"'
    styles = {
        "text": {"num_format": "@"},
        "label": {"text_wrap": True},
        "title": {"bold": True, "align": "center", "font_size": 11},
        "head": {"bg_color": "#203864", "font_color": "#FFFFFF", "align": "center", "text_wrap": True, "border": 1},
        "black": {"bg_color": "#000000", "font_color": "#FFFFFF", "align": "center"},
        "data_head": {"bg_color": "#C00000", "font_color": "#FFFFFF", "bold": True, "text_wrap": True, "border": 1},
        "concept": {"bold": True, "text_wrap": True, "bottom": 3, "left": 1, "right": 1},
        "detail": {"text_wrap": True, "bottom": 3, "left": 3, "right": 3},
        "section": {"bold": True, "bg_color": "#D9D9D9", "border": 1},
        "total_label": {"bold": True, "bg_color": "#BFBFBF", "border": 1, "text_wrap": True},
        "total": {"bg_color": "#BFBFBF", "border": 1, "num_format": amount},
        "money": {"num_format": amount, "bottom": 3, "left": 3, "right": 1},
        "empty": {"right": 1},
        "date": {"num_format": "dd/mm/yyyy"},
        "warn": {"bg_color": "#FFF2CC", "text_wrap": True},
        "error": {"bg_color": "#FCE4D6"},
    }
    return {name: book.add_format({**common, **values}) for name, values in styles.items()}


def _groups(snapshot, code, bank_only=False, profile=None):
    groups = defaultdict(lambda: defaultdict(lambda: decimal(0)))
    for movement in snapshot["movements"]:
        for allocation in movement["allocations"]:
            if allocation["code"] != code:
                continue
            bank = movement["counterparty_bank"] or "Banco no identificado"
            account = movement["counterparty_account"] or "Cuenta no identificada"
            key = (bank, account) if bank_only else (movement["partner"] or "Contraparte no identificada", bank, account)
            groups[key][int(movement["date"][5:7])] += decimal(allocation["amount"])
    result = sorted(groups.items(), key=lambda item: (-sum(item[1].values()), item[0]))
    if code == "OUT_DIVIDENDS" and len(result) > 10:
        other = defaultdict(lambda: decimal(0))
        for _, amounts in result[10:]:
            for number, value in amounts.items():
                other[number] += value
        result = result[:10] + [(("Otros socios", "", ""), other)]
    if profile:
        from .template import merge_manual_details
        return merge_manual_details(result, profile["details"].get(code, []), bank_only)
    return [(key, values, None) for key, values in result]


def summary(sheet, snapshot, confirmed, fmt, profile=None):
    meta, months = snapshot["metadata"], snapshot["months"]
    currency = {"GTQ": "Quetzales", "USD": "Dólares"}.get(meta["currency"], meta["currency"])
    symbol = {"GTQ": "Q.", "USD": "$"}.get(meta["currency"], meta["currency"])
    sheet.set_default_row(15)
    sheet.set_column("A:A", 2.86)
    sheet.set_column("B:B", 6.14)
    sheet.set_column("C:C", 25.43)
    sheet.set_column("D:E", 20.14)
    sheet.set_column("F:AE", 12.14)
    sheet.set_row(3, 25.5)
    sheet.set_row(12, 8.25)
    sheet.set_row(15, 3.75)
    sheet.set_row(16, 22.5)

    def merge(first_row, first_col, last_col, value, style="detail"):
        sheet.merge_range(first_row - 1, first_col - 1, first_row - 1, last_col - 1, value, fmt[style])

    merge(4, 2, 3, "Nombre o Razón Social:", "concept")
    merge(4, 4, 5, meta["company"])
    sheet.write_string("F4", "NIT:", fmt["text"])
    sheet.write_string("G4", meta["vat"], fmt["text"])
    merge(5, 2, 7, "Identificación del Banco donde tiene su cuenta", "black")
    for row, label, value in (
        (6, "Nombre del Banco:", meta["bank"] or (profile or {}).get("bank") or "Banco no identificado"),
        (7, "País:", meta["country"]),
        (9, "Tipo de Cuenta:", meta["account_type"]),
        (11, "Nombre de la Cuenta Contable:", meta.get("ledger_account_name", meta["ledger_account"])),
        (12, "No. de Cuenta Contable:", meta.get("ledger_account_code", "")),
    ):
        merge(row, 2, 3, label, "concept")
        merge(row, 4, 7, value)
    merge(8, 2, 3, "No. de Cuenta Bancaria:", "concept")
    sheet.write_string("D8", meta["account_number"] or (profile or {}).get("account") or "Cuenta no identificada", fmt["text"])
    sheet.write_string("E8", "Tipo de Moneda:", fmt["text"])
    merge(8, 6, 7, currency)
    merge(10, 2, 7, "Datos de la Contabilidad", "black")
    merge(14, 2, 31, "CONCILIACIÓN CUADRÁTICA MENSUAL (POR CUENTA BANCARIA)", "title")
    merge(15, 2, 31, "(Valores en %s)" % currency, "title")
    merge(17, 2, 5, "CONCILIACIÓN DE SALDOS BANCARIOS", "head")
    merge(17, 6, 31, "PERÍODOS %s" % snapshot["year"], "head")
    sheet.write_blank("B18", None, fmt["head"])
    merge(18, 3, 5, "CONCEPTOS", "head")
    for idx, label in enumerate((*MONTHS, "TOTALES")):
        merge(18, 6 + idx * 2, 7 + idx * 2, label.upper(), "head")
    sheet.freeze_panes(18, 5)
    sheet.repeat_rows(16, 17)
    sheet.repeat_columns(1, 4)

    # Row descriptors separate layout from formulas so added counterparties do
    # not overwrite subsequent concepts or leave totals at stale addresses.
    rows, index, used = [], {}, set()
    concepts = {item["code"]: item for item in snapshot["concepts"]}

    def add(key, label, values=None, *, code="", style="concept", annual="sum", left=False, formula=None, height=15):
        row = len(rows) + 19
        index[key] = row
        rows.append(dict(key=key, row=row, label=label, values=values, code=code,
                         style=style, annual=annual, left=left, formula=formula, height=height,
                         template_row=(profile or {}).get("row_map", {}).get(key)))
        return row

    def amounts(code):
        return [month["concept_totals"].get(code, 0) for month in months]

    def concept(code, label, reserve=0, bank_only=False, person_label="Nombre", height=15):
        used.add(code)
        add(code, label, amounts(code), code=concepts.get(code, {}).get("report_code", ""), height=height)
        groups = _groups(snapshot, code, bank_only, profile) if reserve else []
        if reserve:
            if profile:
                reserve = len(profile["details"].get(code, []))
            origin_row = (profile or {}).get("row_map", {}).get(code)
            rows.append(dict(row=len(rows) + 19, key="", header=True, bank_only=bank_only, person_label=person_label, template_row=origin_row + 1 if origin_row else None))
            children = []
            for idx in range(max(reserve, len(groups))):
                key, values, source_row = groups[idx] if idx < len(groups) else (("", "") if bank_only else ("", "", ""), {}, origin_row + 2 if origin_row else None)
                child = "%s_detail_%s" % (code, idx)
                add(child, key, [float(values.get(month["number"], 0)) for month in months], style="detail", left=True)
                rows[-1]["template_row"] = source_row
                if any(len(str(value)) > 28 for value in key):
                    rows[-1]["height"] = 30
                children.append(child)
            rows[index[code] - 19]["formula"] = lambda col, keys=children: "=SUM(%s)" % ",".join(cell(key, col - 1) for key in keys)

    def metric(key, label, *, annual="last", code="", style="concept", formula=None):
        add(key, label, [month[key] for month in months], code=code, annual=annual, style=style, formula=formula)

    def other(direction, label, fixed=()):
        codes = [code for code, item in concepts.items() if item["direction"] == direction and code not in used and code not in fixed]
        pending = "unclassified_" + direction
        codes.append(pending)
        codes = [code for code in codes if any(amounts(code)) or code in (profile or {}).get("row_map", {})]
        children = list(fixed) + codes
        key = "other_" + direction
        values = [float(sum((decimal(month["concept_totals"].get(code, 0)) for code in children), decimal(0))) for month in months]
        add(key, label, values, formula=(lambda col: "=SUM(%s)" % ",".join(cell(code, col - 1) for code in children)) if children else None)
        for code in children:
            used.add(code)
            fallback_names = {"OUT_TAXES": "Pago de impuestos", "OUT_FEES": "Comisiones bancarias y ACH", "OUT_CHECKS": "Cheques emitidos", "OUT_INTEREST": "Intereses pagados", "OUT_LOANS_GRANTED": "Préstamos otorgados a socios y terceros"}
            name = concepts.get(code, {}).get("name", fallback_names.get(code, "Pendiente de clasificar"))
            add(code, name, amounts(code), code=concepts.get(code, {}).get("report_code", ""), style="detail", left=True)
        return key

    add("bank_opening", "SALDO INICIAL SEGÚN BANCO", [month.get("reported_bank_opening", month["bank_opening"]) for month in months], annual="first")
    add("in_heading", "DEPÓSITOS", [month["income"] for month in months], code="( + )", style="section",
        formula=lambda col: "=SUM(%s)" % ",".join(cell(key, col) for key in income_keys))
    concept("IN_CUSTOMERS_LOCAL", "Cuentas por cobrar clientes locales")
    concept("IN_CUSTOMERS_FOREIGN", "Cuentas por cobrar clientes del exterior")
    concept("IN_RELATED", "Cuentas y documentos por cobrar relacionadas locales y del exterior (detallar)", 2, person_label="Empresa")
    concept("IN_SHAREHOLDERS", "Cuentas por cobrar a socios", 2, person_label="Nombre del Socio")
    concept("IN_EMPLOYEES", "Cuentas por cobrar empleados")
    concept("IN_ADVANCES", "Anticipo de clientes")
    concept("IN_INTEREST", "Intereses ganados")
    concept("IN_TRANSFER", "Transferencia de fondos entre cuentas bancarias (detallar)", 4, bank_only=True)
    # These two rows follow 'Otros ingresos' in the reference, not within it.
    used.update(("IN_SUPPLIER_REFUNDS", "IN_LOAN_REPAYMENTS"))
    other_in = other("in", "Otros ingresos (especifique)")
    concept("IN_SUPPLIER_REFUNDS", "Reintegro por devoluciones a proveedores")
    concept("IN_LOAN_REPAYMENTS", "Devoluciones por préstamos")
    income_keys = ["IN_CUSTOMERS_LOCAL", "IN_CUSTOMERS_FOREIGN", "IN_RELATED", "IN_SHAREHOLDERS", "IN_EMPLOYEES", "IN_ADVANCES", "IN_INTEREST", "IN_TRANSFER", other_in, "IN_SUPPLIER_REFUNDS", "IN_LOAN_REPAYMENTS"]
    add("available", "TOTAL DEPÓSITOS DEL PERÍODO", [None if month["bank_opening"] is None else float(decimal(month["bank_opening"]) + decimal(month["income"])) for month in months], code="A", style="total_label", annual="available",
        formula=lambda col: "=%s+SUM(%s)" % (cell("bank_opening", col), ",".join(cell(key, col) for key in income_keys)))
    add("out_heading", "EGRESOS", code="( - )", style="section")
    concept("OUT_SUPPLIERS", "Proveedores")
    concept("OUT_OPERATING", "Gastos Operativos")
    concept("OUT_ADVANCES", "Pagos anticipados")
    concept("OUT_LOANS", "Pago a Préstamos (detallar)", 3)
    concept("OUT_DIVIDENDS", "Pago de dividendos (Detalle los 10 más importantes, los demás agrúpelos en una sola línea)", 2, person_label="Nombre del Socio", height=28.5)
    concept("OUT_SHAREHOLDERS", "Cuentas por pagar socios (detallar)", 2, person_label="Nombre del Socio")
    concept("OUT_RELATED_LOCAL", "Cuentas por pagar relacionadas locales (detallar)", 2, person_label="Empresa")
    concept("OUT_RELATED_FOREIGN", "Cuentas por pagar relacionadas del exterior (detallar)", 2, person_label="Empresa")
    concept("OUT_TRANSFER", "Egreso por transferencia de fondos entre cuentas bancarias (detallar)", 4, bank_only=True, height=30)
    other_out = other("out", "Otros egresos (especifique):", ("OUT_TAXES", "OUT_FEES", "OUT_CHECKS", "OUT_INTEREST", "OUT_LOANS_GRANTED"))
    expense_keys = ["OUT_SUPPLIERS", "OUT_OPERATING", "OUT_ADVANCES", "OUT_LOANS", "OUT_DIVIDENDS", "OUT_SHAREHOLDERS", "OUT_RELATED_LOCAL", "OUT_RELATED_FOREIGN", "OUT_TRANSFER", other_out]
    metric("expense", "TOTAL EGRESOS", annual="sum", code="B", style="total_label", formula=lambda col: "=SUM(%s)" % ",".join(cell(key, col) for key in expense_keys))
    metric("bank_end", "SALDO FINAL BANCARIO ( A - B )", style="total_label", formula=lambda col: "=%s-%s" % (cell("available", col), cell("expense", col)))
    has_manual_balances = any(month.get("manual_balance") for month in months)
    if has_manual_balances:
        add(
            "reported_closing", "Saldo final según banco (extracto o captura)", [month["statement_end"] for month in months], annual="last")
    add("bank_adjustments", "AJUSTES")
    metric("deposits", "Depósitos en tránsito", code="( + )")
    metric("checks", "Cheques en circulación", code="( - )")
    add("other_bank", "Otros (especifique)", [-month["payments"] for month in months], annual="last")
    add("bank_note", "Otros pagos pendientes de cargo", style="detail")
    def adjusted_formula(col):
        base = cell("bank_end", col)
        if has_manual_balances:
            base = "IF(ISNUMBER(%s),%s,%s)" % (base, base, cell("reported_closing", col))
        return "=%s+%s-%s+%s" % (base, cell("deposits", col), cell("checks", col), cell("other_bank", col))

    metric("bank_adjusted", "SALDO CONCILIADO", style="total_label", formula=adjusted_formula)
    add("spacer", "", style="label")
    metric("book_balance", "SALDO FINAL SEGÚN LIBROS")
    add("book_adjustments", "AJUSTES")
    # Only identifiable interest still in suspense belongs in the book
    # adjustment. Already recorded interest must not be added a second time.
    interest_moves = {move["move_id"] for move in snapshot["movements"] if move["amount"] > 0 and move["allocations"] and all(item["code"] == "IN_INTEREST" for item in move["allocations"])}
    interest, debits, receipts = [], [], []
    for month in months:
        suspense = [item for item in snapshot["pending"] if item["month"] == month["number"] and item["kind"] == "suspense"]
        interest.append(float(-sum((decimal(item["amount"]) for item in suspense if item["amount"] < 0 and item["move_id"] in interest_moves), decimal(0))))
        debits.append(float(sum((decimal(item["amount"]) for item in suspense if item["amount"] > 0), decimal(0))))
        receipts.append(float(decimal(month["book_adjustment"]) - decimal(interest[-1]) + decimal(debits[-1])))
    add("interest_pending", "Ingresos por intereses ganados", interest, code="( + )", annual="last")
    add("debits_pending", "Nota de débito no operadas", debits, code="( - )", annual="last")
    add("other_book", "Otros (especifique)", receipts, annual="last", formula=lambda col: "=" + cell("receipts_pending", col - 1))
    add("receipts_pending", "Otros ingresos bancarios pendientes de registrar", receipts, annual="last", style="detail", left=True, height=27)
    add("book_spacer", "", style="detail")
    metric("book_adjusted", "SALDO CONCILIADO", style="total_label", formula=lambda col: "=%s+%s-%s+%s" % tuple(cell(key, col) for key in ("book_balance", "interest_pending", "debits_pending", "other_book")))

    def cell(key, col):
        return xl_rowcol_to_cell(index[key] - 1, col)

    def put(row, col, value, style, formula=None):
        if value is None:
            sheet.write_string(row, col, "n.d.", fmt[style])
        elif formula:
            sheet.write_formula(row, col, formula, fmt[style], value)
        else:
            sheet.write_number(row, col, value, fmt[style])

    bank_flows_missing = not snapshot["movements"] and all(month["bank_end"] is None for month in months)
    for item in rows:
        row = item["row"] - 1
        if item.get("header"):
            sheet.write_string(row, 2, "Banco" if item["bank_only"] else item["person_label"], fmt["head"])
            if item["bank_only"]:
                sheet.merge_range(row, 3, row, 4, "No. de Cuenta", fmt["head"])
            else:
                sheet.write_row(row, 3, ["Banco", "No. Cuenta"], fmt["head"])
            for col in range(5, 30, 2):
                sheet.write_string(row, col, "Valor " + symbol, fmt["black"])
            continue
        sheet.set_row(row, item["height"])
        sheet.write_string(row, 1, item["code"], fmt["text"])
        if isinstance(item["label"], tuple):
            if len(item["label"]) == 2:
                sheet.write_string(row, 2, item["label"][0], fmt["detail"])
                sheet.merge_range(row, 3, row, 4, item["label"][1], fmt["detail"])
            else:
                sheet.write_row(row, 2, item["label"], fmt["detail"])
        else:
            merge(item["row"], 3, 5, item["label"], item["style"])
        style = "total" if item["style"] == "total_label" else "money"
        for number in range(13):
            col = 5 + number * 2
            sheet.write_blank(row, col, None, fmt["total"] if style == "total" else fmt["empty"])
            sheet.write_blank(row, col + 1, None, fmt[style])
        if item["values"] is None:
            continue
        offset = 5 if item["left"] else 6
        if bank_flows_missing and index["in_heading"] <= item["row"] < index["bank_end"]:
            # Missing statements do not establish zero deposits/withdrawals.
            for idx in range(len(months)):
                put(row, offset + idx * 2, None, style)
            put(row, 29 if item["left"] else 30, None, style)
            continue
        for idx, value in enumerate(item["values"]):
            col = offset + idx * 2
            formula = item["formula"](col) if item["formula"] else None
            if item["key"] == "bank_opening" and idx and not months[idx].get("manual_opening_used"):
                formula = "=" + cell("bank_end", col - 2)
            put(row, col, value, style, formula)
        annual_col = 29 if item["left"] else 30
        if item["annual"] == "sum":
            value = float(sum((decimal(value) for value in item["values"]), decimal(0)))
            formula = "=SUM(%s)" % ",".join(xl_rowcol_to_cell(row, offset + idx * 2) for idx in range(len(months)))
        elif item["annual"] == "available":
            value = None if months[0]["bank_opening"] is None else float(decimal(months[0]["bank_opening"]) + sum((decimal(month["income"]) for month in months), decimal(0)))
            formula = item["formula"](annual_col)
        else:
            idx = 0 if item["annual"] == "first" else len(months) - 1
            value = item["values"][idx]
            formula = "=" + xl_rowcol_to_cell(row, offset + idx * 2)
        # Subtotals in the total column reference their own detail totals.
        if item["formula"] and item["annual"] == "sum" and item["key"] != "expense":
            formula = item["formula"](annual_col)
        put(row, annual_col, value, style, formula)

    # Controls and provenance sit below the form, leaving its visual hierarchy
    # intact. Books remain an independent source, never a copy of bank totals.
    row = len(rows) + 22
    merge(row, 2, 31, "CONTROL Y TRAZABILIDAD ODOO", "head")
    row += 1
    state = "CIERRE CONSERVADO" if confirmed else "BORRADOR"
    notes = [
        "%s · %s pendientes de revisión · Corte: %s" % (state, len(snapshot["issues"]), months[-1]["cutoff"]),
        "Diarios incluidos: %s · Control de extractos: %s" % (meta["journal"], meta.get("control_journal") or "Pendiente de identificar"),
        "A incluye saldo inicial + depósitos. TOTALES: apertura de enero + ingresos acumulados; saldos al último mes incluido. n.d.: falta respaldo bancario.",
        "Generado: %s · %s" % (meta["generated_at"], meta["generated_by"]),
    ]
    if "accounting_movements" in snapshot:
        notes.append("Fuentes del período: %s transacciones bancarias · %s apuntes de la cuenta bancaria. Consulte Mayor bancario Odoo y Resumen contable Odoo; sus importes no sustituyen los extractos." % (len(snapshot["movements"]), len(snapshot["accounting_movements"])))
    for note in notes:
        merge(row, 2, 31, note, "warn" if "pendientes de revisión" in note and snapshot["issues"] else "label")
        sheet.set_row(row - 1, 24)
        row += 1
    for key, label in (("income", "Depósitos del período (sin saldo inicial)"), ("ledger_bank", "Mayor de la cuenta bancaria"), ("ledger_outstanding", "Mayor de cuentas pendientes"), ("ledger_suspense", "Mayor de cuenta transitoria"), ("statement_end", "Saldo final según banco (extracto o captura)"), ("bank_difference", "Diferencia banco contra saldo informado"), ("difference", "Diferencia banco ajustado menos libros ajustados")):
        merge(row, 3, 5, label, "concept")
        sheet.set_row(row - 1, 27)
        for idx, month in enumerate(months):
            put(row - 1, 6 + idx * 2, None if key == "income" and bank_flows_missing else month[key], "money")
        value = float(sum((decimal(month[key]) for month in months), decimal(0))) if key == "income" else months[-1][key]
        put(row - 1, 30, None if key == "income" and bank_flows_missing else value, "money")
        if key in ("bank_difference", "difference"):
            sheet.conditional_format(row - 1, 5, row - 1, 30, {"type": "cell", "criteria": "not between", "minimum": -float(snapshot["rounding"]) / 2, "maximum": float(snapshot["rounding"]) / 2, "format": fmt["error"]})
        row += 1
    for month in months:
        merge(row, 3, 31, "%s: %s · Corte %s" % (month["name"], month["control"]["name"] if month["control"] else "Sin extracto al cierre", month["cutoff"]), "label")
        row += 1
        if month.get("manual_balance"):
            capture = month["manual_balance"]
            merge(row, 3, 31, "%s: captura manual %s · Apertura %s · Cierre %s · Actualizada %s" % (
                month["name"], capture["name"], capture["opening"], capture["closing"], capture.get("write_date", "")), "label")
            row += 1
    for issue in snapshot["issues"]:
        merge(row, 3, 31, issue["message"], "warn")
        sheet.set_row(row - 1, 24)
        row += 1
    sheet.print_area(3, 1, row - 1, 30)
    return {"rows": rows, "index": index, "form_end": index["book_adjusted"]}
