"""Fail CI if Odoo exits successfully without actually running our tests."""

from pathlib import Path
import re
import sys


text = Path(sys.argv[1]).read_text()
started = re.findall(r"Starting TestQuadraticReconciliation\.(test_\w+)", text)
summary = re.findall(r"(\d+) failed, (\d+) error\(s\) of (\d+) tests", text)
if not started or not summary:
    raise SystemExit("No se encontró una ejecución de las pruebas del módulo.")
failed, errors, total = map(int, summary[-1])
print("Integración Odoo: %s pruebas iniciadas; %s total; %s fallos; %s errores" % (len(started), total, failed, errors))
if failed or errors or total < len(started) or "skipped" in text.lower():
    raise SystemExit("La ejecución de integración no está completa o contiene fallos.")
