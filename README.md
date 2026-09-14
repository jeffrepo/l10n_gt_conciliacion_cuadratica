# Conciliación cuadrática para Odoo 19

Módulo independiente `l10n_gt_conciliacion_cuadratica`, versión `19.0.1.0.0`.
Genera un XLSX por compañía y cuenta bancaria, con resumen de enero al mes de
corte, movimientos clasificados y partidas conciliatorias. Depende únicamente
de `account` y de la biblioteca Python `xlsxwriter`; funciona sobre los modelos
contables comunes a Community y Enterprise. No depende de GDOMEX ni de `account_gt`.

## Instalación

1. Colocar este repositorio con el nombre `l10n_gt_conciliacion_cuadratica` dentro
   de una ruta de addons de Odoo 19.
2. Instalar `xlsxwriter` en el entorno Python del servidor si no está disponible.
3. Actualizar la lista de aplicaciones e instalar **Conciliación cuadrática Guatemala**.
4. Acceder con permisos de Contabilidad. La configuración de conceptos y reglas
   requiere permisos de administrador de Contabilidad.

Ejemplo de instalación por consola:

```bash
odoo-bin -d base_pruebas -i l10n_gt_conciliacion_cuadratica --stop-after-init
```

## Configuración inicial

- En cada diario bancario, pestaña **Conciliación cuadrática**, activar su
  inclusión. Revisar la cuenta bancaria vinculada, moneda y tipo de cuenta.
- Mantener configuradas las cuentas de cobros y pagos pendientes de los métodos
  de pago. El módulo las obtiene automáticamente. Las cuentas pendientes
  adicionales se configuran explícitamente en el diario.
- En **Contabilidad → Conciliación cuadrática → Cortes de extractos**, indicar
  la fecha final que cubre cada documento bancario cuando sea diferente de la
  última transacción. Si queda vacía, se usa la fecha del extracto de Odoo.
  Para conservar un cierre se requiere un extracto de control al último día de
  cada mes incluido, con saldo final real y líneas completas.
- Configurar las reglas usando cuentas de contrapartida, socios/relacionadas
  identificados por contacto, país, método de pago o texto. No se incluyen reglas
  universales que adivinen el plan contable de la empresa. La relación entre un
  contacto y una empresa se expresa en las reglas de esa compañía.

Las condiciones de una regla se combinan con AND; una lista de cuentas o
contrapartes admite cualquiera de sus miembros. Gana la prioridad numérica
menor. Conceptos distintos empatados a la misma prioridad quedan pendientes.
Un país desconocido no se convierte en país extranjero. Un concepto manual
o una distribución por conceptos prevalecen sobre las reglas.

El catálogo inicial distingue ingresos, egresos y su finalidad. Incluye
**Cheques emitidos** para permitir la presentación del formato de referencia.
Configurar su prioridad frente a **Proveedores** y **Gastos operativos** según
el criterio acordado; cada importe se incluye una sola vez. El medio de pago
permanece disponible en el detalle aunque no sea el concepto principal.

## Operación mensual

1. Cargar los movimientos y estados de cuenta en el flujo bancario habitual de
   Odoo y realizar la conciliación. No se suben archivos al asistente de este módulo.
2. Abrir **Generar conciliación**, seleccionar empresa, año, mes de corte y una
   o varias cuentas habilitadas (o todas).
3. Pulsar **Calcular** y revisar los resultados. Los faltantes, diferencias y
   distribuciones incompletas se muestran como pendientes.
4. Corregir la clasificación en **Clasificar movimientos**. Un movimiento puede
   tener un concepto único o una distribución de importes positivos. El resto
   no distribuido queda visible como pendiente. No se modifican importes contables.
5. Generar una nueva versión y descargar los XLSX. Varias cuentas se entregan
   en un ZIP. Cada archivo contiene **Conciliación**, **Movimientos** y
   **Partidas conciliatorias**, con enlaces a sus registros originales en Odoo.
6. Opcionalmente, **Conservar cierre** cuando no queden pendientes de revisión.
   Esto conserva la versión; no publica asientos, no concilia movimientos y no
   modifica las fechas de bloqueo contable de Odoo.

Las versiones son inmutables y no se eliminan desde el módulo. Corregir una
fuente no cambia un resultado anterior. Una nueva versión usa las fuentes
vigentes; no reconstruye registros eliminados ni el estado de conocimiento
anterior a una modificación retroactiva. El XLSX se exporta desde el resultado
conservado, no consultando otra vez los movimientos.

## Criterio de cálculo

Los importes se expresan en la moneda del diario (o de la compañía cuando el
diario no define otra). Los saldos de bancos y los de libros se obtienen de
fuentes separadas:

- **Banco calculado:** saldo inicial respaldado por un extracto más entradas
  menos salidas. No se reinicia con cada extracto para esconder saltos.
- **Control bancario:** comparación con `balance_end_real` del extracto que
  cubre el fin de mes. Los saldos reales deben haber sido cargados desde el banco;
  un saldo autocompletado por Odoo no prueba por sí solo una verificación externa.
- **Saldo según libros:** mayor de la cuenta bancaria + cuentas pendientes
  atribuibles al banco + cuenta transitoria atribuible al banco. Esta es la
  definición de libro de bancos utilizada para el flujo con cuentas pendientes;
  los tres componentes se muestran por separado.
- **Banco ajustado:** banco calculado + depósitos en tránsito − cheques en
  circulación − otros pagos pendientes.
- **Libros ajustados:** saldo según libros menos el residual firmado de la
  transitoria. Cada ajuste tiene su apunte en Partidas conciliatorias; no se
  inventa una contrapartida para hacer coincidir los saldos.
- **Diferencia:** banco ajustado menos libros ajustados.

Los residuales se reconstruyen con los importes conciliados cuya `max_date`
(máxima fecha contable de los dos apuntes) no supera el corte. Así, un cheque
de enero pagado parcialmente en febrero conserva todo el pendiente en enero
y solo el remanente en febrero. Los pagos sin método `check_printing` se
presentan como otros pagos pendientes, no se etiquetan arbitrariamente como cheques.

En moneda extranjera se utilizan `amount_currency` y los importes en moneda
extranjera de las conciliaciones parciales. Nunca se convierten saldos antiguos
con el tipo de cambio actual. Las partidas conservan también su saldo residual
en moneda de la compañía. Una diferencia cambiaria con importe cero en moneda
bancaria no se trata como movimiento de caja en esa moneda.

Los flujos anuales se suman. Los saldos muestran el último mes incluido; el
saldo inicial anual es el de enero. Los meses posteriores al corte quedan vacíos.
Los importes faltantes se muestran como `n.d.` en el XLSX, no como un cero validado.

## Alcance y límites de esta versión

- Se procesan transacciones bancarias y apuntes **publicados**. Los reversos
  publicados permanecen visibles; los borradores no alimentan el reporte.
- El flujo de cuentas pendientes separadas está soportado. Los métodos de pago
  que usan directamente la cuenta bancaria generan un pendiente de configuración:
  no se presenta su conciliación como un cierre validado en esta versión.
- Las cuentas transitorias compartidas se atribuyen mediante el banco del
  movimiento, pago, diario o relación de conciliación. Los apuntes sin atribución
  inequívoca quedan señalados y no se asignan silenciosamente a una empresa/banco.
- Un diario en USD necesita que sus apuntes de banco y pendientes conserven USD.
  Los apuntes que solo conservan otra moneda generan un pendiente de revisión.
- Un mes sin extracto de control al cierre se puede exportar como borrador.
  No se presume cobertura completa simplemente porque no haya movimientos.
- Los números y bancos de contrapartes se muestran cuando están identificados
  en las fuentes; no se elige arbitrariamente una de sus cuentas bancarias.
- El desglose de dividendos presenta los diez mayores socios y agrupa el resto.
- El formato se basa en la estructura de conciliación proporcionada para el
  proyecto. No incluye presentación automática ante SAT ni certificación normativa.

## Pruebas

Las pruebas usan datos sintéticos: este repositorio no contiene archivos
bancarios, NIT, cuentas ni nombres reales de los ejemplos proporcionados.

Pruebas independientes del servidor (requieren `xlsxwriter` y `openpyxl`):

```bash
python -m unittest discover -s tests_unit -v
```

Pruebas de integración, en una base de datos desechable de Odoo 19:

```bash
odoo-bin -d cq_test -i l10n_us,l10n_gt_conciliacion_cuadratica \
  --test-enable --test-tags=/l10n_gt_conciliacion_cuadratica --stop-after-init
```

`l10n_us` se usa solo para los fixtures contables de las pruebas comunes de Odoo;
no es una dependencia del módulo. GitHub Actions instala el módulo en Odoo 19
con PostgreSQL 16 y ejecuta las pruebas de integración y de exportación.
