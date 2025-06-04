## 📌 Macro: Conciliación Contable Automática con Token DIAN

Esta macro automatiza completamente el proceso de conciliación contable mensual a partir de archivos de Excel exportados desde software contable y del token de la DIAN.

### ✅ Funcionalidades:

- 🗂️ Crea automáticamente un nuevo libro Excel donde organiza los datos conciliados.
- 📥 Solicita al usuario seleccionar:
  - El archivo del software contable (compras/ventas)
  - El archivo del token de la DIAN (emitido/recibido)
- 🔍 Extrae y organiza:
  - Rangos de compras y ventas desde el archivo contable
  - Facturas electrónicas recibidas y emitidas desde el token DIAN
  - Notas crédito recibidas y emitidas
- 🧮 Aplica:
  - Filtros personalizados
  - Eliminación de duplicados
  - Fórmulas `SUMIF` para totalizar por tercero
  - Sumatorias generales para revisión rápida

### 🧠 ¿Para qué sirve?

Ideal para contadores o auxiliares que quieren:
- Agilizar la conciliación mensual entre su sistema contable y la DIAN.
- Reducir errores al filtrar y totalizar manualmente.
- Generar un informe claro y exportable con todos los cálculos necesarios.

### 🧰 Salidas automáticas:

- Hoja `compras`
- Hoja `ventas`
- Hoja `compras_token`
- Hoja `ventas_token`
- Hoja `notas_credito_compras`
- Hoja `notas_credito_ventas`

Cada hoja está organizada, sin duplicados, y con totales automáticos.