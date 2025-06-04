# 🏦 Macro EliminarCampos – Limpieza y Exportación de Extractos Bancarios en Excel 💳

Esta macro está diseñada para automatizar el procesamiento de **extractos bancarios** en Excel, facilitando la limpieza de datos y la exportación a un nuevo archivo limpio, listo para análisis o auditoría.

---

## ⚙️ Funcionalidades clave

### 📂 Selección y carga del extracto
- Permite seleccionar un archivo Excel que contiene el extracto bancario.
- Copia los datos de la primera hoja para trabajar en el archivo actual.

### 🏢 Extracción de metadatos importantes
- Extrae el nombre de la empresa (`A4`) y la fecha del extracto (`A8`) para referencia.

### 🔎 Filtrado avanzado de información
- Aplica filtros para eliminar filas con encabezados, textos repetitivos o información irrelevante (como etiquetas de cliente, fechas, movimientos generales, etc.).
- Deja solo las filas con movimientos bancarios reales, excluyendo filas de información auxiliar o encabezados.

### 💾 Creación y guardado de archivo limpio
- Copia los datos filtrados a un nuevo libro Excel sin macros.
- Solicita al usuario el nombre y ubicación para guardar el archivo procesado en formato `.xlsx`.
- Cierra automáticamente el archivo original sin modificarlo.

---

## 🚀 Cómo usarla

1. Ejecuta la macro desde el libro Excel donde deseas importar y limpiar el extracto bancario.
2. Selecciona el archivo del extracto bancario cuando se abra el diálogo.
3. Después del procesamiento, elige dónde guardar el archivo limpio y listo para su análisis.

---

## 🎯 Beneficios

- **Optimiza el análisis bancario** al eliminar manualmente filas innecesarias.
- **Evita errores manuales** en el filtrado de datos bancarios.
- Facilita la **gestión documental** al generar archivos limpios y estándar.
- Compatible con extractos que incluyen encabezados y texto de información general.

---