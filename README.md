# Analítica de Facturación en Salud

**Facturación, radicación, glosas, recaudo y cartera**

Aplicación web en Streamlit para cargar, validar y analizar archivos de facturación de prestadores de servicios de salud. Trabaja solo con el archivo que carga el usuario: lo procesa en memoria, no lo guarda en el servidor y no envía datos a servicios externos.

> Los resultados dependen de la calidad y estructura del archivo cargado. La herramienta facilita el análisis exploratorio y no reemplaza una auditoría financiera, contable, médica o contractual.

---

## Contenido del repositorio

| Archivo | Descripción |
|---|---|
| `app.py` | Aplicación completa (un solo archivo, 17 secciones). |
| `requirements.txt` | Dependencias externas para Streamlit Community Cloud. |
| `README.md` | Este documento. |

No se requiere incluir datos en el repositorio. Cada sesión nueva solicita el archivo.

---

## Archivos admitidos

- **Excel `.xlsx`**: detecta las hojas y selecciona automáticamente `Facturacion` si existe; el usuario puede elegir otra.
- **Excel `.xls`**: soportado mediante `xlrd`.
- **CSV `.csv`**: detecta codificación (UTF-8, Windows-1252, Latin-1) y separador (coma, punto y coma, tabulación, barra vertical). Si la detección falla, el usuario puede elegirlos manualmente.

Límite de carga por defecto de Streamlit: 200 MB.

---

## Estructura esperada del archivo

| Grupo | Columnas | Efecto si faltan |
|---|---|---|
| **Mínimas** | `ID_Factura`, `Fecha_Radicacion` o `Periodo_Cierre`, `Valor_Neto` | Se bloquean los análisis financieros y se indica cómo corregir. |
| **Recomendadas** | `Entidad_Pagadora`, `Estado_Factura`, `Valor_Pagado`, `Saldo_Cartera`, `Glosa_Inicial` | La app continúa; se desactivan los análisis que dependen de ellas. |
| **Opcionales** | `Numero_Radicado`, `Fecha_Servicio`, `Fecha_Emision`, `Fecha_Aceptacion`, `Fecha_Pago`, `Codigo_Entidad`, `Regimen`, `Sede`, `Codigo_Servicio`, `Servicio`, `Linea_Servicio`, `Cantidad`, `Canal_Radicacion`, `Dias_Emision_Radicacion`, `Soporte_Completo`, `Valor_Bruto`, `Descuento`, `Copago_Cuota` | Se omiten los análisis correspondientes con un mensaje explicativo. |

### Tolerancia en los encabezados

Los nombres se normalizan (sin tildes, sin distinguir mayúsculas, espacios y guiones como `_`) y se reconocen sinónimos frecuentes. Ejemplos: `ID Factura` → `ID_Factura`, `Fecha Radicación` → `Fecha_Radicacion`, `Línea de Servicio` → `Linea_Servicio`, `Copago` → `Copago_Cuota`.

La sección **Configurar correspondencia de columnas** permite asignar manualmente cualquier campo o marcarlo como *Columna no disponible*. El mapeo se conserva solo durante la sesión.

### Formatos reconocidos

- **Valores monetarios**: números, símbolos de moneda, separadores de miles y coma o punto decimal (`$ 1.234.567`, `1,234,567.50`, `(500)`). Se detecta la convención predominante de cada columna y se advierte cuando hay ambigüedad.
- **Fechas**: ISO (`2025-03-15`), día/mes/año (`15/03/2025`) y números de serie de Excel. Las fechas no interpretables quedan nulas y se informa cuántas.
- **Periodo de cierre**: se normaliza a `AAAA-MM` (`2025-03`, `202503`, `03/2025`, `mar-2025`). Si no existe, se construye desde `Fecha_Radicacion`.
- **Soporte completo**: `Sí/No`, `S/N`, `1/0`, `Completo/Incompleto`, `True/False`.

---

## Secciones de la aplicación

| # | Sección | Contenido principal |
|---|---|---|
| 1 | Inicio y carga de datos | Carga, selección de hoja o separador, equivalencias de columnas, mapeo manual y vista previa (primeros registros o muestra aleatoria). |
| 2 | Calidad de datos | Puntaje orientativo 0-100 (completitud, unicidad, validez, consistencia), duplicados, nulos, conversiones, inconsistencias y limpieza aplicada. |
| 3 | Resumen ejecutivo | 11 tarjetas KPI con comparación del último periodo frente al anterior y 7 gráficos. |
| 4 | Cierre mensual | Tabla mensual, trimestral o anual con crecimiento, acumulado anual, promedio móvil y periodos destacados. |
| 5 | Entidades pagadoras | Tabla comparativa, rankings, dispersión, mapa de calor, detalle por entidad y análisis de Pareto (HHI). |
| 6 | Servicios y líneas | Top 5/10/15/todos, treemap, burbujas, boxplot, mapa servicio-entidad y evolución mensual. |
| 7 | Sedes y regímenes | Barras agrupadas, barras al 100 %, boxplot y mapa de calor por periodo. |
| 8 | Estados de facturación | Estados detectados dinámicamente, apiladas por periodo y tablas cruzadas. |
| 9 | Glosas | Frecuencia (facturas glosadas / total) y severidad (valor glosado / valor neto), rankings y soporte completo vs incompleto. |
| 10 | Recaudo y cartera | Evolución, ranking de cartera, entidades con alto volumen y bajo recaudo, antigüedad aproximada con fecha de corte. |
| 11 | Tiempos de radicación | Estadísticos (P75, P90, P95), análisis por dimensión, rangos configurables y casos extremos. |
| 12 | Análisis estadístico | Descriptiva con interpretación automática, correlaciones (Pearson/Spearman) y pruebas: t de Welch, Mann-Whitney U, ANOVA, Kruskal-Wallis, chi-cuadrado y Shapiro-Wilk. |
| 13 | Valores atípicos | IQR (multiplicador configurable), puntaje Z o percentiles; global o por servicio, línea o entidad. |
| 14 | Insights y alertas | Semáforos con umbrales editables e insights generados por reglas programadas (sin IA externa). |
| 15 | Explorador de datos | Búsqueda de texto, selección de columnas, orden y filtros rápidos (saldo, glosa, soporte, inconsistencias, atípicos). |
| 16 | Descargas | 12 archivos generados en memoria (CSV, Excel y TXT). |
| 17 | Metodología | Fórmulas, definiciones y limitaciones. |

---

## Barra lateral

- **Navegación** entre las 17 secciones.
- **Filtros globales**: rango de fechas, año, mes, entidad, régimen, sede, servicio, línea, estado, canal, soporte, glosa, saldo y rango de valor neto. Incluye *Seleccionar todos*, *Limpiar filtros* y el conteo de registros visibles. Todas las secciones usan el mismo conjunto filtrado.
- **Moneda (solo visual)**: COP, USD, EUR o sin símbolo. No convierte tasas de cambio.
- **Limpieza controlada**: eliminar filas vacías o duplicados exactos, rellenar categorías vacías con *Sin información*, excluir valores netos o fechas inválidas, y *Restaurar datos originales*. La limpieza se aplica sobre una copia.
- **Parámetros analíticos**: tolerancia de redondeo para las validaciones, rangos de días de radicación (por defecto 3, 7 y 15) y de antigüedad de cartera (30, 60, 90 y 180).
- **Umbrales de alertas**: valores demostrativos editables para % glosa, % cartera, % recaudo, días de radicación, % soportes incompletos y crecimiento.
- **Reiniciar aplicación**: limpia la sesión y vuelve a la pantalla inicial.

---

## Fórmulas principales

| Indicador | Fórmula |
|---|---|
| Porcentaje de glosa | `Glosa_Inicial / Valor_Neto × 100` |
| Porcentaje de recaudo | `Valor_Pagado / Valor_Neto × 100` |
| Porcentaje de cartera | `Saldo_Cartera / Valor_Neto × 100` |
| Ticket promedio | `Valor_Neto / número de facturas únicas` |
| Identidad del valor neto | `Valor_Neto ≈ Valor_Bruto − Descuento − Copago_Cuota` |
| Identidad del saldo | `Saldo_Cartera ≈ Valor_Neto − Glosa_Inicial − Valor_Pagado` |
| Límites IQR | `Q1 − k × IQR` y `Q3 + k × IQR` (k = 1,5 por defecto) |

Si un denominador es cero o nulo, el indicador se muestra como **No calculable**. Las facturas se cuentan con valores únicos de `ID_Factura`. El crecimiento y el promedio móvil solo se calculan entre periodos consecutivos.

---

## Descargas disponibles

| # | Archivo |
|---|---|
| 1-2 | `facturacion_filtrada.csv` / `facturacion_filtrada.xlsx` |
| 3-4 | `cierre_mensual.csv` / `cierre_mensual.xlsx` |
| 5 | `analisis_entidades.xlsx` (tabla comparativa y Pareto) |
| 6 | `analisis_servicios.xlsx` (servicios y líneas) |
| 7 | `analisis_sedes.xlsx` (sedes y regímenes) |
| 8 | `registros_glosados.csv` |
| 9 | `registros_cartera.csv` |
| 10 | `registros_atipicos.csv` |
| 11 | `informe_calidad.xlsx` (resumen, nulos, conversiones, limpieza e inconsistencias) |
| 12 | `insights_facturacion.txt` |

Los CSV usan punto y coma como separador y coma decimal, con codificación UTF-8 con BOM, para abrirse correctamente en Excel con configuración regional de Colombia. Con más de 20.000 registros filtrados, los archivos de registros individuales se generan cuando el usuario activa *Preparar archivos de registros*.

---

## Ejecución local

Requiere Python 3.10 o superior (se recomienda 3.12).

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS / Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

La aplicación abre en `http://localhost:8501`.

---

## Despliegue en Streamlit Community Cloud

1. Cree un repositorio en GitHub (público o privado).
2. Suba `app.py`, `requirements.txt` y `README.md` a la **raíz** del repositorio, sin subcarpetas.
3. Verifique que ambos archivos aparezcan en la raíz de la rama principal.
4. Ingrese a [share.streamlit.io](https://share.streamlit.io) e inicie sesión.
5. Conecte su cuenta de GitHub y autorice el acceso al repositorio.
6. Seleccione **Create app** → **Deploy a public app from GitHub**.
7. Elija el repositorio y la rama (por ejemplo, `main`).
8. Indique `app.py` como archivo principal (*Main file path*).
9. Opcional: en *Advanced settings*, seleccione Python 3.12.
10. Pulse **Deploy** y espere la instalación de dependencias.
11. Abra la URL que asigna Streamlit y cargue el archivo de facturación.

La aplicación no requiere secretos, bases de datos, API externas ni rutas locales.

---

## Privacidad y datos de salud

> Evite cargar archivos que contengan nombres de pacientes, documentos de identidad, historias clínicas, diagnósticos u otra información personal o sensible. Utilice archivos anonimizados y siga las políticas de tratamiento de datos de su organización.

- El archivo se procesa en memoria durante la sesión y no se escribe en disco.
- No se envían datos a servicios externos ni a modelos de inteligencia artificial.
- Al cerrar la sesión o pulsar *Reiniciar aplicación*, los datos se descartan.
- En una app pública de Streamlit Community Cloud, cualquier persona con el enlace puede usarla con sus propios archivos; cada sesión está aislada. Para restringir el acceso, despliéguela como app privada.

---

## Limitaciones

- La herramienta es exploratoria. No es un sistema oficial de facturación, contabilidad, auditoría ni historia clínica.
- Las correlaciones y pruebas estadísticas describen asociaciones; no demuestran causalidad.
- Los rangos de tiempo, de antigüedad y los umbrales de semáforos son parámetros analíticos editables, no estándares normativos ni condiciones contractuales.
- La antigüedad de cartera es aproximada (fecha de corte − fecha de radicación); no reemplaza el cálculo según las condiciones de pago pactadas con cada pagador.
- Las diferencias en la identidad del saldo pueden reflejar glosas levantadas o conciliadas, notas crédito o pagos posteriores, no necesariamente errores.
- El último periodo puede estar incompleto y los periodos recientes suelen mostrar menor recaudo.
- El puntaje de calidad es orientativo y no reemplaza una auditoría de datos.

---

## Validación realizada

Probada con Streamlit 1.50 y 1.65, pandas 2.2 y 3.0, y Plotly 5.24 y 7.1, usando datos sintéticos en estos casos:

| Caso | Resultado |
|---|---|
| Excel con varias hojas (selección automática de `Facturacion`) | Las 17 secciones funcionan |
| CSV con `;`, Windows-1252, encabezados alternos, `$ 1.234.567`, fechas dd/mm/aaaa, 25 fechas y 5 valores inválidos | Conversión correcta e informe de fallas |
| Archivo sin `Sede` / con `Sede` desasignada en el mapeo | Se desactiva solo el análisis por sede |
| Archivo sin `Valor_Neto` | Bloqueo controlado con instrucciones y mapeo manual |
| Archivo con solo 3 columnas mínimas | Funciona con avisos en las secciones no disponibles |
| Archivo vacío o solo con encabezados | Mensaje claro, sin cálculos |
| Filtros sin resultados | Advertencia y opción para limpiar filtros |
| 210.000 registros (48 MB) | Procesamiento inicial ≈ 12 s; navegación ≈ 1-2 s por sección |
