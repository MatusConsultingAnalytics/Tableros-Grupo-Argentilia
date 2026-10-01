# Tablero Ejecutivo — Grupo Gastronómico Argentilia

Tablero publicado en GitHub Pages, protegido con contraseña y actualizado automáticamente de lunes a viernes a las 12:00 (hora CDMX).

## Fuentes de información

| Pestañas | Fuente | Actualización |
|---|---|---|
| Ranking, Cumplimiento, Detalle, Mix, Días, 2025 vs 2026 | `Plantilla Captura Argentilia 2026.xlsx` (Drive) | Automática diaria |
| Rentabilidad, Plan de recuperación | Estados de resultados de Contraloría, carpeta de Drive `1jDKGW4dhuy9JKvZ2g8Zuwx17qeE7kBcT` | Automática: se refleja en la siguiente corrida después de subir los archivos |

### Reglas para los estados financieros

1. Un Excel por restaurante y mes, con el nombre exacto: `UNIDAD MES AÑO.xlsx`
   - Unidades válidas: `ARGENTILIA LEON`, `ARGENTILIA QRO`, `FRASCATI`, `MIKOH`, `MANOLO`
   - Ejemplo: `ARGENTILIA LEON SEPTIEMBRE 2026.xlsx`
2. Si Contraloría corrige un mes, se reemplaza el archivo con el mismo nombre (sin "v2" ni "final").
3. No renombrar ni reordenar las hojas del archivo de Contraloría.
4. La cifra oficial de cada mes es la columna `TOTAL <MES> <AÑO>`.
5. El tablero muestra el último mes que tengan **todas** las unidades; si falta alguna, avisa cuáles.
6. Los archivos consolidados (p. ej. `AGOSTO 2026.xlsx`) se ignoran.
7. Si el archivo más reciente de una unidad trae un mes sin presupuesto de venta (vacío, 0 o −1), el presupuesto de ese mes se toma del archivo de cierre del propio mes (p. ej. `ARGENTILIA QRO FEBRERO 2026.xlsx`) y se avisa en las notas de calidad de datos.

## Archivos

- `generar_tablero.py` — genera el tablero completo (`index.html`).
- `rentabilidad.py` — pestaña Rentabilidad: lectura de estados de resultados, cálculos, textos automáticos y palancas.
  - Umbrales del semáforo, fechas de apertura y responsables de palancas se configuran al inicio del archivo.
- `plan_recuperacion.py` — pestaña Plan de recuperación: simulador de cheque promedio sobre el mismo modelo de Rentabilidad.
  - Solo se mueve el cheque; comensales fijos. Utilidad adicional = venta adicional × (1 − % costo de venta).
  - Cumplimiento del grupo **sin compensar**: el excedente de una unidad no cubre el faltante de otra.
  - "Cheque para meta" calcula el ajuste exacto por unidad. El cálculo corre en el navegador; el ajuste inicial (+$15) se configura en `AJUSTE_INICIAL`.
- `.github/workflows/automatizacion.yml` — genera, cifra y publica.

## Secretos requeridos (Settings → Secrets and variables → Actions)

| Secreto | Contenido |
|---|---|
| `TABLERO_PASSWORD` | Contraseña para abrir el tablero. Sin este secreto el workflow **no publica**. |
| `GDRIVE_SA_KEY` | JSON completo de la llave de la cuenta de servicio de Google con acceso de lectura a la carpeta de estados financieros y a la plantilla. |

## Seguridad

- El sitio publicado está cifrado (StatiCrypt, AES-256): sin la contraseña el contenido es ilegible, incluso viendo el código fuente.
- Solo se publica la carpeta cifrada y sin historial previo (`force_orphan`).
- Para cambiar la contraseña basta con actualizar `TABLERO_PASSWORD` y ejecutar el workflow manualmente.
