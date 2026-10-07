# Validador de retenciones de 5ta categoría (Perú)

Plantilla Excel para calcular y validar, **mes a mes y en forma anual por tramos**, las retenciones del
Impuesto a la Renta de 5ta categoría de varios trabajadores. El procedimiento sigue el art. 40 del
Reglamento de la LIR (D.S. N° 122-94-EF) y la escala del art. 53 de la LIR.

Archivo: [`Validador_Retenciones_5ta_Categoria.xlsx`](Validador_Retenciones_5ta_Categoria.xlsx)
(capacidad: 100 trabajadores y 1,500 filas de ingresos; ejercicio 2026, UIT S/ 5,500).

## Hojas

| Hoja | Uso |
|---|---|
| Instrucciones | Pasos, leyenda de colores, procedimiento normativo, supuestos y base legal. |
| Parámetros | Ejercicio, UIT por año, 7 UIT / 3 UIT, % bonificación Ley 30334, escala por tramos, divisores mensuales, tolerancia. |
| Conceptos | Catálogo editable: cada concepto tiene un tipo (ordinario, gratificación, bonificación, extraordinario, inafecto o retención de otro empleador). |
| Trabajadores | Datos de cada persona: fechas de ingreso y cese, seguro (EsSalud/EPS), opciones de proyección y gastos adicionales para la DJ. |
| Ingresos | Registro de ingresos con **frecuencia** (mensual, bimestral, trimestral, cuatrimestral, semestral, julio y diciembre, pago único) y **origen** (empresa principal u otro empleador). Los montos se distribuyen solos en Ene–Dic. |
| Retenciones Planilla | Retención efectivamente descontada cada mes (boleta / PLAME). |
| Cálculo Mensual | 12 filas por trabajador: proyección, impuesto por tramo, divisor, retención ordinaria y por extraordinarios, comparación con la planilla. |
| Resumen Anual | Renta e impuesto anual por tramos, retenciones calculadas vs. planilla y estimación de la DJ anual con la deducción adicional de hasta 3 UIT. |
| Ficha Trabajador | Vista imprimible de una persona: se elige el código y se ven los 12 meses y los tramos. |

El archivo incluye 4 trabajadores de ejemplo (códigos `EJ-001` a `EJ-004`): un aumento de sueldo,
utilidades, bonos semestrales y trimestrales, un ingreso a mitad de año con empleador anterior, un
empleo simultáneo con cese y un caso en el tramo de 30%. Hay 2 diferencias intencionales con la
planilla para mostrar la validación. Borre su contenido (sin eliminar filas) antes de usar la plantilla.

## Regenerar o ampliar la plantilla

```bash
python3 generar_validador_5ta.py --trabajadores 300 --filas-ingresos 5000 --sin-ejemplos
```

Requiere Python 3 y `openpyxl`. Todas las celdas de cálculo son fórmulas de Excel, de modo que la
plantilla recalcula al modificar los datos. Si genera un archivo nuevo, ábralo en Excel para que
calcule los resultados.

> Es una herramienta de apoyo: verifique los criterios con la normativa vigente y los pronunciamientos
> de SUNAT.
