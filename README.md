# BOOST CN v10.0 · Motor Excel

PWA estática para GitHub Pages. `data/Boost_CN.xlsx` es la única fuente de información: GitHub Actions lee sus cinco pestañas, normaliza relaciones, valida todas las métricas y regenera los JSON consumidos por el dashboard.

La interfaz concentra el resumen ejecutivo y el detalle operativo por franja en una sola vista Inicio. La navegación mínima está integrada al encabezado y no utiliza panel lateral.

## Actualizar datos

1. Reemplazar `data/Boost_CN.xlsx` conservando el nombre y las pestañas.
2. Subir el archivo a la rama `main`.
3. El workflow `.github/workflows/actualizar-excel.yml` procesa, valida y publica los JSON automáticamente.

No es necesario modificar `index.html`, JavaScript ni los JSON al actualizar información. El Service Worker obtiene la lista de fragmentos desde `data/manifest-data.json` y consulta los datos con prioridad de red para evitar mostrar una versión anterior.

La lectura se realiza por nombre de pestaña y encabezado. El cruce de productos usa `Condicion[Producto]`, normalizando espacios, mayúsculas y acentos. `Base_AT` se relaciona por `CeCo + Semana + Año`; el motor detecta y corrige de forma auditable el desplazamiento conocido entre `Semana`, `CeCo` y `Tienda`. `Item_ADT` se relaciona por `CeCo + Semana`.

## Validación local opcional

```bash
python tools/process_excel.py
python tools/validate_excel_engine.py
```
