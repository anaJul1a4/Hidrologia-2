# Hidrologia-2 · Tarea 2 (evapotranspiración y diseño hidrológico)

Sitio de diapositivas de la Tarea 2 de Hidrología 2026-2 (UNAL Medellín). Punto 1: cuenca del río Chicama hasta
El Tambo (Perú), la misma de la [Tarea 1](https://anajul1a4.github.io/Hidrologia-Chicama-El-Tambo/). Punto 2: Cottbus (Alemania).

**Origen.** El sitio parte de la versión de Marlin Julie Schwebel
([mschwebel-eng/Tarea2](https://github.com/mschwebel-eng/Tarea2)). Sobre esa base se agregaron:

- la línea gráfica de la Tarea 1 (capa final del CSS de `index.html`) y los bloques «Qué muestra / Cómo se construyó»;
- la diapositiva `t1-resumen` (la Tarea 1 como punto de partida);
- la diapositiva `p1-balance` (balance diario del suelo, adaptado de la figura #s14 de las notas del curso);
- diapositivas de conceptos (`t1-conceptos`, `p2-conceptos`), de las obras (`p24-obra`, `p24b-obra`),
  la comparación con el Valle de Aburrá (`p2-aburra`) y los escenarios de cambio climático (`p25-clima`).

## Estructura

```
index.html                      diapositivas (navegación con ← → o los puntos inferiores)
assets/, img/                   Plotly e imágenes
codigo/punto1/00_balance_cubeta.py   balance diario del suelo (diapositiva p1-balance)
codigo/punto2/10_geometria_obras.py  ortofoto, alcantarillas reales de OSM e hidráulica de las obras (requiere Pillow y pyproj)
codigo/punto2/11_escenarios_clima.py escenarios de cambio climático sobre la IDF y la alcantarilla
datos/punto2/                   ortofoto DOP20 (GeoBasis-DE/LGB, dl-de/by-2-0) y extracto de OpenStreetMap (ODbL)
datos/punto1/                   entradas recortadas a la cuenca y al periodo (CAMELS-PE 1998–2016)
datos/punto1/tarea1/            productos de la Tarea 1 que se reutilizan
productos/punto1/               salidas: balance_cubeta_datos.js (lo lee index.html), CSV mensual y anual
```

## Reproducir

```
python codigo/punto1/00_balance_cubeta.py   balance diario del suelo (diapositiva p1-balance)
codigo/punto2/10_geometria_obras.py  ortofoto, alcantarillas reales de OSM e hidráulica de las obras (requiere Pillow y pyproj)
codigo/punto2/11_escenarios_clima.py escenarios de cambio climático sobre la IDF y la alcantarilla
datos/punto2/                   ortofoto DOP20 (GeoBasis-DE/LGB, dl-de/by-2-0) y extracto de OpenStreetMap (ODbL)
```

Solo usa la biblioteca estándar de Python 3. Si existe `datos/punto1/eto_era5land_diaria_cuenca.csv`
(columnas `date`, `eto_mm`: ETo FAO-56 diaria de ERA5-Land del apartado 1.1.a), la usa como segunda demanda en lugar de
PISCOeo × 0,906.

## Fuentes de datos

- CAMELS-PE, cuenca `PE_47E0A2A8`: `prec` (PISCOp v2.1), `pet` (PISCOeo_pm, termina en 2016-12) y `flow_obs` diarios.
- IMERG V07 Final mensual promediado en la cuenca, y caudal mensual rellenado: Tarea 1 del grupo.
