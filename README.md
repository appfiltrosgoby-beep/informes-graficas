# Dashboard Goby - Version Estatica

Aplicacion web 100% cliente (HTML, CSS y JavaScript) para cargar un Excel y generar las graficas del tablero de talleres aliados.

## Estructura

- `index.html`: interfaz principal.
- `static/app.js`: lectura del Excel en navegador y render de graficas.
- `static/styles.css`: estilos visuales.
- `static/logo-goby.svg`: logo alterno.
- `GOBY MARCA REGISTRADA.png`: logo principal.

## Uso local (sin servidor)

1. Abre `index.html` directamente en tu navegador.
2. Carga un archivo Excel (`.xlsx`, `.xlsm`, `.xls`).
3. Presiona `Procesar y generar dashboard`.

## Formato del Excel

Hojas esperadas por nombre aproximado:

- `treemap`
- `cobertura`
- `unidades_mes`
- `referencias`
- `consumo_taller`

Columnas minimas por hoja:

- `treemap`: Ciudad, Taller/Nombre, Cantidad
- `cobertura`: Mes, Despachados, Base datos
- `unidades_mes`: Mes, Total
- `referencias`: Referencia, Cantidad
- `consumo_taller`: Nombre, Cantidad

Si falta una hoja o columna, la app muestra advertencias y renderiza lo disponible.
"# informes-graficas" 
