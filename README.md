# WAPSA Dashboard

Dashboard de tickets del equipo WAPSA, publicado en GitHub Pages.

## Archivos del proyecto

| Archivo | Descripcion |
|---------|-------------|
| `index.html` | Dashboard web (no modificar) |
| `data.xlsx` | Datos que consume el dashboard (se genera automaticamente) |
| `convert.py` | Script que transforma el export de Azure DevOps |
| `actualizar.bat` | Doble click para actualizar todo |
| `input/` | Carpeta donde se pegan los exports de Azure DevOps |

## Como actualizar el dashboard

1. Exportar el archivo de Azure DevOps (xlsx)
2. Pegarlo en la carpeta `input/` (borrar el anterior si queres)
3. Doble click en `actualizar.bat`
4. Esperar ~1 minuto — el dashboard se actualiza solo

## Ver el dashboard

https://justinfmv.github.io/WAPSA---DASHBOARD/
