# Plan de trabajo: Azure DevOps CSV → Dashboard GitHub Pages

## Arquitectura del flujo

```
Azure DevOps
    │
    ▼ (exportas manualmente)
"WAPSA Team - Epics.xlsx"  (fuente de datos raw)
    │
    ▼ (script Python: convert.py)
data.xlsx  ←── hoja "Hoja2", columnas exactas del dashboard
    │
    ▼ (git add + git commit + git push)
GitHub Pages → index.html + data.xlsx (dashboard actualizado)
```

---

## Columnas que necesita data.xlsx (hoja: `Hoja2`)

| Campo en dashboard | Columna en data.xlsx |
|--------------------|----------------------|
| ID | `ID` |
| Título | `Title` |
| Estado | `Estados` |
| Sprint | `Iteration Path` |
| Aplicación | `Aplicación` |
| Tipo Caso | `Tipo Caso` |
| Tipo Evento | `Tipo Evento` |
| Tipo Estabilización | `Tipo Estabilización` |
| Asignado | `Assigned To` |
| Épica | `Epic` |
| Inicio | `Start Date` |
| Fin SyC | `Finish Date` |
| Fin Chinalco | `Finish Date Chinalco` |

---

## Pasos pendientes

### 1. Instalar Python
```
winget install Python.Python.3.12
```
- Marcar "Add Python to PATH" durante la instalación.

### 2. Instalar dependencias Python
```
pip install openpyxl pandas
```

### 3. Analizar estructura del archivo fuente
- Abrir `WAPSA Team - Epics (5).xlsx` y mapear sus columnas a las que necesita `data.xlsx`
- Identificar columnas que vienen de Azure DevOps vs. columnas que se agregan manualmente

### 4. Crear script `convert.py`
- Lee `WAPSA Team - Epics.xlsx` (o el CSV exportado de Azure DevOps)
- Mapea y transforma columnas
- Genera `data.xlsx` con hoja `Hoja2`

### 5. Configurar GitHub Pages
- Verificar que el repo tenga GitHub Pages activado en la rama `main`
- Confirmar que `index.html` y `data.xlsx` estén en la raíz

### 6. Flujo de actualización (cada vez que hay datos nuevos)
```bash
# 1. Exportar nuevo archivo desde Azure DevOps y reemplazar el xlsx fuente
# 2. Ejecutar el script de conversión
python convert.py

# 3. Subir cambios
git add data.xlsx
git commit -m "Actualizar datos - <fecha>"
git push
```
El dashboard en GitHub Pages se actualiza automáticamente en ~1 minuto.

---

## Archivos del proyecto

| Archivo | Rol |
|---------|-----|
| `WAPSA Team - Epics.xlsx` | Fuente de datos raw (Azure DevOps) |
| `convert.py` | Script de transformación (por crear) |
| `data.xlsx` | Excel final que consume el dashboard |
| `index.html` | Dashboard (GitHub Pages) |
| `PLAN.md` | Este archivo |

---

## Estado actual

- [x] Python instalado (3.14.3)
- [x] Dependencias instaladas (`openpyxl`, `pandas`)
- [x] Estructura del xlsx fuente analizada
- [x] Script `convert.py` creado y probado (207 tickets, 10 épicas)
- [ ] GitHub Pages configurado
- [ ] Primer ciclo de actualización probado
