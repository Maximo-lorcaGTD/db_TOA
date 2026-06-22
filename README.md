# TOA con Playwright y PostgreSQL

Esta versión reemplaza Selenium por **Playwright para Python** y utiliza los XPath del proyecto original. Mantiene login, descargas por fecha y sección, cierre de sesión, consolidado, publicación atómica en PostgreSQL y envío opcional a Power Automate.

## Comportamiento de cada descarga

Cada archivo sigue este flujo:

```text
Descargar XLSX → validar encabezados → leer filas → guardar filas en memoria → eliminar XLSX
```

El archivo descargado no queda almacenado en las carpetas de sección. Aunque esté vacío, se considera una descarga válida, se contabiliza y se elimina después de leer sus encabezados. Si no se pueden descargar o leer todos los archivos esperados, PostgreSQL no se reemplaza.

Solo se conserva `manifest.json` dentro de `TOA_RUNS_DIR` para auditoría. Puedes eliminarlo automáticamente con:

```env
CLEAN_RUN_MANIFEST=true
```

El Excel indicado por `RUTA_BBDD_TOA` es el consolidado final utilizado para Power Automate; no corresponde a uno de los archivos descargados desde TOA.

## Cambios

- No usa `selenium`, `webdriver`, `msedgedriver.exe` ni `RUTA_DRIVER`.
- Abre Microsoft Edge mediante `TOA_BROWSER_CHANNEL=msedge`.
- Usa los XPath originales para usuario, contraseña, botón, sesión cargada, avatar y cierre de sesión.
- Los XPath tienen valores predeterminados dentro del script, pero también se pueden modificar desde `.env`.
- Las descargas usan `page.expect_download()` y `download.save_as()`.
- Cada archivo fuente se elimina inmediatamente después de ser leído.
- Si falta una descarga o falla la lectura de un archivo, PostgreSQL conserva la fotografía anterior.

## Instalación en Windows

Abre PowerShell dentro de la carpeta del proyecto:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m playwright install msedge
```

### Alternativa con Chromium

```powershell
py -m playwright install chromium
```

Luego deja esta variable vacía:

```env
TOA_BROWSER_CHANNEL=
```

## Configuración

```powershell
Copy-Item .env.example .env
notepad .env
```

Completa credenciales, rutas, PostgreSQL y un webhook vigente. No uses `RUTA_DRIVER`.

Durante las primeras pruebas:

```env
TOA_HEADLESS=false
TOA_SLOW_MO_MS=100
```

## Ejecución

```powershell
py .\toa_descargas_postgresql.py
```

También se incluye el mismo código con el nombre descriptivo:

```powershell
py .\toa_descargas_postgresql_playwright.py
```

No ejecutes ambos archivos al mismo tiempo.

## Programador de tareas

```text
Programa:
C:\ruta\proyecto\.venv\Scripts\python.exe

Argumentos:
C:\ruta\proyecto\toa_descargas_postgresql.py

Iniciar en:
C:\ruta\proyecto
```

## Conteo esperado

```text
cantidad de fechas × cantidad de registros ACTIVOS en InputsTOA.xlsx
```

Con 18 fechas y 16 secciones se esperan 288 archivos. Una fotografía incompleta no reemplaza `toa_actual`.
