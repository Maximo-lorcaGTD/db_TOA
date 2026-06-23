# Automatización TOA con Playwright

Este proyecto automatiza el ingreso a Oracle Field Service (TOA), descarga los archivos correspondientes a las fechas y secciones configuradas, lee cada archivo, consolida la información en un único Excel y luego elimina los archivos temporales descargados.

La solución utiliza **Playwright con Microsoft Edge**. No utiliza Selenium, EdgeDriver ni PostgreSQL.

## Flujo de ejecución

```text
Inicio de sesión en TOA
        ↓
Descarga por fecha y sección
        ↓
Validación del archivo descargado
        ↓
Lectura de sus registros
        ↓
Eliminación inmediata del archivo temporal
        ↓
Consolidación de todas las filas en memoria
        ↓
Creación o reemplazo del Excel consolidado
        ↓
Envío opcional a Power Automate
        ↓
Registro de la ejecución y cierre de sesión
```

Cada archivo descargado se utiliza únicamente como archivo temporal. Después de leerlo correctamente, el script lo elimina antes de continuar con la siguiente descarga.

Los archivos vacíos también se consideran válidos: se revisan, se contabilizan y se eliminan sin agregar filas al consolidado.

## Instalación

Abre PowerShell dentro de la carpeta del proyecto y ejecuta:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install --upgrade pip
py -m pip install playwright python-dotenv openpyxl requests
py -m playwright install msedge
```

No se necesita `msedgedriver.exe` ni la variable `RUTA_DRIVER`.

## Configuración del archivo `.env`

Crea un archivo llamado exactamente `.env` en la misma carpeta donde se encuentra el script.

Ejemplo:

```env
# =====================================
# Credenciales TOA
# =====================================
TOA_USERNAME=TU_USUARIO
TOA_PASSWORD=TU_CONTRASEÑA

# =====================================
# Selectores del inicio de sesión
# =====================================
TOA_XPATH_USER=/html/body/div/div/div/div/form/div[2]/div[1]/div[1]/div/input
TOA_XPATH_PASS=/html/body/div/div/div/div/form/div[2]/div[2]/div[1]/div/input
TOA_XPATH_BTN=/html/body/div/div/div/div/form/div[2]/div[5]/button
TOA_XPATH_BTN_ALT=/html/body/div/div/div/div/form/div[2]/div[6]/button

# Elemento utilizado para confirmar que la sesión cargó
TOA_XPATH_READY=/html/body/div[14]/div[1]/main/div/div[2]/div[3]/div[1]/div[3]/div[1]/div[2]/div[2]/div/div[2]/div[3]/div[2]/div[2]/div[5]/div[1]/button[3]/span

# Elementos utilizados para cerrar sesión
TOA_XPATH_AVATAR=/html/body/div[14]/div[1]/app:global-header/header/div[5]/button/visuals:technician-avatar
TOA_XPATH_SIGNOUT=/html/body/div[26]/div[2]/app:global-header/div/ul/li[4]/a

# =====================================
# Navegador
# =====================================
TOA_BROWSER_CHANNEL=msedge
TOA_HEADLESS=false
TOA_SLOW_MO_MS=100

# =====================================
# Rutas locales
# =====================================
RUTA_INPUTS_TOA=C:\RUTA\InputsTOA.xlsx
RUTA_EXC_MONITOREO=C:\RUTA\Monitoreo.xlsx
RUTA_BBDD_TOA=C:\RUTA\BBDD_TOA.xlsx
DOWNLOAD_DIR=C:\Users\USUARIO\Downloads
TOA_RUNS_DIR=C:\RUTA\toa_runs

# =====================================
# Rango de fechas
# =====================================
TOA_DAYS_BACK=3
TOA_DAYS_FORWARD=14

# =====================================
# Power Automate
# =====================================
PA_ENABLED=true
PA_WEBHOOK_URL=PEGA_AQUI_UN_WEBHOOK_VIGENTE
PA_DESTINATION_NAME=BBDD_TOA.xlsx

# =====================================
# Limpieza
# =====================================
CLEAN_RUN_MANIFEST=false
```

No agregues comillas alrededor de las rutas, credenciales o XPath.

## Archivo `InputsTOA.xlsx`

El script utiliza únicamente las filas cuyo campo `ESTADO_EN_SCRIPT` sea igual a `ACTIVO`.

El archivo debe incluir estas columnas:

```text
Nombre Archivo Completo
Subgrupo TOA
Zona Subgrupo TOA
URL
RUTA_Carpeta_Destino
ESTADO_EN_SCRIPT
providerId
downloadId
recursively
```

La cantidad esperada de descargas se calcula de forma dinámica:

```text
Cantidad de fechas × cantidad de registros ACTIVOS
```

El rango incluye ambos extremos. Por ejemplo:

```env
TOA_DAYS_BACK=3
TOA_DAYS_FORWARD=14
```

representa 18 fechas: desde 3 días antes hasta 14 días después de la fecha actual.

Para trabajar con 17 fechas usando 3 días anteriores, puedes configurar:

```env
TOA_DAYS_BACK=3
TOA_DAYS_FORWARD=13
```

## Tratamiento de los archivos descargados

Por cada descarga, el script realiza lo siguiente:

1. Guarda temporalmente el archivo en `DOWNLOAD_DIR`.
2. Comprueba que sea un archivo Excel legible.
3. Lee sus encabezados y registros.
4. Agrega las filas válidas al consolidado en memoria.
5. Identifica si el archivo está vacío.
6. Elimina el archivo descargado.
7. Continúa con la siguiente fecha y sección.

En consola debería aparecer un mensaje parecido a:

```text
[DESCARGA] CENTRO OK → archivo.xlsx | filas leídas: 25 | archivo eliminado
```

El Excel individual no permanece en el equipo después de ser procesado.

## Consolidado final

Cuando todas las descargas terminan correctamente, el script crea o reemplaza el archivo definido en:

```env
RUTA_BBDD_TOA=C:\RUTA\BBDD_TOA.xlsx
```

Este archivo contiene la fotografía completa y actual obtenida durante la ejecución.

El consolidado final no se elimina inmediatamente, porque se utiliza para:

- mantener una copia consolidada local;
- enviarlo a Power Automate;
- reemplazar el archivo utilizado posteriormente en SharePoint u otro destino.

## Control de ejecución incompleta

El script compara:

- archivos esperados;
- archivos descargados correctamente;
- archivos vacíos;
- archivos con datos;
- errores de descarga o lectura.

Si falta un archivo o se produce un error bloqueante, la ejecución se marca como incompleta y no debe reemplazar el consolidado vigente.

## Power Automate

Cuando esta variable está habilitada:

```env
PA_ENABLED=true
```

el consolidado se convierte a Base64 y se envía al webhook configurado en:

```env
PA_WEBHOOK_URL=
```

Para ejecutar el proceso sin enviar el archivo a Power Automate:

```env
PA_ENABLED=false
```

Por seguridad, no guardes el archivo `.env` en repositorios Git ni compartas públicamente el webhook, usuario o contraseña.

## Archivo de monitoreo

La ruta configurada en:

```env
RUTA_EXC_MONITOREO=C:\RUTA\Monitoreo.xlsx
```

se utiliza para registrar información de cada ejecución, como:

- fecha y hora de inicio;
- fecha y hora de término;
- cantidad de descargas por sección;
- resultado de la ejecución;
- errores encontrados;
- resultado del consolidado.

El libro debe contener una tabla llamada `Ejecuciones` dentro de una hoja llamada `Ejecuciones`.

## Manifiesto de auditoría

Cada ejecución puede generar un archivo `manifest.json` dentro de `TOA_RUNS_DIR`.

Este archivo registra:

- fecha procesada;
- sección;
- nombre del archivo;
- cantidad de filas leídas;
- si estaba vacío;
- confirmación de eliminación después de la lectura;
- errores de la ejecución.

Para eliminar automáticamente el manifiesto y su carpeta después de una ejecución exitosa:

```env
CLEAN_RUN_MANIFEST=true
```

Para conservarlo como evidencia de auditoría:

```env
CLEAN_RUN_MANIFEST=false
```

## Ejecución manual

Con el entorno virtual activo:

```powershell
py .\toa_descargas_postgresql.py
```

Aunque el archivo todavía pueda conservar el nombre histórico `toa_descargas_postgresql.py`, esta versión del proceso no utiliza PostgreSQL. Se recomienda renombrarlo a:

```text
toa_descargas_playwright.py
```

Luego se ejecutaría con:

```powershell
py .\toa_descargas_playwright.py
```

## Ejecución cada 30 minutos

Se puede configurar mediante el Programador de tareas de Windows.

### Programa o script

```text
C:\RUTA\PROYECTO\.venv\Scripts\python.exe
```

### Agregar argumentos

```text
C:\RUTA\PROYECTO\toa_descargas_playwright.py
```

### Iniciar en

```text
C:\RUTA\PROYECTO
```

Configura el desencadenador para repetirse cada 30 minutos.

Antes de programarlo, comprueba manualmente que una ejecución completa termine correctamente.

## Errores frecuentes

### Faltan variables en `.env`

Comprueba que el archivo se llame exactamente `.env`, esté junto al script y contenga las variables obligatorias.

```powershell
Get-ChildItem -Force
```

### Playwright no encuentra Microsoft Edge

```powershell
py -m playwright install msedge
```

### El campo de usuario o contraseña no aparece

Revisa los valores de:

```env
TOA_XPATH_USER=
TOA_XPATH_PASS=
TOA_XPATH_BTN=
```

### El archivo descargado no puede leerse

Comprueba que TOA esté entregando un archivo `.xlsx` real y no una respuesta CSV, HTML o una página de error.

### El archivo no se elimina

Verifica que no esté abierto en Excel y que el usuario que ejecuta el script tenga permisos sobre `DOWNLOAD_DIR`.

## Seguridad

- No publiques el archivo `.env`.
- No guardes contraseñas dentro del código Python.
- No compartas públicamente el webhook de Power Automate.
- Si un webhook se expone, reemplázalo por uno nuevo.
- Utiliza una cuenta TOA con los permisos mínimos necesarios.

## Resultado esperado

Una ejecución exitosa debe finalizar con un resumen parecido a:

```text
[RESUMEN] Esperados: 272 | Descargados: 272
[CONSOLIDADO] Archivos con datos: 210 | Archivos vacíos: 62 | Filas: 15430
[WEBHOOK] Archivo enviado exitosamente a Power Automate
---- EJECUCIÓN COMPLETA ----
```

Las cantidades son referenciales y dependen del rango de fechas y de las secciones activas en `InputsTOA.xlsx`.
