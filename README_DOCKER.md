# Automatizacion TOA con Docker

Esta guia explica como ejecutar `toa_proceso_mejorado.py` dentro de un contenedor Playwright y como publicar la imagen con GitHub Actions.

El contenedor ejecuta `scheduler.py`. Por defecto corre una vez al iniciar y luego repite el proceso cada 30 minutos, alineado al reloj.

## Archivos del proyecto

```text
.
|-- Dockerfile
|-- compose.yaml
|-- .dockerignore
|-- .env.docker.example
|-- requirements.txt
|-- toa_proceso_mejorado.py
|-- scheduler.py
|-- healthcheck.py
|-- tests/
|-- .github/workflows/docker-image.yml
`-- data/
    |-- InputsTOA.xlsx
    |-- Monitoreo.xlsx
    `-- BBDD_TOA.xlsx
```

`data/` se monta como `/data` dentro del contenedor. No se copia dentro de la imagen Docker.

## Preparacion local

Abre PowerShell en la raiz del proyecto.

1. Copia la plantilla de entorno:

```powershell
Copy-Item .env.docker.example .env
```

2. Edita `.env` y completa credenciales, XPath y endpoints.

3. Crea la carpeta de datos:

```powershell
New-Item -ItemType Directory -Force data
```

4. Deja los archivos operativos dentro de `data/`:

- `InputsTOA.xlsx`: obligatorio.
- `Monitoreo.xlsx`: obligatorio si quieres registrar ejecuciones.
- `BBDD_TOA.xlsx`: opcional; si no existe, el proceso puede crear/reemplazar el consolidado segun la configuracion.

Ejemplo:

```powershell
Copy-Item .\InputsTOA.xlsx .\data\InputsTOA.xlsx
```

## Variables importantes para Docker

En Docker las rutas deben apuntar a `/data`, no a rutas Windows.

```env
RUTA_INPUTS_TOA=/data/InputsTOA.xlsx
RUTA_EXC_MONITOREO=/data/Monitoreo.xlsx
RUTA_BBDD_TOA=/data/BBDD_TOA.xlsx
TOA_RUNS_DIR=/data/toa_runs
DOWNLOAD_DIR=/data/toa_runs/downloads
TOA_BROWSER_CHANNEL=chromium
TOA_HEADLESS=true
```

Para ejecutar una sola vez y salir, usa:

```env
RUN_ONCE=true
RUN_ON_START=true
STOP_ON_ERROR=true
```

Para dejarlo como proceso programado:

```env
RUN_ONCE=false
RUN_ON_START=true
SCHEDULER_INTERVAL_MINUTES=30
STOP_ON_ERROR=false
```

## Ejecutar con Docker Compose

Construir y levantar:

```powershell
docker compose up -d --build
```

Ver logs:

```powershell
docker compose logs -f toa-scheduler
```

Ver estado:

```powershell
docker compose ps
```

Detener:

```powershell
docker compose down
```

Reconstruir desde cero si cambiaste dependencias o Dockerfile:

```powershell
docker compose build --no-cache
docker compose up -d
```

## Healthcheck

Docker ejecuta:

```text
python /app/healthcheck.py
```

El healthcheck lee `/data/toa_scheduler_state.json`. El contenedor se marca como no saludable si:

- no existe el archivo de estado;
- el estado esta vencido;
- el scheduler queda en `FATAL`.

Variable util:

```env
HEALTHCHECK_MAX_AGE_MINUTES=40
```

## Validacion antes de publicar

Antes de subir cambios, ejecuta:

```powershell
python -B -m unittest discover -s tests -v
docker compose config --quiet
```

Si Docker Desktop esta abierto, prueba tambien:

```powershell
docker build -t database-toa:test .
```

## GitHub Actions

El workflow esta en:

```text
.github/workflows/docker-image.yml
```

Hace dos cosas:

1. Ejecuta los tests Python.
2. Construye la imagen Docker.

En pull requests solo valida. En push a `main`, ademas publica en GitHub Container Registry.

Imagen publicada:

```text
ghcr.io/OWNER/REPO:latest
ghcr.io/OWNER/REPO:sha-<commit>
```

Reemplaza `OWNER/REPO` por el nombre real de tu repositorio. Por ejemplo, si el repositorio es `maximo/database_toa`:

```text
ghcr.io/maximo/database_toa:latest
```

## Activar GitHub Actions y GHCR

1. Sube el repositorio a GitHub.
2. En GitHub, entra a `Settings > Actions > General`.
3. En `Workflow permissions`, selecciona `Read and write permissions`.
4. Guarda los cambios.
5. Haz push a `main`.
6. Abre la pestana `Actions` y revisa el workflow `Docker image`.

No necesitas crear un token manual para publicar en GHCR. El workflow usa `secrets.GITHUB_TOKEN` con permiso `packages: write`.

## Ejecutar una imagen publicada

Primero inicia sesion si el paquete esta privado:

```powershell
docker login ghcr.io
```

Luego descarga la imagen:

```powershell
docker pull ghcr.io/OWNER/REPO:latest
```

Ejecuta una sola vez:

```powershell
docker run --rm --env-file .env -v ${PWD}/data:/data ghcr.io/OWNER/REPO:latest
```

Ejecuta en segundo plano:

```powershell
docker run -d --name toa-scheduler --restart unless-stopped --env-file .env -v ${PWD}/data:/data ghcr.io/OWNER/REPO:latest
```

Ver logs del contenedor publicado:

```powershell
docker logs -f toa-scheduler
```

Detenerlo:

```powershell
docker stop toa-scheduler
docker rm toa-scheduler
```

## Seguridad

- No subas `.env`.
- No subas `data/`.
- No subas Excel con datos reales.
- No guardes usuarios, passwords, tokens ni webhooks en el codigo.
- Usa GitHub Actions solo para construir la imagen, no para ejecutar el proceso real contra TOA.

## Errores comunes

### Docker Desktop no esta corriendo

Error tipico:

```text
failed to connect to the docker API
```

Solucion: abre Docker Desktop y espera a que el engine Linux quede activo.

### El contenedor no encuentra `InputsTOA.xlsx`

Verifica que exista:

```powershell
Get-ChildItem .\data
```

Y que `.env` use:

```env
RUTA_INPUTS_TOA=/data/InputsTOA.xlsx
```

### La imagen funciona local pero GitHub Actions falla al publicar

Revisa:

- que el push sea a `main`;
- que `Workflow permissions` tenga escritura;
- que el repositorio permita publicar paquetes;
- que el nombre de la imagen sea `ghcr.io/OWNER/REPO`.
