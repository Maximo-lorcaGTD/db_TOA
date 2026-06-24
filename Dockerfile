FROM mcr.microsoft.com/playwright/python:v1.59.0-noble

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    TZ=America/Santiago

WORKDIR /app

COPY requirements.txt /app/requirements.txt
RUN python -m pip install --no-cache-dir -r /app/requirements.txt

COPY toa_proceso_mejorado.py scheduler.py healthcheck.py /app/

RUN mkdir -p /data/toa_runs/downloads \
    && chmod +x /app/scheduler.py /app/healthcheck.py

HEALTHCHECK --interval=1m --timeout=10s --start-period=5m --retries=3 CMD python /app/healthcheck.py

CMD ["python", "-u", "/app/scheduler.py"]