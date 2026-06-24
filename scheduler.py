#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Planificador secuencial para ejecutar el proceso TOA cada N minutos.

Características:
- Ejecuta inmediatamente al levantar el contenedor si RUN_ON_START=true.
- Después alinea las ejecuciones al reloj (por defecto :00 y :30).
- Nunca ejecuta dos procesos simultáneos.
- Usa un lock para impedir dos contenedores/planificadores sobre el mismo volumen.
- Guarda estado y resultados en JSON para observabilidad/healthcheck.
"""

from __future__ import annotations

import json
import os
import signal
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

try:
    import fcntl
except ImportError:  # pragma: no cover - solo aplica en Windows
    fcntl = None

try:
    import msvcrt
except ImportError:  # pragma: no cover - solo aplica en Linux/Docker
    msvcrt = None


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "si", "sí", "on"}


def env_int(name: str, default: int, minimum: int = 1) -> int:
    raw = os.getenv(name, str(default)).strip()
    try:
        value = int(raw)
    except ValueError as exc:
        raise RuntimeError(f"{name} debe ser un número entero: {raw!r}") from exc
    if value < minimum:
        raise RuntimeError(f"{name} debe ser igual o mayor que {minimum}")
    return value


TIMEZONE_NAME = os.getenv("TZ", "America/Santiago")
TZ = ZoneInfo(TIMEZONE_NAME)
INTERVAL_MINUTES = env_int("SCHEDULER_INTERVAL_MINUTES", 30)
RUN_ON_START = env_bool("RUN_ON_START", True)
RUN_ONCE = env_bool("RUN_ONCE", False)
STOP_ON_ERROR = env_bool("STOP_ON_ERROR", False)
SCRIPT_PATH = Path(os.getenv("TOA_SCRIPT_PATH", "/app/toa_proceso_mejorado.py"))
STATE_FILE = Path(os.getenv("SCHEDULER_STATE_FILE", "/data/toa_scheduler_state.json"))
LOCK_FILE = Path(os.getenv("SCHEDULER_LOCK_FILE", "/data/.toa_scheduler.lock"))

stop_requested = False
current_child: subprocess.Popen | None = None


def now_local() -> datetime:
    return datetime.now(TZ)


def atomic_write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)
        file.flush()
        os.fsync(file.fileno())
    os.replace(temporary, path)


def next_aligned_time(current: datetime, interval_minutes: int) -> datetime:
    """Obtiene el siguiente múltiplo de N minutos desde medianoche."""
    current = current.astimezone(TZ)
    minutes_today = current.hour * 60 + current.minute
    next_total = ((minutes_today // interval_minutes) + 1) * interval_minutes
    day_offset, minute_of_day = divmod(next_total, 24 * 60)
    hour, minute = divmod(minute_of_day, 60)

    target = current.replace(
        hour=hour,
        minute=minute,
        second=0,
        microsecond=0,
    ) + timedelta(days=day_offset)
    return target


def write_state(**changes) -> dict:
    state = {}
    if STATE_FILE.exists():
        try:
            state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            state = {}

    state.update(changes)
    state["scheduler_updated_at"] = now_local().isoformat()
    state["timezone"] = TIMEZONE_NAME
    state["interval_minutes"] = INTERVAL_MINUTES
    atomic_write_json(STATE_FILE, state)
    return state


def handle_signal(signum, _frame) -> None:
    global stop_requested, current_child
    stop_requested = True
    print(f"[SCHEDULER] Señal {signum} recibida. Finalizando...", flush=True)

    if current_child is not None and current_child.poll() is None:
        try:
            current_child.terminate()
        except Exception:
            pass


signal.signal(signal.SIGTERM, handle_signal)
signal.signal(signal.SIGINT, handle_signal)


def run_toa_process() -> int:
    global current_child

    if not SCRIPT_PATH.is_file():
        raise FileNotFoundError(f"No se encontró el script principal: {SCRIPT_PATH}")

    started = now_local()
    print("=" * 78, flush=True)
    print(f"[SCHEDULER] Inicio de ejecución: {started.isoformat()}", flush=True)
    print(f"[SCHEDULER] Script: {SCRIPT_PATH}", flush=True)

    write_state(
        status="RUNNING",
        current_run_started_at=started.isoformat(),
        last_error=None,
    )

    current_child = subprocess.Popen(
        [sys.executable, "-u", str(SCRIPT_PATH)],
        cwd=str(SCRIPT_PATH.parent),
        env=os.environ.copy(),
    )
    return_code = current_child.wait()
    current_child = None

    finished = now_local()
    duration_seconds = round((finished - started).total_seconds(), 3)
    success = return_code == 0

    print(
        f"[SCHEDULER] Fin: {finished.isoformat()} | "
        f"código={return_code} | duración={duration_seconds}s",
        flush=True,
    )

    write_state(
        status="IDLE" if success else "ERROR",
        last_run_started_at=started.isoformat(),
        last_run_finished_at=finished.isoformat(),
        last_run_duration_seconds=duration_seconds,
        last_exit_code=return_code,
        last_success_at=finished.isoformat() if success else None,
        last_error=None if success else f"El proceso terminó con código {return_code}",
        current_run_started_at=None,
    )

    return return_code


def sleep_until(target: datetime) -> None:
    write_state(status="WAITING", next_run_at=target.isoformat())
    print(f"[SCHEDULER] Próxima ejecución: {target.isoformat()}", flush=True)

    while not stop_requested:
        remaining = (target - now_local()).total_seconds()
        if remaining <= 0:
            return
        time.sleep(min(remaining, 5.0))


def lock_file(lock_handle) -> None:
    if fcntl is not None:
        fcntl.flock(lock_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        return

    if msvcrt is not None:
        lock_handle.seek(0)
        if not lock_handle.read(1):
            lock_handle.write("\0")
            lock_handle.flush()
        lock_handle.seek(0)
        try:
            msvcrt.locking(lock_handle.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise BlockingIOError(
                "Ya existe otro planificador TOA usando el mismo volumen/lock."
            ) from exc
        return

    raise RuntimeError("Este sistema no soporta bloqueo de archivo.")


def unlock_file(lock_handle) -> None:
    if fcntl is not None:
        fcntl.flock(lock_handle.fileno(), fcntl.LOCK_UN)
        return

    if msvcrt is not None:
        lock_handle.seek(0)
        msvcrt.locking(lock_handle.fileno(), msvcrt.LK_UNLCK, 1)


def acquire_lock():
    LOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
    lock_handle = LOCK_FILE.open("a+", encoding="utf-8")
    try:
        lock_file(lock_handle)
    except BlockingIOError as exc:
        lock_handle.close()
        raise RuntimeError(
            "Ya existe otro planificador TOA usando el mismo volumen/lock."
        ) from exc

    lock_handle.seek(0)
    lock_handle.truncate()
    lock_handle.write(str(os.getpid()))
    lock_handle.flush()
    return lock_handle


def main() -> int:
    if INTERVAL_MINUTES > 24 * 60:
        raise RuntimeError("SCHEDULER_INTERVAL_MINUTES no puede superar 1440")

    lock_handle = acquire_lock()
    print(
        f"[SCHEDULER] Activo | intervalo={INTERVAL_MINUTES} min | "
        f"zona={TIMEZONE_NAME} | run_on_start={RUN_ON_START}",
        flush=True,
    )

    try:
        if RUN_ON_START and not stop_requested:
            code = run_toa_process()
            if RUN_ONCE:
                return code
            if code != 0 and STOP_ON_ERROR:
                return code
        elif RUN_ONCE:
            return run_toa_process()

        while not stop_requested:
            target = next_aligned_time(now_local(), INTERVAL_MINUTES)
            sleep_until(target)
            if stop_requested:
                break

            code = run_toa_process()
            if code != 0 and STOP_ON_ERROR:
                return code

        write_state(status="STOPPED", next_run_at=None)
        return 0
    finally:
        try:
            unlock_file(lock_handle)
        finally:
            lock_handle.close()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"[SCHEDULER] ERROR CRÍTICO: {exc!r}", file=sys.stderr, flush=True)
        try:
            write_state(status="FATAL", last_error=repr(exc), next_run_at=None)
        except Exception:
            pass
        raise SystemExit(1)
