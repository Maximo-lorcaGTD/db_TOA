from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


def env_int(name: str, default: int, minimum: int = 1) -> int:
    raw = os.getenv(name, str(default)).strip()
    try:
        value = int(raw)
    except ValueError as exc:
        raise RuntimeError(f"{name} debe ser un número entero: {raw!r}") from exc
    if value < minimum:
        raise RuntimeError(f"{name} debe ser igual o mayor que {minimum}")
    return value


def main() -> int:
    try:
        state_path = Path(
            os.getenv("SCHEDULER_STATE_FILE", "/data/toa_scheduler_state.json")
        )
        timezone = ZoneInfo(os.getenv("TZ", "America/Santiago"))
        max_age_minutes = env_int("HEALTHCHECK_MAX_AGE_MINUTES", 40)
    except Exception as exc:
        print(f"Configuración inválida: {exc!r}")
        return 1

    if not state_path.is_file():
        print(f"No existe estado: {state_path}")
        return 1

    try:
        state = json.loads(state_path.read_text(encoding="utf-8"))
        updated = datetime.fromisoformat(state["scheduler_updated_at"])
        if updated.tzinfo is None:
            updated = updated.replace(tzinfo=timezone)
    except Exception as exc:
        print(f"Estado inválido: {exc!r}")
        return 1

    age_seconds = (
        datetime.now(timezone) - updated.astimezone(timezone)
    ).total_seconds()
    if age_seconds > max_age_minutes * 60:
        print(f"Estado desactualizado: {age_seconds:.0f}s")
        return 1

    if state.get("status") == "FATAL":
        print(f"Estado fatal: {state.get('last_error')}")
        return 1

    print(f"OK status={state.get('status')} age={age_seconds:.0f}s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
