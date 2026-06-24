from __future__ import annotations

import json
import os
import sys
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch
from zoneinfo import ZoneInfo

# ============================================================
# RUTA RAÍZ DEL PROYECTO
# ============================================================
# Debe agregarse ANTES de importar los módulos locales.
PROJECT_ROOT = Path(__file__).resolve().parents[1]

project_root_str = str(PROJECT_ROOT)
if project_root_str not in sys.path:
    sys.path.insert(0, project_root_str)

# ============================================================
# DEPENDENCIAS EXTERNAS
# ============================================================
import pandas as pd

# ============================================================
# MÓDULOS LOCALES DEL PROYECTO
# ============================================================
import healthcheck
import scheduler
import toa_proceso_mejorado as toa

# Carpeta temporal dentro de la raíz del proyecto.
TEST_TMP_DIR = PROJECT_ROOT / ".test_tmp"


def workspace_temp_file(name: str) -> Path:
    """Crea una ruta temporal estable dentro del proyecto."""
    TEST_TMP_DIR.mkdir(parents=True, exist_ok=True)

    path = TEST_TMP_DIR / name
    if path.exists():
        path.unlink()

    return path


class ToaTransformTests(unittest.TestCase):
    def test_api_schema_is_consistent(self) -> None:
        toa.validate_api_schema()
        self.assertEqual(len(toa.API_HEADERS), 77)

    def test_downloaded_xlsx_is_normalized_for_api(self) -> None:
        xlsx_path = workspace_temp_file("toa_test_input.xlsx")
        self.addCleanup(lambda: xlsx_path.exists() and xlsx_path.unlink())

        pd.DataFrame(
            {
                "Fecha": ["2026-06-23"],
                "Tipo de actividad": [" certificación "],
                "Lugar de la Reparación": ["Nodo"],
                "Número de OT": ["OT-1"],
            }
        ).to_excel(
            xlsx_path,
            index=False,
            engine="openpyxl",
        )

        dataframe = toa.consume_downloaded_xlsx(
            str(xlsx_path),
            "CENTRO",
            "ZONA CENTRO",
            toa.API_HEADERS,
        )
        records = toa.dataframe_to_api_records(dataframe)

        self.assertEqual(len(dataframe), 1)
        self.assertEqual(list(dataframe.columns), toa.API_HEADERS)
        self.assertEqual(len(records), 1)

        record = records[0]

        self.assertEqual(record["lugar_reparacion"], "Nodo")
        self.assertEqual(record["numero_ot"], "OT-1")
        self.assertEqual(record["subgrupo_toa"], "CENTRO")
        self.assertEqual(
            record["zona_subgrupo_toa"],
            "ZONA CENTRO",
        )

        # Comprueba que no existan NaN u otros valores inválidos para JSON.
        json.dumps(
            record,
            ensure_ascii=False,
            allow_nan=False,
        )


class SchedulerTests(unittest.TestCase):
    def test_next_aligned_time_uses_next_interval(self) -> None:
        current = datetime(
            2026,
            6,
            23,
            10,
            30,
            1,
            tzinfo=ZoneInfo("America/Santiago"),
        )

        target = scheduler.next_aligned_time(current, 30)

        self.assertEqual(target.hour, 11)
        self.assertEqual(target.minute, 0)
        self.assertEqual(target.second, 0)
        self.assertEqual(target.microsecond, 0)

    def test_lock_file_uses_fcntl_when_available(self) -> None:
        class FakeFcntl:
            LOCK_EX = 1
            LOCK_NB = 2
            LOCK_UN = 4

            def __init__(self) -> None:
                self.calls: list[tuple[int, int]] = []

            def flock(self, fileno: int, flags: int) -> None:
                self.calls.append((fileno, flags))

        class FakeHandle:
            def fileno(self) -> int:
                return 123

        fake_fcntl = FakeFcntl()
        fake_handle = FakeHandle()

        with patch.object(scheduler, "fcntl", fake_fcntl):
            scheduler.lock_file(fake_handle)
            scheduler.unlock_file(fake_handle)

        self.assertEqual(
            fake_fcntl.calls,
            [
                (
                    123,
                    fake_fcntl.LOCK_EX | fake_fcntl.LOCK_NB,
                ),
                (
                    123,
                    fake_fcntl.LOCK_UN,
                ),
            ],
        )


class HealthcheckTests(unittest.TestCase):
    def test_recent_state_is_healthy(self) -> None:
        state_path = workspace_temp_file(
            "scheduler_state_test.json"
        )
        self.addCleanup(
            lambda: state_path.exists() and state_path.unlink()
        )

        state_path.write_text(
            json.dumps(
                {
                    "scheduler_updated_at": datetime.now(
                        ZoneInfo("America/Santiago")
                    ).isoformat(),
                    "status": "IDLE",
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        with patch.dict(
            os.environ,
            {
                "SCHEDULER_STATE_FILE": str(state_path),
                "TZ": "America/Santiago",
                "HEALTHCHECK_MAX_AGE_MINUTES": "40",
            },
            clear=False,
        ):
            result = healthcheck.main()

        self.assertEqual(result, 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
