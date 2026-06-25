from __future__ import annotations

import json
import os
import shutil
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


class FakeApiResponse:
    def __init__(self, status_code: int, payload: dict | None = None) -> None:
        self.status_code = status_code
        self._payload = payload or {"ok": True}
        self.text = json.dumps(self._payload, ensure_ascii=False)
        self.content = self.text.encode("utf-8")

    def json(self) -> dict:
        return self._payload


class FakeApiSession:
    def __init__(self, statuses: list[int]) -> None:
        self.statuses = statuses
        self.headers: dict[str, str] = {}
        self.calls: list[dict] = []
        self.closed = False

    def post(self, url: str, json=None, timeout=None):
        self.calls.append({"url": url, "json": json, "timeout": timeout})
        status = self.statuses.pop(0)
        payload = {"ok": 200 <= status < 300, "status": status}
        return FakeApiResponse(status, payload)

    def close(self) -> None:
        self.closed = True


class ToaApiPublicationTests(unittest.TestCase):
    def make_records(self, total: int) -> list[dict]:
        return [
            {key: f"{index}-{key}" for key in toa.API_HEADERS}
            for index in range(total)
        ]

    def api_env(self) -> dict[str, str]:
        return {
            "TOA_API_EMPTY_URL": "http://api.local/vaciar",
            "TOA_API_CREATE_URL": "http://api.local/crear",
            "TOA_API_BATCH_SIZE": "300",
            "TOA_API_TIMEOUT": "30",
            "TOA_API_BATCH_DELAY_SECONDS": "0",
            "TOA_API_EMPTY_RETRIES": "1",
            "TOA_API_CREATE_RETRIES": "1",
            "TOA_API_RETRY_WAIT_SECONDS": "0",
            "TOA_API_DRY_RUN": "false",
            "TOA_API_PAYLOAD_KEY": "",
        }

    def test_api_publication_empties_then_sends_batches(self) -> None:
        failure_dir = TEST_TMP_DIR / "api_success"
        self.addCleanup(lambda: shutil.rmtree(failure_dir, ignore_errors=True))
        fake_session = FakeApiSession([200, 201, 201, 201])

        with patch.dict(os.environ, self.api_env(), clear=False):
            with patch.object(toa.requests, "Session", lambda: fake_session):
                result = toa.publicar_toa_api_por_lotes(
                    self.make_records(650),
                    str(failure_dir),
                )

        self.assertTrue(fake_session.closed)
        self.assertEqual([call["url"] for call in fake_session.calls], [
            "http://api.local/vaciar",
            "http://api.local/crear",
            "http://api.local/crear",
            "http://api.local/crear",
        ])
        self.assertEqual(
            [len(call["json"]) for call in fake_session.calls[1:]],
            [300, 300, 50],
        )
        self.assertEqual(result["registros_enviados"], 650)
        self.assertEqual(result["lotes_procesados"], 3)
        self.assertEqual(result["lotes_fallidos"], 0)
        self.assertEqual(result["estado_final"], "EXITOSO")

    def test_api_publication_stops_on_failed_batch(self) -> None:
        failure_dir = TEST_TMP_DIR / "api_failure"
        self.addCleanup(lambda: shutil.rmtree(failure_dir, ignore_errors=True))
        fake_session = FakeApiSession([200, 201, 500])

        with patch.dict(os.environ, self.api_env(), clear=False):
            with patch.object(toa.requests, "Session", lambda: fake_session):
                with self.assertRaises(toa.ApiPublicationError) as captured:
                    toa.publicar_toa_api_por_lotes(
                        self.make_records(650),
                        str(failure_dir),
                    )

        summary = captured.exception.summary
        self.assertEqual(len(fake_session.calls), 3)
        self.assertEqual(summary["registros_enviados"], 300)
        self.assertEqual(summary["lotes_procesados"], 2)
        self.assertEqual(summary["lotes_fallidos"], 1)
        self.assertEqual(summary["estado_final"], "ERROR_LOTE")
        self.assertEqual(summary["errores"][0]["lote"], 2)
        self.assertEqual(summary["errores"][0]["rango_registros"], [301, 600])
        self.assertTrue((failure_dir / "lote_api_fallido_0002.json").is_file())


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
