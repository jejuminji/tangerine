from concurrent.futures import wait
import threading
import time
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

import app


class AnalysisProgressTests(unittest.TestCase):
    def test_slow_request_keeps_reporting_on_main_thread(self):
        work = app._empty_work("2026-11-10")
        reports = []
        request_threads = []
        main_thread = threading.get_ident()

        def slow_response(**kwargs):
            request_threads.append(threading.get_ident())
            time.sleep(0.06)
            return SimpleNamespace(
                id="test-response", status="completed", usage=None,
                output_parsed=app.ExtractedBatch(orders=[]),
            )

        def progress(stage, elapsed):
            reports.append((stage, elapsed, threading.get_ident()))

        client = MagicMock()
        client.responses.parse.side_effect = slow_response
        with (
            patch.object(app.st, "session_state", SimpleNamespace(work=work)),
            patch.object(app, "OpenAI", return_value=client),
            patch.object(app, "_save_work"),
            patch.object(app, "_load_order_prompt", return_value="Test prompt"),
            patch.object(app, "wait", side_effect=lambda requests, timeout: wait(requests, timeout=0.005)),
            patch.dict(app.os.environ, {"OPENAI_API_KEY": "test-placeholder", "OPENAI_MODEL": "test-model"}),
            self.assertLogs(app.logger, level="INFO") as logs,
        ):
            count, notice = app._analyze_message("private input text", progress=progress)
        self.assertEqual(count, 0)
        self.assertEqual(notice["kind"], "warning")
        self.assertGreaterEqual(sum(stage == "AI 분석 응답 대기 중" for stage, _, _ in reports), 2)
        self.assertTrue(all(thread == main_thread for _, _, thread in reports))
        self.assertNotEqual(request_threads[0], main_thread)
        self.assertTrue(reports[-1][0].startswith("분석 완료"))
        self.assertGreater(reports[-1][1], 0)
        self.assertNotIn("private input text", "\n".join(logs.output))
        self.assertNotIn("test-placeholder", "\n".join(logs.output))

    def test_request_failure_is_reported_and_saved(self):
        work = app._empty_work("2026-11-10")
        reports = []
        client = MagicMock()
        client.responses.parse.side_effect = RuntimeError("simulated failure")
        with (
            patch.object(app.st, "session_state", SimpleNamespace(work=work)),
            patch.object(app, "OpenAI", return_value=client),
            patch.object(app, "_save_work") as save,
            patch.object(app, "_load_order_prompt", return_value="Test prompt"),
            patch.dict(app.os.environ, {"OPENAI_API_KEY": "test-placeholder", "OPENAI_MODEL": "test-model"}),
            self.assertLogs(app.logger, level="INFO"),
        ):
            count, notice = app._analyze_message("Test order", progress=lambda stage, elapsed: reports.append(stage))
        self.assertEqual((count, notice["kind"]), (0, "error"))
        self.assertEqual(reports[-1], "분석 실패: RuntimeError")
        self.assertEqual(work["messages"][0]["status"], "failed")
        self.assertEqual(save.call_count, 2)


if __name__ == "__main__":
    unittest.main()
