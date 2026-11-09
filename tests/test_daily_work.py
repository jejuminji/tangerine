import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app


class SessionState(dict):
    def __getattr__(self, name):
        return self[name]

    def __setattr__(self, name, value):
        self[name] = value


class DailyWorkTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        for name, value in (
            ("DATA_DIR", self.root),
            ("LEGACY_WORK_FILE", self.root / "current_work.json"),
        ):
            replacement = patch.object(app, name, value)
            replacement.start()
            self.addCleanup(replacement.stop)

    def test_dates_are_saved_and_reopened_independently(self):
        for day, text in (("2026-11-08", "yesterday"), ("2026-11-09", "today")):
            work = app._empty_work(day)
            work["messages"] = [{"raw_text": text}]
            app._save_work(work)
        yesterday, error = app._load_work("2026-11-08")
        self.assertIsNone(error)
        self.assertEqual(yesterday["messages"][0]["raw_text"], "yesterday")
        today, error = app._load_work("2026-11-09")
        self.assertIsNone(error)
        self.assertEqual(today["messages"][0]["raw_text"], "today")
        empty, error = app._load_work("2026-11-10")
        self.assertIsNone(error)
        self.assertEqual(empty["messages"], [])

    def test_save_after_midnight_keeps_original_work_date(self):
        work = app._empty_work("2026-11-08")
        with patch.object(app, "_now", return_value="2026-11-09T00:01:00+09:00"):
            app._save_work(work)
        self.assertTrue(app._work_file("2026-11-08").exists())
        self.assertFalse(app._work_file("2026-11-09").exists())

    def test_legacy_is_copied_by_last_updated_korean_date(self):
        legacy = app._empty_work("2026-11-08")
        legacy.pop("work_date")
        legacy["updated_at"] = "2026-11-08T16:00:00+00:00"
        contents = json.dumps(legacy).encode()
        app.LEGACY_WORK_FILE.write_bytes(contents)
        loaded, error = app._load_work("2026-11-09")
        self.assertIsNone(error)
        self.assertEqual(loaded["work_id"], legacy["work_id"])
        self.assertEqual(app.LEGACY_WORK_FILE.read_bytes(), contents)
        self.assertEqual(app._work_file("2026-11-09").read_bytes(), contents)

    def test_migration_never_overwrites_existing_daily_work(self):
        existing = app._empty_work("2026-11-09")
        app._save_work(existing)
        before = app._work_file("2026-11-09").read_bytes()
        legacy = app._empty_work("2026-11-09")
        legacy["updated_at"] = "2026-11-09T12:00:00+09:00"
        app.LEGACY_WORK_FILE.write_text(json.dumps(legacy), encoding="utf-8")
        app._migrate_legacy_work()
        self.assertEqual(app._work_file("2026-11-09").read_bytes(), before)

    def test_corrupt_daily_file_is_reported_and_preserved(self):
        target = app._work_file("2026-11-09")
        target.parent.mkdir()
        target.write_text("invalid json", encoding="utf-8")
        work, error = app._load_work("2026-11-09")
        self.assertIsNotNone(error)
        self.assertEqual(work["orders"], [])
        self.assertEqual(target.read_text(encoding="utf-8"), "invalid json")

    def test_switching_dates_reloads_work_and_clears_editor_state(self):
        earlier = app._empty_work("2026-11-08")
        earlier["messages"] = [{"message_id": "old", "raw_text": "earlier"}]
        app._save_work(earlier)
        state = SessionState(work=app._empty_work("2026-11-09"), orders_editor={"edited_rows": {0: {}}})
        with patch.object(app.st, "session_state", state):
            app._initialize_state("2026-11-08")
        self.assertEqual(state.work["work_id"], earlier["work_id"])
        self.assertNotIn("orders_editor", state)


if __name__ == "__main__":
    unittest.main()
