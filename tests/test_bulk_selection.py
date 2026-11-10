import unittest
from unittest.mock import patch

import app


class SessionState(dict):
    def __getattr__(self, name):
        return self[name]

    def __setattr__(self, name, value):
        self[name] = value


class BulkSelectionTests(unittest.TestCase):
    def setUp(self):
        work = app._empty_work("2026-11-10")
        for index in range(2):
            row = dict(zip(app.TEMPLATE_COLUMNS, ["Recipient", "01012345678", 1, "Address", "Sender", "01087654321", "10kg"]))
            row.update(order_id=str(index), approved=False)
            app._refresh_row_status(row)
            work["orders"].append(row)
        self.state = SessionState(work=work, delete_order_ids=set(), bulk_delete=False, bulk_approval=False)
        replacement = patch.object(app.st, "session_state", self.state)
        replacement.start()
        self.addCleanup(replacement.stop)

    def test_bulk_delete_selects_and_clears_orders_without_selecting_total(self):
        self.state.bulk_delete = True
        app._toggle_all_delete()
        frame = app._editor_frame(include_total=True)
        self.assertTrue(frame.iloc[:2]["삭제 선택"].all())
        self.assertFalse(frame.iloc[-1]["삭제 선택"])
        self.assertEqual(len(self.state.work["orders"]), 2)
        self.state.bulk_delete = False
        app._toggle_all_delete()
        self.assertFalse(app._editor_frame()["삭제 선택"].any())

    def test_bulk_approval_respects_required_fields_and_preserves_delete_selection(self):
        self.state.delete_order_ids = {"0"}
        self.state.work["orders"][1]["수하인전화"] = ""
        with patch.object(app, "_save_work"):
            self.state.bulk_approval = True
            app._toggle_all_approval()
            self.assertTrue(self.state.work["orders"][0]["approved"])
            self.assertFalse(self.state.work["orders"][1]["approved"])
            self.assertEqual(self.state.notice["kind"], "error")
            self.assertEqual(self.state.delete_order_ids, {"0"})
            self.state.bulk_approval = False
            app._toggle_all_approval()
        self.assertFalse(any(row["approved"] for row in self.state.work["orders"]))

    def test_individual_delete_selection_is_kept_after_editor_refresh(self):
        edited = app._editor_frame()
        edited.loc[1, "삭제 선택"] = True
        app._apply_editor_changes(edited, rerun=False)
        self.assertEqual(self.state.delete_order_ids, {"1"})
        self.assertTrue(app._editor_frame().loc[1, "삭제 선택"])


if __name__ == "__main__":
    unittest.main()
