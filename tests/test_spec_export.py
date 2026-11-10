import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import app


class SessionState(dict):
    def __getattr__(self, name):
        return self[name]

    def __setattr__(self, name, value):
        self[name] = value


def make_order_row():
    extracted = app.ExtractedOrder(
        sender_explicitly_stated=False,
        sender="",
        sender_phone="",
        recipient="Test recipient",
        phone="01012345678",
        base_address="Test address",
        detail_address="",
        quantity_explicitly_stated=False,
        quantity=None,
        product_name="",
        evidence="Test order",
        review_reasons=[],
        uncertain_candidates=[],
    )
    return app._make_order_row(extracted, "test-group", "test-message")


class SpecExportTests(unittest.TestCase):
    def test_download_filename_uses_next_day_and_exported_box_total(self):
        for work_date, expected_date in (
            ("2026-11-10", "2026_11_11"),
            ("2026-11-30", "2026_12_01"),
            ("2026-12-31", "2027_01_01"),
        ):
            with self.subTest(work_date=work_date):
                state = SessionState(work={"work_date": work_date})
                with patch.object(app.st, "session_state", state):
                    self.assertEqual(
                        app._download_filename([{"수량": 20}, {"수량": 6}]),
                        expected_date + "_26박스.xlsx",
                    )
                    self.assertEqual(
                        app._download_filename([{"수량": 6}]),
                        expected_date + "_6박스.xlsx",
                    )

    def test_new_order_preview_places_default_spec_after_sender_phone(self):
        row = make_order_row()
        state = SessionState(work={"orders": [row]})
        with patch.object(app.st, "session_state", state):
            frame = app._editor_frame()
        columns = list(frame.columns)
        self.assertEqual(columns[columns.index("송하인전화") + 1], "규격")
        self.assertEqual(frame.loc[0, "규격"], "감귤10kg")

    def test_existing_approved_order_gets_default_without_losing_approval(self):
        row = make_order_row()
        row.pop("규격")
        row["approved"] = True
        work = app._empty_work("2026-11-10")
        work["orders"] = [row]
        state = SessionState(work=work, load_error=None)
        with patch.object(app.st, "session_state", state), patch.object(app, "_save_work") as save:
            app._initialize_state("2026-11-10")
        self.assertEqual(row["규격"], "감귤10kg")
        self.assertTrue(row["approved"])
        save.assert_called_once_with(work)

    def test_existing_custom_spec_is_preserved(self):
        row = make_order_row()
        row["규격"] = "5kg"
        work = app._empty_work("2026-11-10")
        work["orders"] = [row]
        state = SessionState(work=work, load_error=None)
        with patch.object(app.st, "session_state", state), patch.object(app, "_save_work"):
            app._initialize_state("2026-11-10")
        self.assertEqual(row["규격"], "5kg")

    def test_old_default_spec_is_upgraded(self):
        row = make_order_row()
        row["규격"] = "10kg"
        work = app._empty_work("2026-11-10")
        work["orders"] = [row]
        state = SessionState(work=work, load_error=None, orders_editor={"edited_rows": {}})
        with patch.object(app.st, "session_state", state), patch.object(app, "_save_work") as save:
            app._initialize_state("2026-11-10")
        self.assertEqual(row["규격"], "감귤10kg")
        self.assertNotIn("orders_editor", state)
        save.assert_called_once_with(work)

    def test_export_writes_default_and_custom_specs_to_h_with_template_style(self):
        with tempfile.TemporaryDirectory() as temporary:
            template = Path(temporary) / "template.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Sheet2"
            for column, title in enumerate(app.TEMPLATE_COLUMNS, start=2):
                sheet.cell(1, column, title)
            sheet["H2"] = "Old example"
            sheet["H2"].font = Font(bold=True, color="FF8800")
            sheet["H2"].alignment = Alignment(horizontal="center")
            sheet["H20"] = "Old example to remove"
            sheet.column_dimensions["D"].fill = PatternFill(fill_type="solid", fgColor="DDEEF3")
            sheet.column_dimensions["D"].border = Border(left=Side(style="thin"), right=Side(style="thin"))
            workbook.save(template)
            default_row = make_order_row()
            default_row.pop("규격")
            custom_row = make_order_row()
            custom_row["규격"] = "5kg"
            custom_row["수량"] = 3
            with patch.object(app, "TEMPLATE_FILE", template):
                result = app._build_template_excel([default_row, custom_row])
            exported = load_workbook(BytesIO(result))["Sheet2"]
        self.assertEqual(exported["H1"].value, "규격")
        self.assertEqual(exported["H2"].value, "감귤10kg")
        self.assertEqual(exported["H3"].value, "5kg")
        self.assertEqual(exported["G2"].value, default_row["송하인전화"])
        self.assertTrue(exported["H2"].font.bold)
        self.assertEqual(exported["H3"].alignment.horizontal, "center")
        self.assertEqual(exported.max_row, 4)
        self.assertEqual(exported["C4"].value, "총수량")
        self.assertEqual(exported["D4"].value, "=SUM(D2:D3)")
        self.assertEqual(exported["D4"].data_type, "f")
        self.assertTrue(exported["D4"].font.bold)
        self.assertEqual(exported["D4"].font.color.rgb, "00000000")
        self.assertEqual(exported["D4"].number_format, "0")
        self.assertTrue(all(exported.cell(4, column).fill.patternType is None for column in range(2, 9)))
        self.assertIsNone(exported.column_dimensions["D"].fill.patternType)
        for edge in ("left", "right", "top", "bottom"):
            self.assertIsNone(getattr(exported.column_dimensions["D"].border, edge))
            for column in range(2, 9):
                self.assertIsNone(getattr(exported.cell(4, column).border, edge))
        self.assertFalse(exported.row_dimensions[4].hidden)
        self.assertIsNone(exported["B4"].value)
        self.assertIsNone(exported["H4"].value)

    def test_total_formula_handles_empty_and_single_order_exports(self):
        with tempfile.TemporaryDirectory() as temporary:
            template = Path(temporary) / "template.xlsx"
            workbook = Workbook()
            workbook.active.title = "Sheet2"
            workbook.save(template)
            for rows, total_cell, formula in (
                ([], "D2", "=SUM(0)"),
                ([make_order_row()], "D3", "=SUM(D2:D2)"),
            ):
                with self.subTest(total_cell=total_cell):
                    with patch.object(app, "TEMPLATE_FILE", template):
                        exported = load_workbook(BytesIO(app._build_template_excel(rows)))
                    self.assertEqual(exported["Sheet2"][total_cell].value, formula)
                    self.assertEqual(exported.calculation.calcMode, "auto")
                    self.assertTrue(exported.calculation.fullCalcOnLoad)

    def test_preview_total_does_not_add_or_modify_orders(self):
        rows = [make_order_row(), make_order_row()]
        rows[1]["수량"] = 3
        self.assertEqual(app._box_total(rows), 4)
        self.assertEqual(len(rows), 2)
        self.assertEqual([row["수량"] for row in rows], [1, 3])
        self.assertEqual(app._box_total([]), 0)
        with patch.object(app.st, "session_state", SessionState(work={"orders": rows})):
            preview = app._editor_frame(include_total=True)
        self.assertEqual(len(preview), 3)
        self.assertEqual(preview.iloc[-1]["수하인전화"], "총수량")
        self.assertEqual(preview.iloc[-1]["수량"], 4)
        self.assertFalse(preview.iloc[-1]["검수완료"])
        self.assertFalse(preview.iloc[-1]["삭제 선택"])
        self.assertEqual(len(rows), 2)


if __name__ == "__main__":
    unittest.main()
