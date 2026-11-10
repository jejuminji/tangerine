import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import app


class ReanalysisTests(unittest.TestCase):
    def test_identical_message_is_analyzed_again_and_adds_new_orders(self):
        work = app._empty_work("2026-11-10")
        order = app.ExtractedOrder(
            sender_explicitly_stated=False,
            sender="",
            sender_phone="",
            recipient="Test recipient",
            phone="01012345678",
            base_address="Test address",
            detail_address="",
            quantity_explicitly_stated=True,
            quantity=2,
            product_name="",
            evidence="Test order",
            review_reasons=[],
            uncertain_candidates=[],
        )
        response = SimpleNamespace(
            id="test-response",
            status="completed",
            output_parsed=app.ExtractedBatch(orders=[order]),
            usage=None,
        )
        client = MagicMock()
        client.responses.parse.return_value = response
        with (
            patch.object(app.st, "session_state", SimpleNamespace(work=work)),
            patch.object(app, "OpenAI", return_value=client),
            patch.object(app, "_save_work"),
            patch.object(app, "_load_order_prompt", return_value="Test prompt"),
            patch.dict(app.os.environ, {"OPENAI_API_KEY": "test-placeholder", "OPENAI_MODEL": "test-model"}),
        ):
            first_count, first_notice = app._analyze_message("Same order text")
            original_order = work["orders"][0].copy()
            second_count, second_notice = app._analyze_message("Same order text")
        self.assertEqual(client.responses.parse.call_count, 2)
        self.assertEqual((first_count, second_count), (1, 1))
        self.assertEqual((first_notice["kind"], second_notice["kind"]), ("success", "success"))
        self.assertEqual(len(work["messages"]), 2)
        self.assertEqual(len(work["orders"]), 2)
        self.assertEqual(work["orders"][0], original_order)
        self.assertNotEqual(work["messages"][0]["message_id"], work["messages"][1]["message_id"])
        self.assertNotEqual(work["orders"][0]["order_id"], work["orders"][1]["order_id"])


if __name__ == "__main__":
    unittest.main()
