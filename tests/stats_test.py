import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "backend"))

from app import _build_stats


class StatsTest(unittest.TestCase):
    def test_expected_and_confirmed_use_total_guest_numbers(self):
        guests = [
            {"party_size": 3, "confirmed_size": 2, "table_no": "1", "invite_status": "已发送"},
            {"party_size": 2, "confirmed_size": 0, "table_no": "", "invite_status": "未发送"},
            {"party_size": 4, "confirmed_size": 0, "table_no": "", "invite_status": "未发送"},
        ]

        stats = _build_stats(guests, {"budget_total": 10})

        self.assertEqual(stats["expected"], 9)
        self.assertEqual(stats["confirmed"], 2)
        self.assertEqual(stats["seated"], 2)
        self.assertEqual(stats["unseated"], 6)


if __name__ == "__main__":
    unittest.main()
