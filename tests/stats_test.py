import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "backend"))

from app import _build_stats


class StatsTest(unittest.TestCase):
    def test_expected_and_confirmed_use_total_guest_numbers(self):
        guests = [
            {"party_size": 3, "confirmed_size": 2, "confirm_status": "已确认", "table_no": "1", "invite_status": "已发送", "wechat_sent": "已发送"},
            {"party_size": 2, "confirmed_size": 1, "confirm_status": "待确认", "table_no": "", "invite_status": "未发送", "wechat_sent": "未发送"},
            {"party_size": 4, "confirmed_size": 0, "confirm_status": "不参加", "table_no": "", "invite_status": "未发送", "wechat_sent": "未发送"},
        ]

        stats = _build_stats(guests, {"budget_total": 10})

        self.assertEqual(stats["expected"], 9)
        self.assertEqual(stats["confirmed"], 3)


if __name__ == "__main__":
    unittest.main()
