from __future__ import annotations

from datetime import date
import sys
from pathlib import Path
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from lib.financial_year import get_financial_year


class FinancialYearTests(unittest.TestCase):
    def test_financial_year_rollover(self) -> None:
        self.assertEqual(get_financial_year(date(2026, 3, 31)), "FY2025-2026")
        self.assertEqual(get_financial_year(date(2026, 4, 1)), "FY2026-2027")


if __name__ == "__main__":
    unittest.main()
