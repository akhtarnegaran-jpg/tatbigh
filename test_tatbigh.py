
# -*- coding: utf-8 -*-
import unittest
import pandas as pd
from main import normalize_fa, MatchConfig, run_matching

class TestTatbigh(unittest.TestCase):
    def test_persian_normalization(self):
        self.assertEqual(normalize_fa("  تبادكان  "), "تبادکان")
        self.assertEqual(normalize_fa("ناحیه‌ ۴"), "ناحیه 4")

    def test_exact_normalized_match(self):
        a = pd.DataFrame({"نام": ["كاشمر", "داورزن"], "کد": ["01", "02"]})
        b = pd.DataFrame({"منطقه": ["کاشمر"], "تعداد": ["7"]})
        result = run_matching(
            a, b, ["نام"], ["منطقه"],
            ["نام", "کد"], ["منطقه", "تعداد"],
            MatchConfig("normalized", 90, 4, False)
        )
        self.assertEqual(len(result["نتیجه نهایی"]), 1)
        self.assertEqual(len(result["فقط در فایل اول"]), 1)

if __name__ == "__main__":
    unittest.main()
