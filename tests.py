import unittest
from unittest import TestCase
from run import generate_rrule_pattern


class GenerateRrulePatternTests(TestCase):
    def test_weekly_sunday(self):
        real_pattern = "RRULE:FREQ=WEEKLY;BYDAY=SU;INTERVAL=1"
        pattern = generate_rrule_pattern("U")
        self.assertEqual(pattern, real_pattern)

    def test_weekly_sunday_and_monday(self):
        real_pattern = "RRULE:FREQ=WEEKLY;BYDAY=SU,MO;INTERVAL=1"
        pattern = generate_rrule_pattern("UM")
        self.assertEqual(pattern, real_pattern)
        pattern = generate_rrule_pattern("MU")
        self.assertEqual(pattern, real_pattern)

    def test_empty_freq_code(self):
        pattern = generate_rrule_pattern("")
        self.assertEqual(pattern, "")

    def test_monthly(self):
        real_pattern = "RRULE:FREQ=MONTHLY"
        pattern = generate_rrule_pattern("MONTHLY")
        self.assertEqual(pattern, real_pattern)


if __name__ == "__main__":
    unittest.main()
