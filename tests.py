import unittest
from unittest import TestCase
from pandas import Series
from run import generate_rrule_pattern, ControlAction


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


class MissingEntriesTests(TestCase):
    """
    mandatory_entries: "list[str]" = ['Name', 'Duration',
                                      'Time', 'Start Date']
    """
    func = ControlAction._mandatory_entries_missing

    def test_not_missing(self):
        series = Series({"Name": "Yusuf",
                         "Duration": "0:05",
                         "Time": "10pm",
                         "Start Date": "Not Today"
                         })
        self.assertEqual(self.func(series), False)

    def test_name_missing(self):
        series = Series({"Name": "",
                         "Duration": "0:05",
                         "Time": "10pm",
                         "Start Date": "Not Today"
                         })
        self.assertEqual(self.func(series), True)


if __name__ == "__main__":
    unittest.main()
