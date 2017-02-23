"""

  This file keeps all fixes for issues found

"""
import os
import datetime
from textwrap import dedent
from unittest import TestCase
import pyexcel as pe

from pyexcel_xlsxw import save_data


class TestBugFix(TestCase):

    def test_pyexcel_issue_5(self):
        """pyexcel issue #5

        datetime is not properly parsed
        """
        s = pe.load(os.path.join("tests",
                                 "test-fixtures",
                                 "test-date-format.xls"))
        s.save_as("issue5.xlsx")
        s2 = pe.load("issue5.xlsx")
        assert s[0, 0] == datetime.datetime(2015, 11, 11, 11, 12, 0)
        assert s2[0, 0] == datetime.datetime(2015, 11, 11, 11, 12, 0)

    def test_pyexcel_issue_8_with_physical_file(self):
        """pyexcel issue #8

        formular got lost
        """
        tmp_file = "issue_8_save_as.xlsx"
        s = pe.load(os.path.join("tests",
                                 "test-fixtures",
                                 "test8.xlsx"))
        s.save_as(tmp_file)
        s2 = pe.load(tmp_file)
        self.assertEqual(str(s), str(s2))
        content = dedent("""
        CNY:
        +----------+----------+------+---+-------+
        | 01/09/13 | 02/09/13 | 1000 | 5 | 13.89 |
        +----------+----------+------+---+-------+
        | 02/09/13 | 03/09/13 | 2000 | 6 | 33.33 |
        +----------+----------+------+---+-------+
        | 03/09/13 | 04/09/13 | 3000 | 7 | 58.33 |
        +----------+----------+------+---+-------+""").strip("\n")
        self.assertEqual(str(s2), content)
        os.unlink(tmp_file)

    def test_pyexcel_issue_8_with_memory_file(self):
        """pyexcel issue #8

        formular got lost
        """
        tmp_file = "issue_8_save_as.xlsx"
        f = open(os.path.join("tests",
                              "test-fixtures",
                              "test8.xlsx"),
                 "rb")
        s = pe.load_from_memory('xlsx', f.read())
        s.save_as(tmp_file)
        s2 = pe.load(tmp_file)
        self.assertEqual(str(s), str(s2))
        content = dedent("""
        CNY:
        +----------+----------+------+---+-------+
        | 01/09/13 | 02/09/13 | 1000 | 5 | 13.89 |
        +----------+----------+------+---+-------+
        | 02/09/13 | 03/09/13 | 2000 | 6 | 33.33 |
        +----------+----------+------+---+-------+
        | 03/09/13 | 04/09/13 | 3000 | 7 | 58.33 |
        +----------+----------+------+---+-------+""").strip("\n")
        self.assertEqual(str(s2), content)
        os.unlink(tmp_file)

    def test_excessive_columns(self):
        tmp_file = "date_field.xlsx"
        s = pe.get_sheet(file_name=os.path.join("tests", "fixtures", tmp_file))
        assert s.number_of_columns() == 2

    def test_workbook_options(self):
        cell_content = "= Hello World ="
        tmp_file = "workbook_options.xlsx"
        options = {'strings_to_formulas': False}
        data = {"Sheet 1": [[cell_content]]}
        save_data(tmp_file, data, options=options)
        sheet = pe.get_sheet(file_name=tmp_file)
        self.assertEqual(sheet[0][0], cell_content)
        os.unlink(tmp_file)
