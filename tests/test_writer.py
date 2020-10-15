import os

from base import PyexcelWriterBase, PyexcelHatWriterBase
from pyexcel_xls import get_data
from pyexcel_xlsxw import xlsxw as xlsx

from nose.tools import eq_


class TestNativeXLWriter:
    def test_write_book(self):
        self.content = {
            "Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]],
            "Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]],
            "Sheet3": [[u"X", u"Y", u"Z"], [1, 4, 7], [2, 5, 8], [3, 6, 9]],
        }
        self.testfile = "xlwriter.xlsx"
        writer = xlsx.XLSXWriter(self.testfile, "xlsx")
        writer.write(self.content)
        writer.close()
        content = get_data(self.testfile)
        for key in content.keys():
            content[key] = list(content[key])
        eq_(content, self.content)

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)


class TestXLSnCSVWriter(PyexcelWriterBase):
    def setUp(self):
        self.testfile = "test.xlsx"
        self.testfile2 = "test.csv"

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
        if os.path.exists(self.testfile2):
            os.unlink(self.testfile2)


class TestXLSHatWriter(PyexcelHatWriterBase):
    def setUp(self):
        self.testfile = "test.xlsx"

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
