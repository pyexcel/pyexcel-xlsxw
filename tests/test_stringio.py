import os

import pyexcel
from base import create_sample_file1
from nose.tools import eq_


class TestStringIO:
    def test_ods_stringio(self):
        odsfile = "cute.xlsx"
        create_sample_file1(odsfile)
        with open(odsfile, "rb") as f:
            content = f.read()
            r = pyexcel.get_sheet(file_type="xlsx", file_content=content)
            result = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", 1.1, 1]
            actual = list(r.enumerate())
            eq_(result, actual)
        if os.path.exists(odsfile):
            os.unlink(odsfile)

    def test_xls_output_stringio(self):
        data = [[1, 2, 3], [4, 5, 6]]
        io = pyexcel.save_as(dest_file_type="xlsx", array=data)
        r = pyexcel.get_sheet(file_type="xlsx", file_content=io.getvalue())
        result = [1, 2, 3, 4, 5, 6]
        actual = list(r.enumerate())
        eq_(result, actual)
