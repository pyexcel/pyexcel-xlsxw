import os
import datetime

import pyexcel as pe


class TestDateFormat:
    def test_writing_date_format(self):
        excel_filename = "testdateformat.xlsx"
        data = [
            [
                datetime.date(2014, 12, 25),
                datetime.time(11, 11, 11),
                datetime.datetime(2014, 12, 25, 11, 11, 11),
            ]
        ]
        pe.save_as(dest_file_name=excel_filename, array=data)
        r = pe.get_sheet(file_name=excel_filename, library="pyexcel-xls")
        assert isinstance(r[0, 0], datetime.date) is True
        assert r[0, 0].strftime("%d/%m/%y") == "25/12/14"
        assert isinstance(r[0, 1], datetime.time) is True
        assert r[0, 1].strftime("%H:%M:%S") == "11:11:11"
        assert isinstance(r[0, 2], datetime.date) is True
        assert r[0, 2].strftime("%d/%m/%y %H:%M:%S") == "25/12/14 11:11:11"
        os.unlink(excel_filename)
