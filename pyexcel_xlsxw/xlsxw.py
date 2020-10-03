"""
    pyexcel_xlsxw
    ~~~~~~~~~~~~~~~~~~~

    The lower level xlsx file format writer using xlsxwriter

    :copyright: (c) 2016 by Onni Software Ltd & its contributors
    :license: New BSD License
"""
import xlsxwriter
from pyexcel_io.plugin_api.abstract_sheet import ISheetWriter
from pyexcel_io.plugin_api.abstract_writer import IWriter


class XLSXSheetWriter(ISheetWriter):
    """
    xlsx sheet writer
    """

    def __init__(self, ods_book, ods_sheet, sheet_name, **_):
        self._native_book = ods_book
        self._native_sheet = ods_sheet
        self.current_row = 0

    def write_row(self, array):
        """
        write a row into the file
        """
        for i in range(0, len(array)):
            self._native_sheet.write(self.current_row, i, array[i])
        self.current_row += 1

    def close(self):
        pass


class XLSXWriter(IWriter):
    """
    xlsx writer
    """

    def __init__(
        self,
        file_alike_object,
        file_type,
        constant_memory=True,
        default_date_format="dd/mm/yy",
        **keywords
    ):
        """
        Open a file for writing

        Please note that this writer configure xlsxwriter's BookWriter to use
        constant_memory by default.

        :param keywords: **default_date_format** control the date time format.
                         **constant_memory** if true, reduces the memory
                         footprint when writing large files. Other parameters
                         can be found in `xlsxwriter's documentation
                         <http://xlsxwriter.readthedocs.io/workbook.html>`_
        """
        if "single_sheet_in_book" in keywords:
            keywords.pop("single_sheet_in_book")
        self._native_book = xlsxwriter.Workbook(
            file_alike_object, options=keywords
        )

    def create_sheet(self, name):
        return XLSXSheetWriter(
            self._native_book, self._native_book.add_worksheet(name), name
        )

    def close(self):
        """
        This call actually save the file
        """
        self._native_book.close()
        self._native_book = None
