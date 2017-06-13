"""
    pyexcel_xlsxw
    ~~~~~~~~~~~~~~~~~~~

    The lower level xlsx file format writer using xlsxwriter

    :copyright: (c) 2016 by Onni Software Ltd & its contributors
    :license: New BSD License
"""
import xlsxwriter

from pyexcel_io.book import BookWriter
from pyexcel_io.sheet import SheetWriter


class XLSXSheetWriter(SheetWriter):
    """
    xlsx sheet writer
    """
    def set_sheet_name(self, name):
        self.current_row = 0

    def write_row(self, array):
        """
        write a row into the file
        """
        for i in range(0, len(array)):
            self._native_sheet.write(self.current_row, i, array[i])
        self.current_row += 1


class XLSXWriter(BookWriter):
    """
    xlsx writer
    """
    def __init__(self):
        BookWriter.__init__(self)
        self._native_book = None

    def open(self, file_name, **keywords):
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
        keywords.setdefault('default_date_format', 'dd/mm/yy')
        keywords.setdefault('constant_memory', True)
        BookWriter.open(self, file_name, **keywords)

        self._native_book = xlsxwriter.Workbook(
            file_name, keywords
         )

    def create_sheet(self, name):
        return XLSXSheetWriter(self._native_book,
                               self._native_book.add_worksheet(name), name)

    def close(self):
        """
        This call actually save the file
        """
        self._native_book.close()
        self._native_book = None
