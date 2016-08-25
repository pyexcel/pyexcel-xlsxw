"""
    pyexcel_xlsxw
    ~~~~~~~~~~~~~~~~~~~

    The lower level xlsx file format writer using xlsxwriter

    :copyright: (c) 2016 by Onni Software Ltd & its contributors
    :license: New BSD License
"""
# flake8: noqa
# this line has to be place above all else
# because of dynamic import
_FILE_TYPE = 'xlsx'
__pyexcel_io_plugins__ = [_FILE_TYPE]


from pyexcel_io.io import isstream, store_data as write_data


def save_data(afile, data, file_type=None, **keywords):
    """standalone module function for writing module supported file type"""
    if isstream(afile) and file_type is None:
        file_type = _FILE_TYPE
    write_data(afile, data, file_type=file_type, **keywords)
