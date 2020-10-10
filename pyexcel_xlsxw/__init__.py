"""
    pyexcel_xlsxw
    ~~~~~~~~~~~~~~~~~~~

    The lower level xlsx file format writer using xlsxwriter

    :copyright: (c) 2016 by Onni Software Ltd & its contributors
    :license: New BSD License
"""
from pyexcel_io.io import get_data as read_data
from pyexcel_io.io import isstream
from pyexcel_io.io import store_data as write_data

# flake8: noqa
# this line has to be place above all else
# because of dynamic import
from pyexcel_io.plugins import IOPluginInfoChainV2

__FILE_TYPE__ = "xlsx"
IOPluginInfoChainV2(__name__).add_a_writer(
    relative_plugin_class_path="xlsxw.XLSXWriter",
    locations=["file", "memory"],
    file_types=[__FILE_TYPE__],
    stream_type="binary",
)


def save_data(afile, data, file_type=None, **keywords):
    """standalone module function for writing module supported file type"""
    if isstream(afile) and file_type is None:
        file_type = __FILE_TYPE__
    write_data(afile, data, file_type=file_type, **keywords)
