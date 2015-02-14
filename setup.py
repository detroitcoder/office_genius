# setup_interp.py
# A distutils setup script for the "interp" sample.

from distutils.core import setup
import py2exe

setup(
    options={"py2exe": {"packages":"encodings",
                        "excludes": ["pywin", "pywin.dialogs", "pywin.dialogs.list", "win32ui"],
                        "dll_excludes": ['w9xpopen.exe'],
                        "bundle_files": 3}},
    com_server=["basic_excel_object"],
    zipfile=None
    )
