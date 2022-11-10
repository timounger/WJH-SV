# This Python file uses the following encoding: utf-8
"""
*****************************************************************************
 @file    generate_version_file.py
 @brief   WJH-SV - Utility Script to generate a version info .txt file for the executable
*****************************************************************************
"""

import sys
import os

from PyInstaller.utils.win32.versioninfo import *

sys.path.append('../')
from Source import wjh_sv as sv

versionInfo = VSVersionInfo(
    ffi=FixedFileInfo(
        # filevers and prodvers should be always a tuple with four items: (1, 2, 3, 4)
        # Set not needed items to zero 0.
        filevers=(sv.I_VERSION_NUM_1, sv.I_VERSION_NUM_2, sv.I_VERSION_NUM_3, sv.I_VERSION_NUM_4),
        prodvers=(sv.I_VERSION_NUM_1, sv.I_VERSION_NUM_2, sv.I_VERSION_NUM_3, sv.I_VERSION_NUM_4),
        # Contains a bitmask that specifies the Boolean attributes of the file
        mask=0x0,
        # Contains a bitmask that specifies the Boolean attributes of the file.
        flags=0x0,
        # THe operating system for which this flag was designed
        # 0x4 - NT and there no need to change it.
        OS=0x0,
        # The general type of file.
        # 0x1 - the function is not defined for this fileType
        subtype=0x0,
        # Creation date and time stamp.
        date=(0, 0)
        ),
    kids=[
        StringFileInfo(
            [
                StringTable(
                    u'040904E4',
                    [
                        StringStruct(u'FileDescription', sv.S_WJHSV_DESCRIPTION),
                        StringStruct(u'FileVersion', sv.S_VERSION),
                        StringStruct(u'LegalCopyright', sv.S_COPYRIGHT.replace("©", "(c)")),
                        StringStruct(u'ProductName', sv.S_WJHSV_APPLICATION_NAME),
                        StringStruct(u'ProductVersion', sv.S_VERSION)
                    ])
                ]),
        VarFileInfo([VarStruct(u'Translation', [1033, 1252])])
        ]
    )

if __name__ == "__main__":
    s_workpath = sys.argv[1]
    #s_workpath = "build"
    s_filename = "wjh_sv_version_info.txt"
    s_version_file = os.path.join(s_workpath, s_filename)
    print(f'Generate version file {s_version_file} (Version: {sv.S_VERSION})')
    os.mkdir(s_workpath) if not os.path.exists(s_workpath) else print(f'Directory {s_workpath} already exists')
    with open(s_version_file, "w") as version_file:
        version_file.write(str(versionInfo))
    sys.exit()
    