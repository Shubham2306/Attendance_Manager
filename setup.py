application_title = "Attendace Manager"
application_main_file = "bukcy.py"

import sys

from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"

includes = ["random", "functools", "openpyxl", "webbrowser"]

setup(
    name = application_title,
    version = "1.0.9",
    description = "Attendance Manager",
    options = {"build.exe" : {"includes" : includes}},
    icon = "icon.ico",
    executables = [Executable(application_main_file, base = base)]

)