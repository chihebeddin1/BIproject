# create_exe.py - Create standalone executable
import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning
build_exe_options = {
    "packages": ["os", "sys", "PyQt5", "pandas", "pyodbc", "matplotlib", "numpy"],
    "excludes": ["tkinter"],
    "include_files": []
}

# GUI applications require a different base on Windows
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="DataWarehouseDashboard",
    version="1.0",
    description="Data Warehouse Dashboard",
    options={"build_exe": build_exe_options},
    executables=[Executable("dashboard_desktop.py", base=base)]
)