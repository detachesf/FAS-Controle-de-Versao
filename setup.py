import sys
from cx_Freeze import setup, Executable
from Dlls import include_files

base = None
if sys.platform == "win32":
    base = "Win32GUI"
if sys.platform == "win64":
    base = "Win64GUI"

target = Executable(
    script="FASgtkui.py",
    base=base,
    icon="static/CHESF1.ico",
    targetName='FASgtkui.exe',
    shortcutName="FAS",
    shortcutDir="DesktopFolder"
    )
setup(
    name="FAS",
    author="Hugo Everaldo Salvador Bezerra",
    version="2.1.1",
    description="Ferramenta de Automatização para Projetos de Sistemas Supervisórios",
    options={'build_exe': {
        'includes': ["gi", 'xml', 'datetime', 'os', 'bs4','sys', 'traceback', 'openpyxl', 're',
                     'xlsxwriter', 'difflib', 'operator', 'pickle', 'threading'],
        'excludes': ['wx', 'email', 'pydoc_data', 'curses'],
        'packages': ["gi", 'xml'],
        'include_files': include_files
    },
    'bdist_msi': {'initial_target_dir': 'C:\\FAS_2.1.1'}},
    executables=[target]
)
