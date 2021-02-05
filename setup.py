import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"
if sys.platform == "win64":
    base = "Win64GUI"
    
setup(
    name="FAS",
    author="Hugo Everaldo Salvador Bezerra",
    version="2.0.13",
    description="Ferramenta de Automatização para Projetos de Sistemas Supervisórios",
    options={'build_exe': {
        'includes': ["gi", 'xml', 'datetime', 'os', 'bs4','sys', 'traceback', 'openpyxl', 're'
                     'xlsxwriter', 'difflib', 'operator', 'pickle', 'threading'],
        'excludes': ['wx', 'email', 'pydoc_data', 'curses'],
        'packages': ["gi", 'xml', 'os'],
        'include_files': includeFiles
    }},
    executables=[
        Executable("FASgtkui.py",
                   base=base
                   )
    ]
)
