import sys
from cx_Freeze import setup, Executable

buildOpt = {'excludes': ['email', 'asyncio', 'html', 'http', 'ctypes', 'multiprocessing', 'xml', 'test', 'unittest',
                         'pydoc_data', 'lib2to3', 'curses', 'msilib', 'logging', 'concurrent', 'pkg_resources', 'sqlite3',
                         'distutils', 'setuptools', '_distutils_hack']}

base = None

if sys.platform == 'win64':
    base = 'Win64GUI'
elif sys.platform == 'win32':
    base = 'Win32GUI'

setup(
    name='XChange',
    version='1.0',
    description='convert spread sheets',
    options={'build_exe': buildOpt},
    executables=[Executable('main.py', base=base)]
)
