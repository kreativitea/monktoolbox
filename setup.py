# setup.py
from distutils.core import setup
import py2exe

setup(console = ["monktoolbox.py"],
      options = {"py2exe":{"excludes": ["win32com.gen_py"]}})
