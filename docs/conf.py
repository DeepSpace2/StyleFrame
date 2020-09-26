import os
import sys

sys.path.insert(0, os.path.abspath('..'))

add_module_names = False
extensions = [
'sphinx.ext.autodoc',
]
html_theme = "sphinx_rtd_theme"
master_doc = 'index'
project = 'styleframe'
