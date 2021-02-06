import os
import sys

sys.path.insert(0, os.path.abspath('..'))

add_module_names = False
extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.intersphinx',
]
html_theme = "furo"
intersphinx_mapping = {
    'python': ('https://docs.python.org/3', None),
    'pandas': ('https://pandas.pydata.org/docs/', None)
}
master_doc = 'index'
project = 'styleframe'
