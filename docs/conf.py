import os
import sys

sys.path.insert(0, os.path.abspath('..'))

add_module_names = False

extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.intersphinx',
    'sphinx.ext.viewcode',
    'sphinx_copybutton'
]

html_theme = 'furo'

html_theme_options = {
    "source_repository": "https://github.com/DeepSpace2/styleframe/",
    "source_branch": "devel",
    "source_directory": "docs/",
    "dark_css_variables": {
        "color-api-name": "#2b8cee"
    },
    "light_css_variables": {
        "color-api-name": "#2962ff"
    }
}

intersphinx_mapping = {
    'python': ('https://docs.python.org/3', None),
    'pandas': ('https://pandas.pydata.org/docs/', None),
    'numpy': ('https://numpy.org/doc/stable/', None)
}

master_doc = 'index'

project = 'styleframe'
