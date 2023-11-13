import sys

from pandas import DataFrame

from .container import Container
from .series import Series
from .style_frame import StyleFrame
from .styler import Styler
from .command_line.commandline import CommandLineInterface
from .version import _version_, _versions_, _openpyxl_version_, _pandas_version_, _python_version_

from . import deprecations

if 'utrunner' not in sys.argv[0]:
    from styleframe.tests import tests

ExcelWriter = StyleFrame.ExcelWriter
read_excel = StyleFrame.read_excel

# applymap is deprecated in pandas > 2, nasty hack to support < 2 and > 2 versions at the same time
if not hasattr(DataFrame, 'map'):
    DataFrame.map = DataFrame.applymap
