import sys

# Python 2
if sys.version_info < (3, 0):
    from style_frame import StyleFrame
    from container import Container
    from styler import Styler, number_formats, colors
    from version import _version_

# Python 3
else:
    from StyleFrame.style_frame import StyleFrame
    from StyleFrame.container import Container
    from StyleFrame.styler import Styler, number_formats, colors
    from StyleFrame.version import _version_


