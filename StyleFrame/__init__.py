import sys

# Python 2
if sys.version_info < (3, 0):
    # noinspection PyUnresolvedReferences
    from style_frame import StyleFrame
    # noinspection PyUnresolvedReferences
    from container import Container
    # noinspection PyUnresolvedReferences
    from styler import Styler, number_formats, colors
    # noinspection PyUnresolvedReferences,PyPackageRequirements
    from version import _version_

# Python 3
else:
    from StyleFrame.style_frame import StyleFrame
    from StyleFrame.container import Container
    from StyleFrame.styler import Styler, number_formats, colors
    from StyleFrame.version import _version_

from StyleFrame.tests import style_frame_tests as tests
