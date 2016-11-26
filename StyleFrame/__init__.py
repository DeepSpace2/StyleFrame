import sys

# Python 2
if sys.version_info < (3, 0):
    # noinspection PyUnresolvedReferences
    from style_frame import StyleFrame
    # noinspection PyUnresolvedReferences
    from container import Container
    # noinspection PyUnresolvedReferences
    from styler import Styler
    # noinspection PyUnresolvedReferences
    import utils
    # noinspection PyUnresolvedReferences,PyPackageRequirements
    from version import _version_

# Python 3
else:
    from StyleFrame.style_frame import StyleFrame
    from StyleFrame.container import Container
    from StyleFrame.styler import Styler
    from StyleFrame import utils
    from StyleFrame.version import _version_

from StyleFrame import warnings_conf

if 'utrunner' not in sys.argv[0]:
    from StyleFrame.tests import style_frame_tests as tests
