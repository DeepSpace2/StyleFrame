import unittest

from StyleFrame.command_line.tests.commandline_tests import CommandlineInterfaceTest
from StyleFrame.tests.container_tests import ContainerTest
from StyleFrame.tests.series_tests import SeriesTest
from StyleFrame.tests.style_frame_tests import StyleFrameTest
from StyleFrame.tests.styler_tests import StylerTests


def run():
    test_classes = [ContainerTest, StyleFrameTest, CommandlineInterfaceTest, SeriesTest, StylerTests]
    for test_class in test_classes:
        suite = unittest.TestLoader().loadTestsFromTestCase(test_class)
        unittest.TextTestRunner().run(suite)


if __name__ == '__main__':
    unittest.main()
