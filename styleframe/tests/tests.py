import unittest

from styleframe.command_line.tests.commandline_tests import CommandlineInterfaceTest
from styleframe.tests.container_tests import ContainerTest
from styleframe.tests.series_tests import SeriesTest
from styleframe.tests.style_frame_tests import StyleFrameTest
from styleframe.tests.styler_tests import StylerTests


def run():
    test_classes = [ContainerTest, StyleFrameTest, CommandlineInterfaceTest, SeriesTest, StylerTests]
    for test_class in test_classes:
        suite = unittest.TestLoader().loadTestsFromTestCase(test_class)
        unittest.TextTestRunner().run(suite)


if __name__ == '__main__':
    unittest.main()
