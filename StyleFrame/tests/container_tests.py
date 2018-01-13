import unittest

from StyleFrame import Container


class ContainerTest(unittest.TestCase):
    def setUp(self):
        self.cont_1 = Container(1)
        self.cont_2 = Container(2)

    def test__gt__(self):
        self.assertGreater(self.cont_2, self.cont_1)
        self.assertGreater(self.cont_2, 1)

    def test__ge__(self):
        self.assertGreaterEqual(self.cont_1, self.cont_1)
        self.assertGreaterEqual(self.cont_2, self.cont_1)
        self.assertFalse(self.cont_1 >= self.cont_2)
        self.assertFalse(self.cont_1 >= 3)

    def test__lt__(self):
        self.assertLess(self.cont_1, self.cont_2)
        self.assertLess(self.cont_2, 3)
        self.assertFalse(self.cont_2 < self.cont_1)

    def test__le__(self):
        self.assertLessEqual(self.cont_1, self.cont_1)
        self.assertLessEqual(self.cont_1, self.cont_2)
        self.assertLessEqual(self.cont_1, 3)
        self.assertFalse(self.cont_2 < self.cont_1)

    def test__add__(self):
        self.assertEqual(self.cont_1 + self.cont_1, self.cont_2)
        self.assertEqual(self.cont_1 + 1, self.cont_2)

    def test__sub__(self):
        self.assertEqual(self.cont_2 - self.cont_1, self.cont_1)
        self.assertEqual(self.cont_2 - 1, self.cont_1)

    def test__div__(self):
        self.assertEqual(self.cont_2 / self.cont_2, self.cont_1)
        self.assertEqual(self.cont_2 / self.cont_1, self.cont_2)
        self.assertEqual(self.cont_2 / 3, Container(2/3))

    def test__mul__(self):
        self.assertEqual(self.cont_1 * self.cont_1, self.cont_1)
        self.assertEqual(self.cont_2 * 1, self.cont_2)

    def test__mod__(self):
        self.assertEqual(self.cont_2 % self.cont_1, Container(0))
        self.assertEqual(self.cont_2 % 1, Container(0))

    def test__pow__(self):
        self.assertEqual(self.cont_2 ** 2, Container(4))

    def test__int__(self):
        self.assertEqual(int(self.cont_2), 2)

    def test__float__(self):
        self.assertEqual(float(self.cont_1), 1.0)
