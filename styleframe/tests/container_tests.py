import unittest

from styleframe import Container


class ContainerTest(unittest.TestCase):
    def setUp(self):
        self.cont_0 = Container(0)
        self.cont_1 = Container(1)
        self.cont_2 = Container(2)
        self.cont_str = Container('a string')
        self.cont_empty_str = Container('')
        self.cont_false = Container(False)
        self.cont_true = Container(True)

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

    def test__radd__(self):
        self.assertEqual(1 + self.cont_1, self.cont_2)

    def test__sub__(self):
        self.assertEqual(self.cont_2 - self.cont_1, self.cont_1)
        self.assertEqual(self.cont_2 - 1, self.cont_1)

    def test__rsub__(self):
        self.assertEqual(1 - self.cont_1, self.cont_0)

    def test__truediv__(self):
        self.assertEqual(self.cont_2 / self.cont_2, self.cont_1)
        self.assertEqual(self.cont_2 / 1, self.cont_2)

    def test__rtruediv__(self):
        self.assertEqual(1 / self.cont_1, self.cont_1)

    def test__floordiv__(self):
        self.assertEqual(self.cont_2 // 3, Container(2//3))
        self.assertEqual(self.cont_2 // Container(3), Container(2//3))

    def test__rfloordiv__(self):
        self.assertEqual(2 // self.cont_1, Container(2//1))

    def test__mul__(self):
        self.assertEqual(self.cont_1 * self.cont_1, self.cont_1)
        self.assertEqual(self.cont_2 * 1, self.cont_2)

    def test__rmul__(self):
        self.assertEqual(2 * self.cont_0, self.cont_0)
        self.assertEqual(2 * self.cont_1, self.cont_2)

    def test__mod__(self):
        self.assertEqual(self.cont_2 % self.cont_1, self.cont_0)
        self.assertEqual(self.cont_2 % 1, self.cont_0)

    def test__rmod__(self):
        self.assertEqual(1 % self.cont_2, self.cont_1)

    def test__pow__(self):
        self.assertEqual(self.cont_2 ** 2, Container(4))

    def test__int__(self):
        self.assertEqual(int(self.cont_2), 2)

    def test__float__(self):
        self.assertEqual(float(self.cont_1), 1.0)

    def test__len__(self):
        self.assertEqual(len(self.cont_empty_str), 0)
        self.assertEqual(len(self.cont_empty_str), len(self.cont_empty_str.value))
        self.assertEqual(len(self.cont_str), 8)
        self.assertEqual(len(self.cont_str), len(self.cont_str.value))

    def test__bool__(self):
        self.assertEqual(bool(self.cont_0), False)
        self.assertEqual(bool(self.cont_0), bool(self.cont_0.value))
        self.assertEqual(bool(self.cont_empty_str), False)
        self.assertEqual(bool(self.cont_empty_str), bool(self.cont_empty_str.value))
        self.assertEqual(bool(self.cont_false), False)
        self.assertEqual(bool(self.cont_false), bool(self.cont_false.value))
        self.assertEqual(bool(self.cont_1), True)
        self.assertEqual(bool(self.cont_1), bool(self.cont_1.value))
        self.assertEqual(bool(self.cont_true), True)
        self.assertEqual(bool(self.cont_true), bool(self.cont_true.value))
