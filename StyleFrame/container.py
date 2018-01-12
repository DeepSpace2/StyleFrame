# coding:utf-8
import sys
import datetime as dt
import pandas as pd

PY2 = sys.version_info[0] == 2

# Python 2
if PY2:
    from styler import Styler
# Python 3
else:
    from StyleFrame.styler import Styler

try:
    pd_timestamp = pd.Timestamp
except AttributeError:
    pd_timestamp = pd.tslib.Timestamp


class Container(object):
    """
    A container class used to store value and style pairs.
    Value can be any datatype, and style is a Styler object
    """
    def __init__(self, value, styler=None):
        self.value = value
        if styler is None:
            if isinstance(self.value, pd_timestamp):
                self.style = Styler(number_format='DD/MM/YY HH:MM')
            elif isinstance(self.value, dt.date):
                self.style = Styler(number_format='DD/MM/YY')
            elif isinstance(self.value, dt.time):
                self.style = Styler(number_format='HH:MM')
            else:
                self.style = Styler()
        else:
            self.style = styler

    def __hash__(self):
        return hash(self.value)

    def __str__(self):
        if PY2:
            return unicode(self.value)
        return str(self.value)

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return other.value == self.value
        else:
            return other == self.value

    def __gt__(self, other):
        if isinstance(other, self.__class__):
            return other.value < self.value
        else:
            return other < self.value

    def __ge__(self, other):
        if isinstance(other, self.__class__):
            return other.value <= self.value
        else:
            return other <= self.value

    def __lt__(self, other):
        if isinstance(other, self.__class__):
            return other.value > self.value
        else:
            return other > self.value

    def __le__(self, other):
        if isinstance(other, self.__class__):
            return other.value >= self.value
        else:
            return other >= self.value

    def __add__(self, other):
        if isinstance(other, self.__class__):
            return self.value + other.value
        return self.value + other

    def __sub__(self, other):
        if isinstance(other, self.__class__):
            return self.value - other.value
        return self.value - other

    def __div__(self, other):
        if isinstance(other, self.__class__):
            return self.value / other.value
        return self.value / other

    def __truediv__(self, other):
        if isinstance(other, self.__class__):
            return self.value / other.value
        return self.value / other

    def __floordiv__(self, other):
        if isinstance(other, self.__class__):
            return self.value // other.value
        return self.value // other

    def __mul__(self, other):
        if isinstance(other, self.__class__):
            return self.value * other.value
        return self.value * other

    def __mod__(self, other):
        if isinstance(other, self.__class__):
            return self.value % other.value
        return self.value % other

    def __pow__(self, power, modulo=None):
        return self.value ** power

    def __int__(self):
        return int(self.value)

    def __float__(self):
        return float(self.value)

    def __bool__(self):
        return bool(self.value)

    def __len__(self):
        return len(self.value)
