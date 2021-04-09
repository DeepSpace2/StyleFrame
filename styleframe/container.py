import datetime as dt
import pandas as pd

from . import utils

from styleframe.styler import Styler

try:
    pd_timestamp = pd.Timestamp
except AttributeError:
    pd_timestamp = pd.tslib.Timestamp


class Container:
    """
    A container class used to store value and style pairs.
    Value can be any datatype, and style is a Styler object
    """
    def __init__(self, value, styler=None):
        self.value = value
        if styler is None:
            if isinstance(self.value, pd_timestamp):
                self.style = Styler(number_format=utils.number_formats.default_date_time_format)
            elif isinstance(self.value, dt.date):
                self.style = Styler(number_format=utils.number_formats.default_date_format)
            elif isinstance(self.value, dt.time):
                self.style = Styler(number_format=utils.number_formats.default_time_format)
            else:
                self.style = Styler()
        else:
            self.style = styler

    def __hash__(self):
        return hash(self.value)

    def __str__(self):
        return str(self.value)

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return other.value == self.value
        return other == self.value

    def __ne__(self, other):
        if isinstance(other, self.__class__):
            return other.value != self.value
        return other != self.value

    def __gt__(self, other):
        if isinstance(other, self.__class__):
            return other.value < self.value
        return other < self.value

    def __ge__(self, other):
        if isinstance(other, self.__class__):
            return other.value <= self.value
        return other <= self.value

    def __lt__(self, other):
        if isinstance(other, self.__class__):
            return other.value > self.value
        return other > self.value

    def __le__(self, other):
        if isinstance(other, self.__class__):
            return other.value >= self.value
        return other >= self.value

    def __add__(self, other):
        if isinstance(other, self.__class__):
            return Container(self.value + other.value)
        return Container(self.value + other)

    def __radd__(self, other):
        return self.__add__(other)

    def __sub__(self, other):
        if isinstance(other, self.__class__):
            return Container(self.value - other.value)
        return Container(self.value - other)

    def __rsub__(self, other):
        if isinstance(other, self.__class__):
            return Container(other.value - self.value)
        return Container(other - self.value)

    def __truediv__(self, other):
        if isinstance(other, self.__class__):
            return Container(self.value / other.value)
        return Container(self.value / other)

    def __rtruediv__(self, other):
        if isinstance(other, self.__class__):
            return Container(other.value / self.value)
        return Container(other / self.value)

    def __floordiv__(self, other):
        if isinstance(other, self.__class__):
            return Container(self.value // other.value)
        return Container(self.value // other)

    def __rfloordiv__(self, other):
        if isinstance(other, self.__class__):
            return Container(other.value // self.value)
        return Container(other // self.value)

    def __mul__(self, other):
        if isinstance(other, self.__class__):
            return Container(self.value * other.value)
        return Container(self.value * other)

    def __rmul__(self, other):
        return self.__mul__(other)

    def __mod__(self, other):
        if isinstance(other, self.__class__):
            return Container(self.value % other.value)
        return Container(self.value % other)

    def __rmod__(self, other):
        if isinstance(other, self.__class__):
            return Container(other.value % self.value)
        return Container(other % self.value)

    def __pow__(self, power, modulo=None):
        return Container(self.value ** power)

    def __int__(self):
        return int(self.value)

    def __float__(self):
        return float(self.value)

    def __bool__(self):
        return bool(self.value)

    def __len__(self):
        return len(self.value)
