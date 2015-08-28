# coding:utf-8
import datetime as dt
import pandas as pd
from styler import Styler


class Container(object):
    """
    A container class used to store value and style pairs.
    Value can be any datatype, and style is a Styler object
    """
    def __init__(self, value, styler=None):
        self.value = value
        if styler is None:
            if isinstance(self.value, pd.tslib.Timestamp):
                self.style = Styler(number_format='DD/MM/YY HH:MM').create_style()
            elif isinstance(self.value, dt.date):
                self.style = Styler(number_format='DD/MM/YY').create_style()
            elif isinstance(self.value, dt.time):
                self.style = Styler(number_format='HH:MM').create_style()
            else:
                self.style = Styler().create_style()
        else:
            self.style = styler

    def __hash__(self):
        return hash(self.value)

    def __str__(self):
        return unicode(self.value)

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