import pandas as pd

from .styler import Styler


class Series(pd.Series):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # enabling the styler accessor (for now only usable using .loc), for example:
        #         sf.loc[sf['col_name'].style.bg_color == utils.colors.yellow]
        #         sf.loc[~sf['col_name'].style.bold]
        for attr in Styler().__dict__:
            # a dirty hack to avoid hard-coding all of Styler's attributes
            setattr(self, attr, pd.Series(getattr(i, attr) for i in self if isinstance(i, Styler)))

    def isnull(self):
        return pd.Series(i.value for i in self).isnull()

    def notnull(self):
        return pd.Series(i.value for i in self).notnull()

    @property
    def dt(self):
        return pd.Series(i.value for i in self).dt

    @property
    def str(self):
        return pd.Series(i.value for i in self).str

    @property
    def style(self):
        return Series(i.style for i in self)

    @style.setter
    def style(self, value):
        for v in self:
            v.style = value
