import pandas as pd


class Series(pd.Series):
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
