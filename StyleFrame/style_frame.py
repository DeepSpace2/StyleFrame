# coding:utf-8
import pandas as pd
from container import Container
from styler import Styler
import numpy as np

class StyleFrame(object):
    """
    A wrapper class that wraps pandas DataFrame.
    Stores container objects that have values and Styles that will be applied to excel
    """
    def __init__(self, obj):
        if isinstance(obj, pd.DataFrame):
            self.data_df = obj.applymap(lambda x: Container(x))
        elif isinstance(obj, pd.Series):
            self.data_df = obj.apply(lambda x: Container(x) if not isinstance(x, Container) else x.value)
        elif isinstance(obj, dict) or isinstance(obj, list):
            self.data_df = pd.DataFrame(obj).applymap(lambda x: Container(x))
        else:
            raise TypeError("{} __init__ doesn't support {}".format(type(self).__name__, type(obj).__name__))
        self.data_df.columns = [Container(col) for col in self.data_df.columns]

    def __str__(self):
        return str(self.data_df)

    def __unicode__(self):
        return unicode(self.data_df)

    def __len__(self):
        return len(self.data_df)

    def __getitem__(self, item):
        if type(item) == pd.Series:
            return self.data_df.__getitem__(item).index
        return self.data_df.__getitem__(item)

    def __setitem__(self, key, value):
        self.data_df.__setitem__(key, value)

    def __delitem__(self, item):
        return self.data_df.__delitem__(item)

    def __getattr__(self, attr):
        known_attrs = {'ix': self.data_df.ix,
                       'applymap': self.data_df.applymap,
                       'groupby': self.data_df.groupby,
                       'index': self.data_df.index,
                       'columns': self.data_df.columns,
                       'fillna': self.data_df.fillna}
        if attr in known_attrs and hasattr(self.data_df, attr):
            return known_attrs[attr]
        else:
            raise AttributeError("'{}' object has no attribute '{}'".format(type(self).__name__, attr))

    @classmethod
    def read_excel(cls, path, sheetname, **kwargs):
        return StyleFrame(pd.read_excel(path, sheetname=sheetname, **kwargs))

    @classmethod
    def ExcelWriter(cls, path):
        return pd.ExcelWriter(path, engine='openpyxl')

    def to_excel(self, excel_writer, sheet_name='Sheet1', na_rep='', float_format=None, columns=None, header=True, index=False,
                 index_label=None, startrow=0, startcol=0, engine='openpyxl', merge_cells=True, encoding=None, inf_rep='inf',
                 right_to_left=True, columns_to_hide=None):
        """
        Saves the dataframe to excel and applies the styles.
        :param right_to_left: sets the sheet to be right to left.
        :param columns_to_hide: list or tuple of columns to hide, may be column index (starts from 1)
            column name or column letter.
        Read Pandas' documentation about the other parameters
        """
        if index:
            raise ValueError("'index' must be set to False")

        def get_values(x):
            if isinstance(x, Container):
                return x.value
            else:
                try:
                    if np.isnan(x):
                        return na_rep
                    else:
                        return x
                except TypeError:
                    return x

        export_df = self.data_df.applymap(lambda x: get_values(x))
        export_df.columns = [col.value for col in export_df.columns]

        export_df.to_excel(excel_writer, sheet_name=sheet_name, na_rep=na_rep, float_format=float_format, columns=columns,
                           header=header, index=index, index_label=index_label, startrow=startrow, startcol=startcol,
                           engine=engine, merge_cells=merge_cells, encoding=encoding, inf_rep=inf_rep)

        sheet = excel_writer.book.get_sheet_by_name(sheet_name)

        sheet.sheet_view.rightToLeft = right_to_left

        self.data_df.fillna(Container('NaN'), inplace=True)

        ''' Iterating over the dataframe's elements and applying their styles '''
        ''' openpyxl's rows and cols start from 1,1 while the dataframe is 0,0 '''
        for col_index, column in enumerate(self.data_df.columns):
            current_fgColor = self.data_df.columns[col_index].style.fill.fgColor.rgb
            current_size = self.data_df.columns[col_index].style.font.size
            sheet.cell(row=startrow + 1, column=col_index + startcol + 1).style = Styler(color=current_fgColor, bold=True, size=current_size).create_style()
            for row_index, index in enumerate(self.data_df.index):
                sheet.cell(row=row_index + startrow + 2, column=col_index + startcol + 1).style = self.data_df.ix[index, column].style
                
        ''' Iterating over the columns_to_hide and check if the format is columns name, column index as number or letter  '''
        if columns_to_hide is not None:
            if not isinstance(columns_to_hide, (list, tuple)):
                raise TypeError("columns_to_hide must be a list or a tuple")
            
            for column in columns_to_hide:
                if not isinstance(column,(int, str)):
                    raise TypeError("column must be an index, column letter or column name")

                column_as_letter = None
                if column in self.data_df.columns:  # column name
                    column_index = self.data_df.columns.get_loc(column) + startcol + 1  # worksheet columns index start from 1
                    column_as_letter = get_column_letter(column_index)
                elif isinstance(column, int) and column >= 1:  # column index
                    column_as_letter = get_column_letter(column)
                elif column in sheet.column_dimensions:  # column letter
                    column_as_letter = column

                if column_as_letter is None or column_as_letter not in sheet.column_dimensions:
                    raise TypeError("column: %s is out of columns range." % column)

                sheet.column_dimensions[column_as_letter].hidden = True

    def apply_style_by_indexes(self, indexes_to_style=None, cols_name=None, color='white', bold=False, size=12, number_format='General'):
        """
        applies a certain style to the provided indexes in the dataframe in the provided columns
        :param indexes_to_style: indexes to apply the style to
        :param cols_name: the columns to apply the style to, if not provided all the columns will be styled
        :param color: the color to use
        :param bold: bold or not
        :param size: the font size
        :param number_format: style the number format
        :return:
        """
        if cols_name is not None and type(cols_name) not in [list, tuple]:
            raise TypeError("cols_name must be a list or a tuple")
        elif cols_name is None:
            cols_name = list(self.data_df.columns)
        for index in indexes_to_style:
            for col in cols_name:
                self.ix[index, col].style = Styler(color=color, bold=bold, size=size, number_format=number_format).create_style()

    def apply_column_style(self, cols_name=None, color='white', bold=False, size=12, style_header=False, number_format='General'):
        """
        apply style to a whole column
        :param cols_name: the columns to apply the style to
        :param color:the color to use
        :param bold: bold or not
        :param size: the font size
        :param style_header: style the header or not
        :param number_format: style the number format
        :return:
        """
        if type(cols_name) not in [list, tuple]:
            raise TypeError("cols_name must be a list or a tuple")
        if not all(col in self.columns for col in cols_name):
            raise KeyError("one of the columns in {} wasn't found".format(cols_name))
        for col_name in cols_name:
            if style_header:
                self.columns[self.columns.get_loc(col_name)].style = Styler(color=color, bold=bold, size=size, number_format=number_format).create_style()
            for index in self.index:
                self.ix[index, col_name].style = Styler(color=color, bold=bold, size=size, number_format=number_format).create_style()


