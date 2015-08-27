# coding:utf-8
import pandas as pd
from container import Container
from styler import Styler
import numpy as np
from openpyxl.cell import get_column_letter
from copy import deepcopy


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
            self.data_df = pd.DataFrame(obj).applymap(lambda x: x if isinstance(x, Container) else Container(x))
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
            sheet.cell(row=startrow + 1, column=col_index + startcol + 1).style = Styler(bg_color=current_fgColor, bold=True, font_size=current_size).create_style()
            try:  # try to apply change of width
                column_letter = get_column_letter(col_index + startcol + 1)
                sheet.column_dimensions[column_letter].width = self.data_df.columns[col_index].width
            except AttributeError:
                pass  # width attribute does not exists since no change of width has occurred
            for row_index, index in enumerate(self.data_df.index):
                sheet.cell(row=row_index + startrow + 2, column=col_index + startcol + 1).style = self.data_df.ix[index, column].style
        for row_index in xrange(startrow, startrow + len(self.data_df.index)):
            try:  # try to apply change of height
                sheet.row_dimensions[row_index + 1].height = self.data_df.ix[row_index, 0].height
            except AttributeError:
                pass  # height attribute does not exists since no change of height has occurred

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

    def apply_style_by_indexes(self, indexes_to_style=None, cols_to_style=None, bg_color='white', bold=False, font_size=12,
                               number_format='General', row_height=None):
        """
        applies a certain style to the provided indexes in the dataframe in the provided columns
        :param indexes_to_style: indexes to apply the style to
        :param cols_to_style: the columns to apply the style to, if not provided all the columns will be styled
        :param bg_color: the color to use
        :param bold: bold or not
        :param font_size: the font size
        :param number_format: style the number format
        :param row_height: determine the height of the row
        :return:
        """
        if cols_to_style is not None and type(cols_to_style) not in [list, tuple]:
            raise TypeError("cols_name must be a list or a tuple")
        elif cols_to_style is None:
            cols_to_style = list(self.data_df.columns)
        if row_height:
            try:
                float(row_height)
            except ValueError:
                raise ValueError('row height must be numeric value')
            if row_height <= 0:
                raise ValueError('row height must be positive number')
        for index in indexes_to_style:
            for col in cols_to_style:
                self.ix[index, col].style = Styler(bg_color=bg_color, bold=bold, font_size=font_size,
                                                   number_format=number_format).create_style()
            if row_height:
                self.ix[index, 0].height = row_height

    def apply_column_style(self, cols_to_style=None, bg_color='white', bold=False, font_size=12, style_header=False,
                           number_format='General', column_width=None):
        """
        apply style to a whole column
        :param cols_to_style: the columns to apply the style to
        :param bg_color:the color to use
        :param bold: bold or not
        :param font_size: the font size
        :param style_header: style the header or not
        :param number_format: style the number format
        :param column_width: determine the width of the column
        :return:
        """
        if type(cols_to_style) not in [list, tuple]:
            raise TypeError("cols_name must be a list or a tuple")
        if not all(col in self.columns for col in cols_to_style):
            raise KeyError("one of the columns in {} wasn't found".format(cols_to_style))
        if column_width:
            try:
                float(column_width)
            except ValueError:
                raise ValueError('column width must be numeric value')
            if column_width <= 0:
                raise ValueError('column width must be positive number')

        for col_name in cols_to_style:
            if style_header:
                self.columns[self.columns.get_loc(col_name)].style = Styler(bg_color=bg_color, bold=bold, font_size=font_size, number_format=number_format).create_style()
            if column_width:
                self.columns[self.columns.get_loc(col_name)].width = column_width
            for index in self.index:
                self.ix[index, col_name].style = Styler(bg_color=bg_color, bold=bold, font_size=font_size, number_format=number_format).create_style()

    def apply_headers_style(self, bg_color='white', bold=True, font_size=12, number_format='General'):
        for column in self.data_df.columns:
            column.style = Styler(bg_color=bg_color, bold=bold, font_size=font_size, number_format=number_format).create_style()

    def rename(self, columns=None, inplace=False):
        if not isinstance(columns, dict):
            raise TypeError("'columns' must be a dictionary")
        if inplace:
            for column in self.data_df.columns:
                column.value = columns[column]
        else:
            new_style_frame = deepcopy(self)
            for column in new_style_frame.data_df.columns:
                column.value = columns[column]
            return new_style_frame

