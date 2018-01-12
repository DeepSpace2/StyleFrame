# coding:utf-8

import datetime as dt
import numpy as np
import pandas as pd
import sys

from . import utils
from copy import deepcopy
from openpyxl import cell, load_workbook

PY2 = sys.version_info[0] == 2

# Python 2
if PY2:
    # noinspection PyUnresolvedReferences
    from container import Container
    # noinspection PyUnresolvedReferences
    from series import Series
    # noinspection PyUnresolvedReferences
    from styler import Styler

# Python 3
else:
    from StyleFrame.container import Container
    from StyleFrame.styler import Styler
    from StyleFrame.series import Series

try:
    pd_timestamp = pd.Timestamp
except AttributeError:
    pd_timestamp = pd.tslib.Timestamp

str_type = basestring if PY2 else str
unicode_type = unicode if PY2 else str


class StyleFrame(object):
    """
    A wrapper class that wraps pandas DataFrame.
    Stores container objects that have values and Styles that will be applied to excel
    """
    P_FACTOR = 1.3
    A_FACTOR = 13

    def __init__(self, obj, styler_obj=None):
        from_another_styleframe = False
        if styler_obj and not isinstance(styler_obj, Styler):
            raise TypeError('styler_obj must be {}, got {} instead.'.format(Styler.__name__, type(styler_obj).__name__))
        if isinstance(obj, pd.DataFrame):
            if len(obj) == 0:
                self.data_df = deepcopy(obj)
            else:
                self.data_df = obj.applymap(lambda x: Container(x, deepcopy(styler_obj)) if not isinstance(x, Container) else x)
        elif isinstance(obj, pd.Series):
            self.data_df = obj.apply(lambda x: Container(x, deepcopy(styler_obj)) if not isinstance(x, Container) else x)
        elif isinstance(obj, (dict, list)):
            self.data_df = pd.DataFrame(obj).applymap(lambda x: Container(x, deepcopy(styler_obj)) if not isinstance(x, Container) else x)
        elif isinstance(obj, StyleFrame):
            self.data_df = deepcopy(obj.data_df)
            from_another_styleframe = True
        else:
            raise TypeError("{} __init__ doesn't support {}".format(type(self).__name__, type(obj).__name__))
        self.data_df.columns = [Container(col, deepcopy(styler_obj)) if not isinstance(col, Container) else deepcopy(col)
                                for col in self.data_df.columns]
        self.data_df.index = [Container(index, deepcopy(styler_obj)) if not isinstance(index, Container) else deepcopy(index)
                              for index in self.data_df.index]

        self._columns_width = obj._columns_width if from_another_styleframe else {}
        self._rows_height = obj._rows_height if from_another_styleframe else {}
        self._custom_headers_style = obj._custom_headers_style if from_another_styleframe else False

    def __str__(self):
        return str(self.data_df)

    def __unicode__(self):
        return unicode_type(self.data_df)

    def __len__(self):
        return len(self.data_df)

    def __getitem__(self, item):
        if isinstance(item, pd.Series):
            return self.data_df.__getitem__(item).index
        elif isinstance(item, list):
            return StyleFrame(self.data_df.__getitem__(item))
        else:
            return Series(self.data_df.__getitem__(item))

    def __setitem__(self, key, value):
        if isinstance(value, pd.Series):
            self.data_df.__setitem__(Container(key), map(Container, value))
        else:
            self.data_df.__setitem__(Container(key), Container(value))

    def __delitem__(self, item):
        return self.data_df.__delitem__(item)

    def __getattr__(self, attr):
        known_attrs = {'loc': self.data_df.loc,
                       'iloc': self.data_df.iloc,
                       'applymap': self.data_df.applymap,
                       'groupby': self.data_df.groupby,
                       'index': self.data_df.index,
                       'columns': self.data_df.columns,
                       'fillna': self.data_df.fillna}
        # future versions of pandas may not have .ix (deprecated since 0.20)
        if getattr(self.data_df, 'ix'):
            known_attrs['ix'] = self.data_df.ix
        if attr in known_attrs and hasattr(self.data_df, attr):
            return known_attrs[attr]
        else:
            raise AttributeError("'{}' object has no attribute '{}'".format(type(self).__name__, attr))

    @classmethod
    def read_excel(cls, path, sheetname='Sheet1', read_style=False, **kwargs):
        def _read_style():
            sheet = load_workbook(path).get_sheet_by_name(sheetname)
            for col_index, col_name in enumerate(sf.columns, start=1):
                sf.columns[col_index - 1].style = sheet.cell(row=1, column=col_index).style
                for row_index, sf_index in enumerate(sf.index, start=2):
                    sf.loc[sf_index, col_name].style = sheet.cell(row=row_index, column=col_index).style

        sf = cls(pd.read_excel(path, sheetname=sheetname, **kwargs))
        if read_style:
            _read_style()
            sf._custom_headers_style = True
        return sf

    # noinspection PyPep8Naming
    @classmethod
    def ExcelWriter(cls, path):
        return pd.ExcelWriter(path, engine='openpyxl')

    @property
    def row_indexes(self):
        """Excel row indexes.

        StyleFrame row indexes (including the headers) according to the excel file format.
        Mostly used to set rows height.
        Excel indexes format starts from index 1.

        :rtype: tuple
        """

        return tuple(range(1, len(self) + 2))

    def to_excel(self, excel_writer='output.xlsx', sheet_name='Sheet1', na_rep='', float_format=None, columns=None,
                 header=True, index=False, index_label=None, startrow=0, startcol=0, merge_cells=True, encoding=None,
                 inf_rep='inf', allow_protection=False, right_to_left=False, columns_to_hide=None,
                 row_to_add_filters=None, columns_and_rows_to_freeze=None, best_fit=None):
        """Saves the dataframe to excel and applies the styles.

        :param right_to_left: sets the sheet to be right to left.
        :param columns_to_hide: single column, list or tuple of columns to hide, may be column index (starts from 1)
                                column name or column letter.
        :param allow_protection: allow to protect the sheet and the cells that specified as protected.
        :param row_to_add_filters: add filters to the given row, starts from zero (zero is to add filters to columns).
        :param columns_and_rows_to_freeze: column and row string to freeze for example: C3 will freeze columns: A,B and rows: 1,2.

        See Pandas' to_excel documentation about the other parameters
        """

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

        def get_column_as_letter(column_to_convert):
            if not isinstance(column_to_convert, (int, str_type, Container)):
                raise TypeError("column must be an index, column letter or column name")
            column_as_letter = None
            if column_to_convert in self.data_df.columns:  # column name
                column_index = self.data_df.columns.get_loc(
                    column_to_convert) + startcol + 1  # worksheet columns index start from 1
                column_as_letter = cell.get_column_letter(column_index)

            elif isinstance(column_to_convert, int) and column_to_convert >= 1:  # column index
                column_as_letter = cell.get_column_letter(startcol + column_to_convert)
            elif column_to_convert in sheet.column_dimensions:  # column letter
                column_as_letter = column_to_convert

            if column_as_letter is None or column_as_letter not in sheet.column_dimensions:
                raise IndexError("column: %s is out of columns range." % column_to_convert)

            return column_as_letter

        def get_range_of_cells_for_specific_row(row_index):
            start_letter = get_column_as_letter(column_to_convert=self.data_df.columns[0])
            end_letter = get_column_as_letter(column_to_convert=self.data_df.columns[-1])
            return '{start_letter}{start_index}:{end_letter}{end_index}'.format(start_letter=start_letter,
                                                                                start_index=startrow + row_index + 1,
                                                                                end_letter=end_letter,
                                                                                end_index=startrow + row_index + 1)

        if len(self.data_df) > 0:
            export_df = self.data_df.applymap(get_values)

        else:
            export_df = deepcopy(self.data_df)

        export_df.columns = [col.value for col in export_df.columns]
        # noinspection PyTypeChecker
        export_df.index = [row_index.value for row_index in export_df.index]

        if isinstance(excel_writer, str_type):
            excel_writer = self.ExcelWriter(excel_writer)

        export_df.to_excel(excel_writer, sheet_name=sheet_name, na_rep=na_rep, float_format=float_format, index=index,
                           columns=columns, header=header, index_label=index_label, startrow=startrow,
                           startcol=startcol, engine='openpyxl', merge_cells=merge_cells, encoding=encoding,
                           inf_rep=inf_rep)

        sheet = excel_writer.book.get_sheet_by_name(sheet_name)

        sheet.sheet_view.rightToLeft = right_to_left

        self.data_df.fillna(Container('NaN'), inplace=True)

        if index:
            for row_index, index in enumerate(self.data_df.index):
                sheet.cell(row=startrow + row_index + 2, column=startcol + 1).style = index.style
            startcol += 1

        if header and not self._custom_headers_style:
            self.apply_headers_style(Styler.default_header_style())

        # Iterating over the dataframe's elements and applying their styles
        # openpyxl's rows and cols start from 1,1 while the dataframe is 0,0
        for col_index, column in enumerate(self.data_df.columns):
            sheet.cell(row=startrow + 1, column=col_index + startcol + 1).style = column.style.create_style()
            for row_index, index in enumerate(self.data_df.index):
                current_cell = sheet.cell(row=row_index + startrow + 2, column=col_index + startcol + 1)
                data_df_style = self.data_df.loc[index, column].style

                try:
                    if '=HYPERLINK' in unicode_type(current_cell.value):
                        data_df_style.font_color = utils.colors.blue
                        data_df_style.underline = utils.underline.single
                    else:
                        if best_fit and column in best_fit:
                            data_df_style.wrap_text = False
                            data_df_style.shrink_to_fit = False

                    current_cell.style = data_df_style.create_style()

                except AttributeError:  # if the element in the dataframe is not Container creating a default style
                    current_cell.style = Styler().create_style()

        if best_fit:
            self.set_column_width_dict({column: (max(self.data_df[column].str.len()) + self.A_FACTOR) * self.P_FACTOR
                                        for column in best_fit})

        for column in self._columns_width:
            column_letter = get_column_as_letter(column_to_convert=column)
            sheet.column_dimensions[column_letter].width = self._columns_width[column]

        for row in self._rows_height:
            if row + startrow in sheet.row_dimensions:
                sheet.row_dimensions[startrow + row].height = self._rows_height[row]
            else:
                raise IndexError('row: {} is out of range'.format(row))

        if row_to_add_filters is not None:
            try:
                row_to_add_filters = int(row_to_add_filters)
                if (row_to_add_filters + startrow + 1) not in sheet.row_dimensions:
                    raise IndexError('row: {} is out of rows range'.format(row_to_add_filters))
                sheet.auto_filter.ref = get_range_of_cells_for_specific_row(row_index=row_to_add_filters)
            except (TypeError, ValueError):
                raise TypeError("row must be an index and not {}".format(type(row_to_add_filters)))

        if columns_and_rows_to_freeze is not None:
            if not isinstance(columns_and_rows_to_freeze, str_type) or len(columns_and_rows_to_freeze) < 2:
                raise TypeError("columns_and_rows_to_freeze must be a str for example: 'C3'")
            if columns_and_rows_to_freeze[0] not in sheet.column_dimensions:
                raise IndexError("column: %s is out of columns range." % columns_and_rows_to_freeze[0])
            if int(columns_and_rows_to_freeze[1]) not in sheet.row_dimensions:
                raise IndexError("row: %s is out of rows range." % columns_and_rows_to_freeze[1])
            sheet.freeze_panes = sheet[columns_and_rows_to_freeze]

        if allow_protection:
            sheet.protection.autoFilter = False
            sheet.protection.enable()

        # Iterating over the columns_to_hide and check if the format is columns name, column index as number or letter
        if columns_to_hide:
            if not isinstance(columns_to_hide, (list, tuple)):
                columns_to_hide = [columns_to_hide]

            for column in columns_to_hide:
                column_letter = get_column_as_letter(column_to_convert=column)
                sheet.column_dimensions[column_letter].hidden = True

        return excel_writer

    def apply_style_by_indexes(self, indexes_to_style, styler_obj, cols_to_style=None, height=None,):
        """Applies a certain style to the provided indexes in the dataframe in the provided columns

        :param indexes_to_style: indexes to apply the style to
        :param styler_obj: the styler object that contains the style to be applied
        :type styler_obj: Styler
        :param cols_to_style: the columns to apply the style to, if not provided all the columns will be styled
        :param height: non-default height for the given rows
        :return: self
        :rtype: StyleFrame
        """

        if not isinstance(styler_obj, Styler):
            raise TypeError('styler_obj must be {}, got {} instead.'.format(Styler.__name__, type(styler_obj).__name__))

        if isinstance(indexes_to_style, (list, tuple, int)):
            indexes_to_style = self.index[indexes_to_style]

        if isinstance(indexes_to_style, Container):
            indexes_to_style = pd.Index([indexes_to_style])

        default_number_formats = {pd_timestamp: 'DD/MM/YY HH:MM',
                                  dt.date: 'DD/MM/YY',
                                  dt.time: 'HH:MM'}

        indexes_number_format = styler_obj.number_format
        values_number_format = styler_obj.number_format

        if cols_to_style and not isinstance(cols_to_style, (list, tuple)):
            cols_to_style = [cols_to_style]
        elif not cols_to_style:
            cols_to_style = list(self.data_df.columns)
            for i in indexes_to_style:
                if styler_obj.number_format == utils.number_formats.general:
                    indexes_number_format = default_number_formats.get(type(i.value), utils.number_formats.general)

                styler_obj.number_format = indexes_number_format
                i.style = styler_obj

        for index in indexes_to_style:
            for col in cols_to_style:
                if styler_obj.number_format == utils.number_formats.general:
                    values_number_format = default_number_formats.get(
                        type(self.iloc[index.value, self.columns.get_loc(col)].value),
                        utils.number_formats.general)

                styler_obj.number_format = values_number_format
                self.iloc[index.value, self.columns.get_loc(col)].style = styler_obj

        if height:
            # Add offset 2 since rows do not include the headers and they starts from 1 (not 0).
            rows_indexes_for_height_change = [self.index.get_loc(idx) + 2 for idx in indexes_to_style]
            self.set_row_height(rows=rows_indexes_for_height_change, height=height)

        return self

    def apply_column_style(self, cols_to_style, styler_obj, style_header=False, use_default_formats=True, width=None):
        """apply style to a whole column

        :param cols_to_style: the columns to apply the style to
        :param styler_obj: the styler object that contains the style to be applied
        :type styler_obj: Styler
        :param style_header: if True, style the headers as well
        :type style_header: bool
        :param use_default_formats: if True, use predefined styles for dates and times
        :type use_default_formats: bool
        :param width: non-default width for the given columns
        :return: self
        :rtype: StyleFrame
        """

        if not isinstance(styler_obj, Styler):
            raise TypeError('styler_obj must be {}, got {} instead.'.format(Styler.__name__, type(styler_obj).__name__))

        if not isinstance(cols_to_style, (list, tuple, pd.Index)):
            cols_to_style = [cols_to_style]
        if not all(col in self.columns for col in cols_to_style):
            raise KeyError("one of the columns in {} wasn't found".format(cols_to_style))
        for col_name in cols_to_style:
            if style_header:
                self.columns[self.columns.get_loc(col_name)].style = styler_obj
                self._custom_headers_style = True
            for index in self.index:
                if use_default_formats:
                    if isinstance(self.loc[index, col_name].value, pd_timestamp):
                        styler_obj.number_format = utils.number_formats.date_time
                    elif isinstance(self.loc[index, col_name].value, dt.date):
                        styler_obj.number_format = utils.number_formats.date
                    elif isinstance(self.loc[index, col_name].value, dt.time):
                        styler_obj.number_format = utils.number_formats.time_24_hours

                self.loc[index, col_name].style = styler_obj

        if width:
            self.set_column_width(columns=cols_to_style, width=width)

        return self

    def apply_headers_style(self, styler_obj):
        """Apply style to the headers only

        :param styler_obj: the styler object that contains the style to be applied
        :type styler_obj: Styler
        :return: self
        :rtype: StyleFrame
        """

        if not isinstance(styler_obj, Styler):
            raise TypeError('styler_obj must be {}, got {} instead.'.format(Styler.__name__, type(styler_obj).__name__))

        styler_obj = styler_obj

        for column in self.data_df.columns:
            column.style = styler_obj
        self._custom_headers_style = True
        return self

    def set_column_width(self, columns, width):
        """Set the width of the given columns

        :param columns: a single or a list/tuple of column name, index or letter to change their width
        :param width: numeric positive value of the new width
        :return: self
        :rtype: StyleFrame
        """

        if not isinstance(columns, (set, list, tuple, pd.Index)):
            columns = [columns]
        try:
            width = float(width)
        except ValueError:
            raise TypeError('columns width must be numeric value')

        if width <= 0:
            raise ValueError('columns width must be positive')

        for column in columns:
            if not isinstance(column, (int, str_type, Container)):
                raise TypeError("column must be an index, column letter or column name")
            self._columns_width[column] = width

        return self

    def set_column_width_dict(self, col_width_dict):
        """
        :param col_width_dict: dictionary from tuple of columns to new width
        :type col_width_dict: dict
        :return: self
        :rtype: StyleFrame
        """

        if not isinstance(col_width_dict, dict):
            raise TypeError("'col_width_dict' must be a dictionary")
        for cols, width in col_width_dict.items():
            self.set_column_width(cols, width)

        return self

    def set_row_height(self, rows, height):
        """ Set the height of the given rows

        :param rows: a single row index, list of indexes or tuple of indexes to change their height
        :param height: numeric positive value of the new height
        :return: self
        :rtype: StyleFrame
        """

        if not isinstance(rows, (set, list, tuple, pd.Index)):
            rows = [rows]
        try:
            height = float(height)
        except ValueError:
            raise TypeError('rows height must be numeric value')

        if height <= 0:
            raise ValueError('rows height must be positive')
        for row in rows:
            try:
                row = int(row)
            except TypeError:
                raise TypeError("row must be an index")

            self._rows_height[row] = height

        return self

    def set_row_height_dict(self, row_height_dict):
        """
        :param row_height_dict: dictionary from tuple of rows to new height
        :type row_height_dict: dict
        :return: self
        :rtype: StyleFrame
        """

        if not isinstance(row_height_dict, dict):
            raise TypeError("'row_height_dict' must be a dictionary")
        for rows, height in row_height_dict.items():
            self.set_row_height(rows, height)
        return self

    def rename(self, columns=None, inplace=False):
        """Renames the underlying dataframe's columns

        :param columns: a dictionary, old_col_name -> new_col_name
        :type columns: dict
        :param inplace: whether to rename the columns inplace or return a new StyleFrame object
        :return: self if inplace=True, new StyleFrame object if inplace=False
        """

        if not isinstance(columns, dict):
            raise TypeError("'columns' must be a dictionary")

        sf = self if inplace else StyleFrame(self)

        new_columns = [col if col not in columns else Container(columns[col], col.style)
                       for col in sf.data_df.columns]
        sf.data_df.columns = new_columns

        sf._columns_width.update({new_col_name: sf._columns_width.pop(old_col_name)
                                  for old_col_name, new_col_name in columns.items()
                                  if old_col_name in sf._columns_width})

        return sf

    def style_alternate_rows(self, styles):
        """Applies the provided styles to rows in an alternating manner.

        :param styles: styles to apply
        :type styles: list|tuple
        :return: self
        """

        num_of_styles = len(styles)
        split_indexes = (self.index[i::num_of_styles] for i in range(num_of_styles))
        for i, indexes in enumerate(split_indexes):
            self.apply_style_by_indexes(indexes, styles[i])
        return self
