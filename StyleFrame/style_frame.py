# coding:utf-8
import sys
import pandas as pd
import numpy as np
from openpyxl import cell, load_workbook
from copy import deepcopy
import datetime as dt
import warnings

DEPRECATION_MSG = "Directly passing style specifiers\n('bg_color', 'bold', etc.) is deprecated.\nPass a StyleFrame.Styler object instead as styler_obj."

PY2 = sys.version_info[0] == 2

# Python 2
if PY2:
    # noinspection PyUnresolvedReferences
    from container import Container
    # noinspection PyUnresolvedReferences
    from styler import Styler
    # noinspection PyUnresolvedReferences
    import utils

# Python 3
else:
    from StyleFrame.container import Container
    from StyleFrame.styler import Styler
    from StyleFrame import utils


class StyleFrame(object):
    """
    A wrapper class that wraps pandas DataFrame.
    Stores container objects that have values and Styles that will be applied to excel
    """

    def __init__(self, obj, styler_obj=None):
        from_another_styleframe = False
        if styler_obj:
            if not isinstance(styler_obj, Styler):
                raise TypeError('styler_obj must be {}, got {} instead.'.format(Styler.__name__, type(styler_obj).__name__))
            styler_obj = styler_obj.create_style()
        if isinstance(obj, pd.DataFrame):
            if len(obj) == 0:
                self.data_df = deepcopy(obj)
            else:
                self.data_df = obj.applymap(lambda x: Container(x, styler_obj) if not isinstance(x, Container) else x)
        elif isinstance(obj, pd.Series):
            self.data_df = obj.apply(lambda x: Container(x, styler_obj) if not isinstance(x, Container) else x)
        elif isinstance(obj, (dict, list)):
            self.data_df = pd.DataFrame(obj).applymap(lambda x: Container(x, styler_obj) if not isinstance(x, Container) else x)
        elif isinstance(obj, StyleFrame):
            self.data_df = deepcopy(obj.data_df)
            from_another_styleframe = True
        else:
            raise TypeError("{} __init__ doesn't support {}".format(type(self).__name__, type(obj).__name__))
        self.data_df.columns = [Container(col, styler_obj) if not isinstance(col, Container) else deepcopy(col)
                                for col in self.data_df.columns]
        self.data_df.index = [Container(index, styler_obj) if not isinstance(index, Container) else deepcopy(index)
                              for index in self.data_df.index]

        self._columns_width = obj._columns_width if from_another_styleframe else {}
        self._rows_height = obj._rows_height if from_another_styleframe else {}
        self._custom_headers_style = obj._custom_headers_style if from_another_styleframe else False

    def __str__(self):
        return str(self.data_df)

    def __unicode__(self):
        if PY2:
            return unicode(self.data_df)
        return str(self.data_df)

    def __len__(self):
        return len(self.data_df)

    def __getitem__(self, item):
        if isinstance(item, pd.Series):
            return self.data_df.__getitem__(item).index
        elif isinstance(item, list):
            return StyleFrame(self.data_df.__getitem__(item))
        else:
            return self.data_df.__getitem__(item)

    def __setitem__(self, key, value):
        if isinstance(value, pd.Series):
            self.data_df.__setitem__(Container(key), map(Container, value))
        else:
            self.data_df.__setitem__(Container(key), Container(value))

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
    def read_excel(cls, path, sheetname='Sheet1', read_style=False, **kwargs):
        def _read_style():
            sheet = load_workbook(path).get_sheet_by_name(sheetname)
            for row_index, sf_index in enumerate(sf.index, start=1):
                for col_index, col_name in enumerate(sf.columns, start=1):
                    sf.ix[sf_index - 1, col_name].style = sheet.cell(row=row_index, column=col_index).style

        sf = StyleFrame(pd.read_excel(path, sheetname=sheetname, **kwargs))
        if not read_style:
            return sf
        _read_style()
        return sf

    # noinspection PyPep8Naming
    @classmethod
    def ExcelWriter(cls, path):
        return pd.ExcelWriter(path, engine='openpyxl')

    def to_excel(self, excel_writer='output.xlsx', sheet_name='Sheet1', na_rep='', float_format=None, columns=None,
                 header=True, index=False, index_label=None, startrow=0, startcol=0, merge_cells=True, encoding=None,
                 inf_rep='inf', allow_protection=False, right_to_left=False, columns_to_hide=None,
                 row_to_add_filters=None, columns_and_rows_to_freeze=None):
        """Saves the dataframe to excel and applies the styles.

        :param right_to_left: sets the sheet to be right to left.
        :param columns_to_hide: single column, list or tuple of columns to hide, may be column index (starts from 1)
                                column name or column letter.
        :param allow_protection: allow to protect the sheet and the cells that specified as protected.
        :param row_to_add_filters: add filters to the given row, starts from zero (zero is to add filters to columns).
        :param columns_and_rows_to_freeze: column and row string to freeze for example: C3 will freeze columns: A,B and rows: 1,2.
            Read Pandas' documentation about the other parameters
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
            if not isinstance(column_to_convert, (int, basestring if PY2 else str, Container)):
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

        if isinstance(excel_writer, basestring if PY2 else str):
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
            self.apply_headers_style()

        # Iterating over the dataframe's elements and applying their styles
        # openpyxl's rows and cols start from 1,1 while the dataframe is 0,0
        for col_index, column in enumerate(self.data_df.columns):
            sheet.cell(row=startrow + 1, column=col_index + startcol + 1).style = column.style

            for row_index, index in enumerate(self.data_df.index):
                current_cell = sheet.cell(row=row_index + startrow + 2, column=col_index + startcol + 1)
                try:
                    if PY2:
                        if '=HYPERLINK' in unicode(current_cell.value):
                            current_bg_color = current_cell.style.fill.fgColor.rgb
                            current_font_size = current_cell.style.font.size
                            current_cell.style = Styler(bg_color=current_bg_color,
                                                        font_color=utils.colors.blue,
                                                        font_size=current_font_size,
                                                        number_format=utils.number_formats.general,
                                                        underline=utils.underline.single).create_style()
                        else:
                            current_cell.style = self.data_df.ix[index, column].style

                    else:
                        if '=HYPERLINK' in str(current_cell.value):
                            current_bg_color = current_cell.style.fill.fgColor.rgb
                            current_font_size = current_cell.style.font.size
                            current_cell.style = Styler(bg_color=current_bg_color,
                                                        font_color=utils.colors.blue,
                                                        font_size=current_font_size,
                                                        number_format=utils.number_formats.general,
                                                        underline='single').create_style()
                        else:
                            current_cell.style = self.data_df.ix[index, column].style

                except AttributeError:  # if the element in the dataframe is not Container creating a default style
                    current_cell.style = Styler().create_style()

        for column in self._columns_width:
            column_letter = get_column_as_letter(column_to_convert=column)
            sheet.column_dimensions[column_letter].width = self._columns_width[column]

        for row in self._rows_height:
            if row in sheet.row_dimensions:
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
            if not isinstance(columns_and_rows_to_freeze, basestring if PY2 else str) or len(columns_and_rows_to_freeze) < 2:
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

    def apply_style_by_indexes(self, indexes_to_style, cols_to_style=None, styler_obj=None, bg_color=utils.colors.white,
                               bold=False, font='Arial', font_size=12, font_color=utils.colors.black, protection=False,
                               number_format=None, underline=None):
        """Applies a certain style to the provided indexes in the dataframe in the provided columns

        :param indexes_to_style: indexes to apply the style to
        :param cols_to_style: the columns to apply the style to, if not provided all the columns will be styled
        :param bg_color: the color to use.
        :param bold: bold or not
        :param font_size: the font size
        :param font_color: the font color
        :param protection: to protect the cell from changes or not
        :param number_format: modify the number format
        :param underline: the type of text underline
        :return: self
        :rtype: StyleFrame
        """

        if styler_obj:
            if not isinstance(styler_obj, Styler):
                raise TypeError(
                    'styler_obj must be {}, got {} instead.'.format(Styler.__name__, type(styler_obj).__name__))
            styler_obj = styler_obj.create_style()
        else:
            warnings.warn(DEPRECATION_MSG, DeprecationWarning)

        default_number_formats = {pd.tslib.Timestamp: 'DD/MM/YY HH:MM',
                                  dt.date: 'DD/MM/YY',
                                  dt.time: 'HH:MM'}

        indexes_number_format = number_format
        values_number_format = number_format

        if cols_to_style and not isinstance(cols_to_style, (list, tuple)):
            cols_to_style = [cols_to_style]
        elif not cols_to_style:
            cols_to_style = list(self.data_df.columns)
            for i in indexes_to_style:
                if not number_format:
                    indexes_number_format = default_number_formats.get(type(i.value), utils.number_formats.general)

                i.style = styler_obj or Styler(bg_color=bg_color, bold=bold, font=font, font_size=font_size,
                                               font_color=font_color, protection=protection,
                                               number_format=indexes_number_format, underline=underline).create_style()

        if not isinstance(indexes_to_style, (list, tuple, pd.Index)):
            indexes_to_style = [indexes_to_style]

        for index in indexes_to_style:
            for col in cols_to_style:
                if not number_format:
                    values_number_format = default_number_formats.get(type(self.ix[index.value, col].value), utils.number_formats.general)

                self.ix[index.value, col].style = styler_obj or Styler(bg_color=bg_color, bold=bold, font=font,
                                                                       font_size=font_size, font_color=font_color,
                                                                       protection=protection,
                                                                       number_format=values_number_format,
                                                                       underline=underline).create_style()
        return self

    def apply_column_style(self, cols_to_style, styler_obj=None, bg_color=utils.colors.white, font="Arial", bold=False, font_size=12, protection=False,
                           font_color=utils.colors.black, style_header=False, number_format=utils.number_formats.general,
                           underline=None):
        """apply style to a whole column

        :param cols_to_style: the columns to apply the style to
        :param bg_color:the color to use
        :param bold: bold or not
        :param font_size: the font size
        :param font_color: the font color
        :param style_header: style the header or not
        :param number_format: style the number format
        :param protection: to protect the column from changes or not
        :param underline: the type of text underline
        :return: self
        :rtype: StyleFrame
        """

        if styler_obj:
            if not isinstance(styler_obj, Styler):
                raise TypeError(
                    'styler_obj must be {}, got {} instead.'.format(Styler.__name__, type(styler_obj).__name__))
            styler_obj = styler_obj.create_style()
        else:
            warnings.warn(DEPRECATION_MSG, DeprecationWarning)

        if not isinstance(cols_to_style, (list, tuple)):
            cols_to_style = [cols_to_style]
        if not all(col in self.columns for col in cols_to_style):
            raise KeyError("one of the columns in {} wasn't found".format(cols_to_style))
        for col_name in cols_to_style:
            if style_header:
                self.columns[self.columns.get_loc(col_name)].style = styler_obj or Styler(bg_color=bg_color, bold=bold,
                                                                                          font=font, font_size=font_size,
                                                                                          font_color=font_color,
                                                                                          protection=protection,
                                                                                          number_format=number_format,
                                                                                          underline=underline).create_style()
                self._custom_headers_style = True
            for index in self.index:
                if isinstance(self.ix[index, col_name].value, pd.tslib.Timestamp):
                    number_format = utils.number_formats.date_time
                elif isinstance(self.ix[index, col_name].value, dt.date):
                    number_format = utils.number_formats.date
                elif isinstance(self.ix[index, col_name].value, dt.time):
                    number_format = utils.number_formats.time_24_hours
                self.ix[index, col_name].style = styler_obj or Styler(bg_color=bg_color, bold=bold, font=font,
                                                                      font_size=font_size, protection=protection,
                                                                      font_color=font_color, number_format=number_format,
                                                                      underline=underline).create_style()
        return self

    def apply_headers_style(self, styler_obj=None, bg_color=utils.colors.white, bold=True, font="Arial", font_size=12,
                            font_color=utils.colors.black, protection=False, number_format=utils.number_formats.general,
                            underline=None):
        """Apply style to the headers only

        :param bg_color:the color to use
        :param bold: bold or not
        :param font_size: the font size
        :param font_color: the font color
        :param number_format: openpy_style_obj the number format
        :param protection: to protect the column from changes or not
        :param underline: the type of text underline
        :return: self
        :rtype: StyleFrame
        """

        if styler_obj:
            if not isinstance(styler_obj, Styler):
                raise TypeError(
                    'styler_obj must be {}, got {} instead.'.format(Styler.__name__, type(styler_obj).__name__))

            styler_obj = styler_obj.create_style()
        else:
            warnings.warn(DEPRECATION_MSG, DeprecationWarning)

        for column in self.data_df.columns:
            column.style = styler_obj or Styler(bg_color=bg_color, bold=bold, font=font, font_size=font_size,
                                                font_color=font_color, protection=protection, number_format=number_format,
                                                underline=underline).create_style()
        self._custom_headers_style = True
        return self

    def set_column_width(self, columns, width):
        """Set the width of the given columns

        :param columns: a single or a list/tuple of column name, index or letter to change their width
        :param width: numeric positive value of the new width
        :return: self
        :rtype: StyleFrame
        """
        if not isinstance(columns, (set, list, tuple)):
            columns = [columns]
        try:
            width = float(width)
        except TypeError:
            raise TypeError('columns width must be numeric value')

        if width <= 0:
            raise ValueError('columns width must be positive')

        for column in columns:
            if not isinstance(column, (int, basestring if PY2 else str, Container)):
                raise TypeError("column must be an index, column letter or column name")
            self._columns_width[column] = width

        return self

    def set_column_width_dict(self, col_width_dict):
        """
        :param col_width_dict: dictionary from tuple of columns to new width
        :return: self
        :rtype: StyleFrame
        """
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
        if not isinstance(rows, (set, list, tuple)):
            rows = [rows]
        try:
            height = float(height)
        except TypeError:
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
            if not isinstance(height, (int, float)) or height <= 0:
                raise ValueError('rows height must be positive value')
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

