# coding:utf-8
import pandas as pd
from container import Container
from styler import Styler, number_formats, colors
import numpy as np
import openpyxl.cell
from copy import deepcopy
import datetime as dt


class StyleFrame(object):
    """
    A wrapper class that wraps pandas DataFrame.
    Stores container objects that have values and Styles that will be applied to excel
    """

    def __init__(self, obj):
        if isinstance(obj, pd.DataFrame):
            if len(obj) == 0:
                self.data_df = deepcopy(obj)
            else:
                self.data_df = obj.applymap(lambda x: Container(x) if not isinstance(x, Container) else x)
        elif isinstance(obj, pd.Series):
            self.data_df = obj.apply(lambda x: Container(x) if not isinstance(x, Container) else x)
        elif isinstance(obj, dict) or isinstance(obj, list):
            self.data_df = pd.DataFrame(obj).applymap(lambda x: Container(x) if not isinstance(x, Container) else x)
        elif isinstance(obj, StyleFrame):
            self.data_df = deepcopy(obj)
        else:
            raise TypeError("{} __init__ doesn't support {}".format(type(self).__name__, type(obj).__name__))
        self.data_df.columns = [Container(col) if not isinstance(col, Container) else col for col in self.data_df.columns]
        self.data_df.index = [Container(index) if not isinstance(index, Container) else index for index in self.data_df.index]

        self.columns_width = dict()
        self.rows_height = dict()

    def __str__(self):
        return str(self.data_df)

    def __unicode__(self):
        return unicode(self.data_df)

    def __len__(self):
        return len(self.data_df)

    def __getitem__(self, item):
        if isinstance(item,pd.Series):
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
    def read_excel(cls, path, sheetname=0, **kwargs):
        return StyleFrame(pd.read_excel(path, sheetname=sheetname, **kwargs))

    @classmethod
    def ExcelWriter(cls, path):
        return pd.ExcelWriter(path, engine='openpyxl')

    def to_excel(self, excel_writer, sheet_name='Sheet1', na_rep='', float_format=None, columns=None, header=True,
                 index=False, index_label=None, startrow=0, startcol=0, merge_cells=True, encoding=None, inf_rep='inf',
                 allow_protection=False, right_to_left=True, columns_to_hide=None, row_to_add_filters=None,
                 columns_and_rows_to_freeze=None):
        """
        Saves the dataframe to excel and applies the styles.
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
            if not isinstance(column_to_convert, (int, basestring, Container)):
                raise TypeError("column must be an index, column letter or column name")

            column_as_letter = None
            if column_to_convert in self.data_df.columns:  # column name
                column_index = self.data_df.columns.get_loc(
                    column_to_convert) + startcol + 1  # worksheet columns index start from 1
                column_as_letter = openpyxl.cell.get_column_letter(column_index)

            elif isinstance(column_to_convert, int) and column_to_convert >= 1:  # column index
                column_as_letter = openpyxl.cell.get_column_letter(startcol + column_to_convert)
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
            export_df = self.data_df.applymap(lambda x: get_values(x))

        else:
            export_df = deepcopy(self.data_df)

        export_df.columns = [col.value for col in export_df.columns]
        export_df.index = [row_index.value for row_index in export_df.index]

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
        ''' Iterating over the dataframe's elements and applying their styles '''
        ''' openpyxl's rows and cols start from 1,1 while the dataframe is 0,0 '''
        for col_index, column in enumerate(self.data_df.columns):
            current_bg_color = self.data_df.columns[col_index].style.fill.fgColor.rgb
            current_size = self.data_df.columns[col_index].style.font.size
            current_font_color = self.data_df.columns[col_index].style.font.color.index
            current_number_format = self.data_df.columns[col_index].style.number_format
            sheet.cell(row=startrow + 1, column=col_index + startcol + 1).style = Styler(bg_color=current_bg_color,
                                                                                         bold=True,
                                                                                         font_color=current_font_color,
                                                                                         font_size=current_size,
                                                                                         number_format=current_number_format).create_style()
            for row_index, index in enumerate(self.data_df.index):
                current_cell = sheet.cell(row=row_index + startrow + 2, column=col_index + startcol + 1)
                try:
                    if '=HYPERLINK' in unicode(current_cell.value):
                        current_bg_color = current_cell.style.fill.fgColor.rgb
                        current_font_size = current_cell.style.font.size
                        current_cell.style = Styler(bg_color=current_bg_color,
                                                    font_color='blue',
                                                    font_size=current_font_size,
                                                    number_format=number_formats.general,
                                                    underline='single').create_style()
                    else:
                        current_cell.style = self.data_df.ix[index, column].style
                except AttributeError:  # if the element in the dataframe is not Container creating a default style
                    current_cell.style = Styler().create_style()

        for column in self.columns_width:
            column_letter = get_column_as_letter(column_to_convert=column)
            sheet.column_dimensions[column_letter].width = self.columns_width[column]

        for row in self.rows_height:
            if row in sheet.row_dimensions:
                sheet.row_dimensions[startrow + row].height = self.rows_height[row]
            else:
                raise IndexError('row: %s is out of range' % row)

        if row_to_add_filters is not None:
            try:
                row_to_add_filters = int(row_to_add_filters)
                if (row_to_add_filters + startrow + 1) not in sheet.row_dimensions:
                    raise IndexError('row: %s is out of rows range' % row_to_add_filters)
                sheet.auto_filter.ref = get_range_of_cells_for_specific_row(row_index=row_to_add_filters)
            except TypeError:
                raise TypeError("row must be an index and not %s" % type(row_to_add_filters))

        if columns_and_rows_to_freeze is not None:
            if not isinstance(columns_and_rows_to_freeze, basestring) or len(columns_and_rows_to_freeze) < 2:
                raise TypeError("columns_and_rows_to_freeze must be a str for example: 'C3'")
            if columns_and_rows_to_freeze[0] not in sheet.column_dimensions:
                raise IndexError("column: %s is out of columns range." % columns_and_rows_to_freeze[0])
            if int(columns_and_rows_to_freeze[1]) not in sheet.row_dimensions:
                raise IndexError("row: %s is out of rows range." % columns_and_rows_to_freeze[1])
            sheet.freeze_panes = sheet[columns_and_rows_to_freeze]

        if allow_protection:
            sheet.protection.autoFilter = False
            sheet.protection.enable()

        ''' Iterating over the columns_to_hide and check if the format is columns name, column index as number or letter  '''
        if columns_to_hide is not None:
            if not isinstance(columns_to_hide, (list, tuple)):
                columns_to_hide = [columns_to_hide]

            for column in columns_to_hide:
                column_letter = get_column_as_letter(column_to_convert=column)
                sheet.column_dimensions[column_letter].hidden = True

    def apply_style_by_indexes(self, indexes_to_style, cols_to_style=None, bg_color=colors.white, bold=False,
                               font_size=12, font_color=colors.black, protection=False,
                               number_format=number_formats.general):
        """
        applies a certain style to the provided indexes in the dataframe in the provided columns
        :param indexes_to_style: indexes to apply the style to
        :param cols_to_style: the columns to apply the style to, if not provided all the columns will be styled
        :param bg_color: the color to use
        :param bold: bold or not
        :param font_size: the font size
        :param font_color: the font color
        :param protection: to protect the cell from changes or not
        :param number_format: style the number format
        :return:
        """
        if cols_to_style is not None and not isinstance(cols_to_style, (list, tuple)):
            cols_to_style = [cols_to_style]
        elif cols_to_style is None:
            cols_to_style = list(self.data_df.columns)
            for i in indexes_to_style:
                i.style = Styler(bg_color=bg_color, bold=bold, font_size=font_size,
                                 font_color=font_color, protection=protection,
                                 number_format=number_format).create_style()
        if not isinstance(indexes_to_style, (list, tuple, pd.Index)):
            indexes_to_style = [indexes_to_style]

        for index in indexes_to_style:
            for col in cols_to_style:
                self.ix[index.value, col].style = Styler(bg_color=bg_color, bold=bold, font_size=font_size,
                                                         font_color=font_color, protection=protection,
                                                         number_format=number_format).create_style()

    def apply_column_style(self, cols_to_style, bg_color=colors.white, bold=False, font_size=12, protection=False,
                           font_color=colors.black, style_header=False, number_format=number_formats.general):
        """
        apply style to a whole column
        :param cols_to_style: the columns to apply the style to
        :param bg_color:the color to use
        :param bold: bold or not
        :param font_size: the font size
        :param font_color: the font color
        :param style_header: style the header or not
        :param number_format: style the number format
        :param protection: to protect the column from changes or not
        :return:
        """
        if not isinstance(cols_to_style, (list, tuple)):
            cols_to_style = [cols_to_style]
        if not all(col in self.columns for col in cols_to_style):
            raise KeyError("one of the columns in {} wasn't found".format(cols_to_style))
        for col_name in cols_to_style:
            if style_header:
                self.columns[self.columns.get_loc(col_name)].style = Styler(bg_color=bg_color, bold=bold,
                                                                            font_size=font_size, font_color=font_color,
                                                                            protection=protection,
                                                                            number_format=number_format).create_style()
            for index in self.index:
                if isinstance(self.ix[index, col_name].value, pd.tslib.Timestamp):
                    number_format = number_formats.date_time
                elif isinstance(self.ix[index, col_name].value, dt.date):
                    number_format = number_formats.date
                elif isinstance(self.ix[index, col_name].value, dt.time):
                    number_format = number_formats.time_24_hours
                self.ix[index, col_name].style = Styler(bg_color=bg_color, bold=bold, font_size=font_size,
                                                        protection=protection, font_color=font_color,
                                                        number_format=number_format).create_style()

    def apply_headers_style(self, bg_color=colors.white, bold=True, font_size=12, font_color=colors.black,
                            protection=False, number_format=number_formats.general):
        """
        apply style to the headers only
        :param bg_color:the color to use
        :param bold: bold or not
        :param font_size: the font size
        :param font_color: the font color
        :param number_format: style the number format
        :return:
        """
        for column in self.data_df.columns:
            column.style = Styler(bg_color=bg_color, bold=bold, font_size=font_size, font_color=font_color,
                                  protection=protection, number_format=number_format).create_style()

    def set_column_width(self, columns, width):
        """
        set the width of the given columns
        :param columns: a single or a list/tuple of column name, index or letter to change their width
        :param width: numeric positive value of the new width
        :return:
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
            if not isinstance(column, (int, basestring, Container)):
                raise TypeError("column must be an index, column letter or column name")
            self.columns_width[column] = width

    def set_column_width_dict(self, col_width_dict):
        """
        :param col_width_dict: dictionary from tuple of columns to new width
        :return:
        """
        for cols, width in col_width_dict.iteritems():
            self.set_column_width(cols, width)

    def set_row_height(self, rows, height):
        """
        set the height of the given rows
        :param rows: a single row index, list of indexes or tuple of indexes to change their height
        :param height: numeric positive value of the new height
        :return:
        """
        if not isinstance(rows, (set, list, tuple)):
            rows = [rows]
        try:
            height = float(height)
        except TypeError:
            raise TypeError('rows height must be numeric value')

        if height <= 0:
            raise ValueError('rows width must be positive')
        for row in rows:
            try:
                row = int(row)
            except TypeError:
                raise TypeError("row must be an index")

            self.rows_height[row] = height

    def set_row_height_dict(self, row_height_dict):
        """
        :param row_height_dict: dictionary from tuple of rows to new height
        :return:
        """
        for rows, width in row_height_dict.iteritems():
            self.set_row_height(rows, width)

    def rename(self, columns=None, inplace=False):
        """
        rename the underlying dataframe's columns
        :param columns: a dictionary, old_col_name -> new_col_name
        :param inplace: whether to rename the columns inplace or return a new StyleFrame object
        :return: None if inplace=True, StyleFrame if inplace=False
        """
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
