API Documentation
=================

Given that `sf = StyleFrame(...)` :

Styling by indexes
------------------
::

    sf.apply_style_by_indexes(indexes_to_style=None, cols_to_style=None,
                              styler_obj=Styler(bg_color=utils.colors.white,
                              bold=False, font_size=12, font_color=utils.colors.black,
                              number_format=utils.number_formats.general,
                              protection=False), height=None)

Applies a certain style to the provided indexes in the dataframe to the provided columns.
Parameters:
::

    indexes_to_style: indexes to apply the style to
    cols_to_style: the columns to apply the style to, if not provided all the columns will be styled
    styler_obj: a StyleFrame.Styler object
    height: if None, use default excel height otherwise use the given height value for the chosen rows

Styling by columns
------------------
::

    sf.apply_column_style(cols_to_style=None,
                          styler_obj=Styler(bg_color=utils.colors.white, bold=False, font_size=12,
                          font_color=utils.colors.black, style_header=False,
                          number_format=utils.number_formats.general,
                          protection=False), use_default_formats=True, width=None)

Apply a style to a whole column.
Parameters:
::

    cols_to_style: the columns to apply the style to
    styler_obj: a StyleFrame.Styler object
    use_default_formats: if True, use predefined styles for dates and times
    width: if None, use default excel width otherwise use the given width value for the chosen columns

Styling headers only
--------------------
::

    sf.apply_headers_style(styler_obj=Styler(bg_color=colors.white, bold=True, font_size=12, font_color=utils.colors.black,
                           number_format=utils.number_formats.general, protection=False))

Apply style to the headers only.
Parameters:
::

        styler_obj: a StyleFrame.Styler object

Renaming columns
----------------
::

        sf.rename(columns=None, inplace=False)

Rename the underlying dataframe's columns.
Parameters:
::

        columns: a dictionary, old_col_name -> new_col_name
        inplace: whether to rename the columns inplace or return a new StyleFrame object
        return: None if inplace=True, StyleFrame if inplace=False

Setting columns width
---------------------
::

    sf.set_column_width(columns, width)

Set the width of the given columns
Parameters:
::

        columns: a single or a list/tuple of column name, index or letter to change their width
        width: numeric positive value of the new width

::

    sf.set_column_width_dict(self, col_width_dict)

Parameters:
::

        col_width_dict: a dictionary from tuples of columns to the desired width

Setting rows height
-------------------
::

    sf.set_row_height(rows, height)

Set the height of the given rows.
Parameters:
::

        rows: a single row index, list of indexes or tuple of indexes to change their height
        height: numeric positive value of the new height

::

    sf.set_row_height_dict(self, row_height_dict)

Parameters:
::

    row_height_dict: a dictionary from tuples of rows to the desired height

Reading existing Excel file
---------------------------
::

    sf.read_excel(path, sheetname='Sheet1', read_style=False, **kwargs)

Reads an Excel file and returns a StyleFrame object
Parameters:
::

    path: file's path
    sheetname: the sheetname to read from
    read_style: if True the returned StyleFrame object will have the same style the Excel sheet has
    **kwargs: the same kwargs and pandas.read_excel expects

