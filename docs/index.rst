StyleFrame
----------

A library that wraps pandas and openpyxl and allows easy styling of dataframes in excel.

Some usage examples:
====================

StyleFrame constructor supports all the ways you are used to initiate pandas dataframe.
An existing dataframe, a dictionary or a list of dictionaries:
::
    from StyleFrame import StyleFrame

    sf = StyleFrame({'col_a': range(100)})


Applying a style to rows that meet a condition using pandas selecting syntax.
In this example all the cells in the `col_a` column with the value > 50 will have
blue background and a bold, sized 10 font:
::
    sf.apply_style_by_indexes(indexes_to_color=sf[sf['col_a'] > 50],
                              cols_to_style=['col_a'], bg_color='blue', bold=True, font_size=10)

Creating ExcelWriter used to save the excel:
::
    ew = StyleFrame.ExcelWriter(r'C:\my_excel.xlsx')
    sf.to_excel(ew)
    ew.save()

It is also possible to style a whole column or columns, and decide whether to style the headers or not:
::
    sf.apply_column_style(cols_to_style=['a'], bg_color='green', style_header=True)


API:
====
Given that `sf = StyleFrame(...)`:
::
    sf.apply_style_by_indexes(indexes_to_style=None, cols_to_style=None, bg_color=colors.white, bold=False,
                           font_size=12, font_color=colors.black, number_format=number_formats.general):

Applies a certain style to the provided indexes in the dataframe to the provided columns.
Parameters:
::
    indexes_to_style: indexes to apply the style to
    cols_to_style: the columns to apply the style to, if not provided all the columns will be styled
    bg_color: the cell's background color to use
    bold: bold or not
    font_size: the font size
    font_color: the font color
    number_format: Excel's number format to use


::
    sf.apply_column_style(cols_to_style=None, bg_color=colors.white, bold=False, font_size=12, font_color=colors.black,
                          style_header=False, number_format=number_formats.general):

Apply a style to a whole column.
Parameters:
::
    cols_to_style: the columns to apply the style to
    bg_color: the cell's background color to use
    bold: bold or not
    font_size: the font size
    font_color: the font color
    style_header: style the header or not
    number_format: Excel's number format to use


::
    sf.apply_headers_style(bg_color=colors.white, bold=True, font_size=12, font_color=colors.black,
                           number_format=number_formats.general):


Apply style to the headers only.
Parameters:
::
        bg_color: the cell's background color to use
        bold: bold or not
        font_size: the font size
        font_color: the font color
        style_header: style the header or not
        number_format: Excel's number format to use

::
        sf.rename(columns=None, inplace=False):

Rename the underlying dataframe's columns.
Parameters:
::
        columns: a dictionary, old_col_name -> new_col_name
        inplace: whether to rename the columns inplace or return a new StyleFrame object
        return: None if inplace=True, StyleFrame if inplace=False
