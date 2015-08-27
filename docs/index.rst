StyleFrame
----------

A library that wraps pandas and openpyxl and allows easy styling of dataframes in excel.

Some usage examples:
===================

StyleFrame constructor supports all the ways you are used to initiate pandas dataframe.
An existing dataframe, a dictionary or a list of dictionaries:
::
    from StyleFrame import StyleFrame

    sf = StyleFrame({'col_a': range(100)})


Applying a style to rows that meet a condition using pandas selecting syntax.
In this example all the cells in the `col_a` column with the value > 50 will have
blue background and a bold, sized 10 font.::

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

