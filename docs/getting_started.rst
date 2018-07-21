Getting Started
===============

Basic Usage
^^^^^^^^^^^

StyleFrame's ``init`` supports all the ways you are used to initiate pandas dataframe.
An existing dataframe, a dictionary or a list of dictionaries:
::

    from StyleFrame import StyleFrame, Styler, utils

    sf = StyleFrame({'col_a': range(100)})

Applying a style to rows that meet a condition using pandas selecting syntax.
In this example all the cells in the `col_a` column with the value > 50 will have
blue background and a bold, sized 10 font:
::


    sf.apply_style_by_indexes(indexes_to_style=sf[sf['col_a'] > 50],
                              cols_to_style=['col_a'],
                              styler_obj=Styler(bg_color=utils.colors.blue, bold=True, font_size=10))

Creating ExcelWriter used to export StyleFrame to Excel:
::

    ew = StyleFrame.ExcelWriter(r'C:\my_excel.xlsx')
    sf.to_excel(ew)
    ew.save()

It is also possible to style a whole column or columns, and decide whether to style the headers or not:
::

    sf.apply_column_style(cols_to_style=['a'], styler_obj=Styler(bg_color=utils.colors.green),
                          style_header=True)

Example with 'real world' data
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. note:: The data used in this example is the first 500 rows of of Kaggle's
          "StackLite: Stack Overflow questions and tags" dataset available at https://www.kaggle.com/stackoverflow/stacklite

.. note:: These examples are focusing on StyleFrame's styling abilities rather than on pandas data mingling
          abilities (and the subset of these abilities that is available with StyleFrame)

Setting up

::

    import pandas as pd
    from datetime import timedelta
    from StyleFrame import StyleFrame, Styler, utils

    df = pd.read_csv('data.csv', parse_dates=['CreationDate', 'ClosedDate', 'DeletionDate'])
    sf = StyleFrame(df)

Using red background for Id column for rows with questions that were closed less than 5 minutes after creation
::

    sf.apply_style_by_indexes(indexes_to_style=sf[sf['ClosedDate'] - sf['CreationDate'] < timedelta(minutes=5)],
                              styler_obj=Styler(bg_color=utils.colors.red),
                              cols_to_style=['Id'])

Changing the width of the date columns so their content fits nicely
::

    sf.set_column_width(columns=['CreationDate', 'ClosedDate', 'DeletionDate'],
                        width=20)

Using color-scale conditional formatting for the questions' scores, based on percentage
::

    sf.add_color_scale_conditional_formatting(start_type=utils.conditional_formatting_types.percentile,
                                              start_value=0,
                                              start_color=utils.colors.red,
                                              end_type=utils.conditional_formatting_types.percentile,
                                              end_value=100,
                                              end_color=utils.colors.green,
                                              columns_range=['Score'])


Adding filters to the header row, freezing it and exporting to Excel
::

    sf.to_excel('output.xlsx', columns_and_rows_to_freeze='A2', row_to_add_filters=0,
                best_fit=['OwnerUserId', 'AnswerCount']).save()


Entire code

::

    import pandas as pd
    from datetime import timedelta
    from StyleFrame import StyleFrame, Styler, utils

    # data.csv contains the first 500 rows of Kaggle's "StackLite: Stack Overflow questions and tags"
    # dataset available at https://www.kaggle.com/stackoverflow/stacklite
    df = pd.read_csv('data.csv', parse_dates=['CreationDate', 'ClosedDate', 'DeletionDate'])

    sf = StyleFrame(df)

    sf.apply_style_by_indexes(indexes_to_style=sf[sf['ClosedDate'] - sf['CreationDate'] < timedelta(minutes=5)],
                              styler_obj=Styler(bg_color=utils.colors.red),
                              cols_to_style=['Id'])

    sf.set_column_width(columns=['CreationDate', 'ClosedDate', 'DeletionDate'],
                        width=20)

    sf.add_color_scale_conditional_formatting(start_type=utils.conditional_formatting_types.percentile,
                                              start_value=0,
                                              start_color=utils.colors.red,
                                              end_type=utils.conditional_formatting_types.percentile,
                                              end_value=100,
                                              end_color=utils.colors.green,
                                              columns_range=['Score'])

    sf.to_excel('output.xlsx', columns_and_rows_to_freeze='A2', row_to_add_filters=0,
                best_fit=['OwnerUserId', 'AnswerCount']).save()
