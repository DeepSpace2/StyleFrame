API Documentation
=================

utils
-----

The utils module contains the most widely used values for styling elements such as colors and border types for convenience.
It is possible to directly use a value that is not present in the utils module as long as Excel recognises it.

utils.number_formats
^^^^^^^^^^^^^^^^^^^^
::

   general = 'General'
   general_integer = '0'
   general_float = '0.00'
   percent = '0.0%'
   thousands_comma_sep = '#,##0'
   date = 'DD/MM/YY'
   time_24_hours = 'HH:MM'
   time_12_hours = 'h:MM AM/PM'
   date_time = 'DD/MM/YY HH:MM'
   date_time_with_seconds = 'DD/MM/YY HH:MM:SS'


utils.colors
^^^^^^^^^^^^
::

   white = op_colors.WHITE
   blue = op_colors.BLUE
   dark_blue = op_colors.DARKBLUE
   yellow = op_colors.YELLOW
   dark_yellow = op_colors.DARKYELLOW
   green = op_colors.GREEN
   dark_green = op_colors.DARKGREEN
   black = op_colors.BLACK
   red = op_colors.RED
   dark_red = op_colors.DARKRED
   purple = '800080'
   grey = 'D3D3D3'


utils.fonts
^^^^^^^^^^^
::

   aegean = 'Aegean'
   aegyptus = 'Aegyptus'
   aharoni = 'Aharoni CLM'
   anaktoria = 'Anaktoria'
   analecta = 'Analecta'
   anatolian = 'Anatolian'
   arial = 'Arial'
   calibri = 'Calibri'
   david = 'David CLM'
   dejavu_sans = 'DejaVu Sans'
   ellinia = 'Ellinia CLM'


utils.borders
^^^^^^^^^^^^^
::

   dash_dot = 'dashDot'
   dash_dot_dot = 'dashDotDot'
   dashed = 'dashed'
   dotted = 'dotted'
   double = 'double'
   hair = 'hair'
   medium = 'medium'
   medium_dash_dot = 'mediumDashDot'
   medium_dash_dot_dot = 'mediumDashDotDot'
   medium_dashed = 'mediumDashed'
   slant_dash_dot = 'slantDashDot'
   thick = 'thick'
   thin = 'thin'


utils.underline
^^^^^^^^^^^^^^^
::

   single = 'single'
   double = 'double'

Styler Class
------------

Used to represent a style.

Init Arguments
^^^^^^^^^^^^^^
::

   Styler(bg_color=None, bold=False, font=utils.fonts.arial, font_size=12, font_color=None,
          number_format=utils.number_formats.general, protection=False, underline=None,
          border_type=utils.borders.thin)

:bg_color: (str: hex string or color name ie `'yellow'` | utils.color) The background color
:bold: (bool) If true, a bold typeface is used
:font: (str: font name | utils.font) The font to use
:font_size: (int) The font size
:font_color: (str: hex string or color name ie `'yellow'` | utils.color) The font color
:number_format: (str: utils.font_number or any other format Excel supports) The format of the cell's value
:protection: (bool) If true, the cell/column will be write-protected
:underline: (str: utils.underline or any other underline Excel supports) The underline type
:border_type: (str: utils.border_type or any other border type Excel supports) The border type

Methods
^^^^^^^

create_style
""""""""""""

:arguments: None
:returns: `openpyxl` style object.


StyleFrame Class
----------------

Represent a stylized dataframe

Init Arguments
^^^^^^^^^^^^^^
::

   StyleFrame(obj, styler_obj=None)

:obj: Any object that pandas' dataframe can be initialized with: an existing dataframe, a dictionary,
      a list of dictionaries or another StylerFrame.
:styler_obj: (Styler) A Styler object. Will be used as the default style of all cells.

Methods
^^^^^^^

apply_style_by_indexes
""""""""""""""""""""""

:arguments:
   :indexes_to_style: The StyleFrame indexes to style. This usually passed as pandas selecting syntax.
                      For example, ``sf[sf['some_col'] = 20``
   :styler_obj: (Styler) The `Styler` object that represent the style
   :cols_to_style=None: (str | list | tuple) The column names to apply the provided style to. If ``None`` all columns will be styled.
   :height=None: (int) If provided, the new height for the matched indexes.
:returns: self

apply_column_style
""""""""""""""""""

:arguments:
   :cols_to_style: (str | list | tuple) The column names to style.
   :styler_obj: (Styler) A `Styler` object.
   :style_header=False: (bool) If True, the column(s) header will also be styled.
   :use_default_formats=True: (bool) If True, the default formats for date and times will be used.
   :width=None: (int) If provided, the new width for the specified columns.
:returns: self

apply_headers_style
"""""""""""""""""""

:arguments:
   :styler_obj: (Styler) A `Styler` object.
:returns: self

rename
""""""

:arguments:
   :columns=None: (dict) A dictionary from old columns names to new columns names.
   :inplace=False: (bool) If False, a new StyleFrame object will be returned. If True, renames the columns inplace.
:returns: self if inplace is `True`, new StyleFrame object is `False`

set_column_width
""""""""""""""""

:arguments:
    :columns: (str | list| tuple) Column name(s).
    :width: (int) The new width for the specified columns.
:returns: self

set_column_width_dict
"""""""""""""""""""""

:arguments:
   :col_width_dict: (dict) A dictionary from column names to width.
:returns: self

set_row_height
""""""""""""""

:arguments:
   :rows: (int| list | tuple) Row(s) index.
   :height: (int) The new height for the specified indexes.
:returns: self

set_row_height_dict
"""""""""""""""""""

:arguments:
    :row_height_dict: (dict) A dictionary from row indexes to height.
:returns: self

read_excel
""""""""""

:arguments:
   :path: (str) The path to the Excel file to read.
   :sheetname: (str) The sheet name to read from.
   :read_style=False: (bool) If `True` the sheet's style will be loaded to the returned StyleFrame object.
   :kwargs: Any keyword argument pandas' `read_excel` supports.
:returns: StyleFrame object

A classmethod used to create a StyleFrame object from an existing Excel.

to_excel
========

:arguments:
   :allow_protection=False: (bool) Allow to protect the cells that specified as protected. If used ``protection=True``
                             in a Styler object this must be set to `True`.
   :right_to_left=False: (bool) Makes the sheet right-to-left.
   :columns_to_hide=None: (str | list | tuple) Columns names to hide.
   :row_to_add_filters=None: (int) Add filters to the given row index, starts from 0 (which will add filters to header row).
   :columns_and_rows_to_freeze=None: (str) Column and row string to freeze.
                                     For example "C3" will freeze columns: A, B and rows: 1, 2.
:returns: self