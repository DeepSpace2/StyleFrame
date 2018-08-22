API Documentation
=================

utils module
------------

This module contains the most widely used values for styling elements such as colors and border types for convenience.
It is possible to directly use a value that is not present in the utils module as long as Excel recognises it.

.. _utils.number_formats:

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
   time_24_hours_with_seconds = 'HH:MM:SS'
   time_12_hours = 'h:MM AM/PM'
   time_12_hours_with_seconds = 'h:MM:SS AM/PM'
   date_time = 'DD/MM/YY HH:MM'
   date_time_with_seconds = 'DD/MM/YY HH:MM:SS'

decimal_with_num_of_digits
**************************

:arguments:
  :num_of_digits: (int) Number of digits after the decimal point
:returns: A format string that represents a floating point number with the provided number of digits after the
          decimal point. For example, ``utils.number_formats.decimal_with_num_of_digits(2)`` will return ``'0.00'``

.. _utils.colors:

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


.. _utils.fonts:

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


.. _utils.borders:

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


.. _utils.horizontal_alignments:

utils.horizontal_alignments
^^^^^^^^^^^^^^^^^^^^^^^^^^^
::

    general = 'general'
    left = 'left'
    center = 'center'
    right = 'right'
    fill = 'fill'
    justify = 'justify'
    center_continuous = 'centerContinuous'
    distributed = 'distributed'


.. _utils.vertical_alignments:

utils.vertical_alignments
^^^^^^^^^^^^^^^^^^^^^^^^^
::

    top = 'top'
    center = 'center'
    bottom = 'bottom'
    justify = 'justify'
    distributed = 'distributed'


.. _utils.underline:

utils.underline
^^^^^^^^^^^^^^^
::

   single = 'single'
   double = 'double'


.. _utils.fill_pattern_types:

utils.fill_pattern_types
^^^^^^^^^^^^^^^^^^^^^^^^
::

  solid = 'solid'
  dark_down = 'darkDown'
  dark_gray = 'darkGray'
  dark_grid = 'darkGrid'
  dark_horizontal = 'darkHorizontal'
  dark_trellis = 'darkTrellis'
  dark_up = 'darkUp'
  dark_vertical = 'darkVertical'
  gray0625 = 'gray0625'
  gray125 = 'gray125'
  light_down = 'lightDown'
  light_gray = 'lightGray'
  light_grid = 'lightGrid'
  light_horizontal = 'lightHorizontal'
  light_trellis = 'lightTrellis'
  light_up = 'lightUp'
  light_vertical = 'lightVertical'
  medium_gray = 'mediumGray'


.. _utils.conditional_formatting_types:

utils.conditional_formatting_types
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
::

    num = 'num'
    percent = 'percent'
    max = 'max'
    min = 'min'
    formula = 'formula'
    percentile = 'percentile'


styler module
-------------

This module contains classes that represent styles.

.. _styler-class:

Styler Class
^^^^^^^^^^^^

Used to represent a style.

::

   Styler(bg_color=None, bold=False, font=utils.fonts.arial, font_size=12, font_color=None,
          number_format=utils.number_formats.general, protection=False, underline=None,
          border_type=utils.borders.thin, horizontal_alignment=utils.horizontal_alignments.center,
          vertical_alignment=utils.vertical_alignments.center, wrap_text=True, shrink_to_fit=True,
          fill_pattern_type=utils.fill_pattern_types.solid, indent=0, comment_author=None, comment_text=None)

:bg_color: (str: one of :ref:`utils.colors`, hex string or color name ie `'yellow'` Excel supports) The background color
:bold: (bool) If `True`, a bold typeface is used
:font: (str: one of :ref:`utils.fonts` or other font name Excel supports) The font to use
:font_size: (int) The font size
:font_color: (str: one of :ref:`utils.colors`, hex string or color name ie `'yellow'` Excel supports) The font color
:number_format: (str: one of :ref:`utils.number_formats` or any other format Excel supports) The format of the cell's value
:protection: (bool) If `True`, the cell/column will be write-protected
:underline: (str: one of :ref:`utils.underline` or any other underline Excel supports) The underline type
:border_type: (str: one of :ref:`utils.borders` or any other border type Excel supports) The border type
:horizontal_alignment: (str: one of :ref:`utils.horizontal_alignments` or any other horizontal alignment Excel supports) Text's horizontal alignment
:vertical_alignment: (str: one of :ref:`utils.vertical_alignments` or any other vertical alignment Excel supports) Text's vertical alignment
:wrap_text: (bool)
:shrink_to_fit: (bool)
:fill_pattern_type: (str: one of :ref:`utils.fill_pattern_types` or any other fill pattern type Excel supports) Cells's fill pattern type
:indent: (int)
:comment_author: (str)
:comment_text: (str)

Methods
*******

combine
"""""""
A classmethod used to combine :ref:`styler-class` objects. The right-most object has precedence.
For example: ``Styler.combine(Styler(bg_color='yellow', font_size=24), Styler(bg_color='blue'))`` will return
``Styler(bg_color='blue', font_size=24)``

:arguments:
  :styles: Arbitrary number of Styler objects
:returns: :ref:`styler-class` object


to_openpyxl_style
"""""""""""""""""

:arguments: None
:returns: `openpyxl` style object.

style_frame module
------------------

StyleFrame Class
^^^^^^^^^^^^^^^^

Represent a stylized dataframe

::

   StyleFrame(obj, styler_obj=None)

:obj: Any object that pandas' dataframe can be initialized with: an existing dataframe, a dictionary,
      a list of dictionaries or another StyleFrame.
:styler_obj: (:ref:`styler-class`) A Styler object. Will be used as the default style of all cells.

Methods
*******

apply_style_by_indexes
""""""""""""""""""""""

:arguments:
   :indexes_to_style: (list | tuple | int | Container) The StyleFrame indexes to style. This usually passed as pandas selecting syntax.
                      For example, ``sf[sf['some_col'] = 20]``
   :styler_obj: (:ref:`styler-class`) `Styler` object that contains the style which will be applied to indexes in `indexes_to_style`
   :cols_to_style=None: (None | str | list | tuple | set) The column names to apply the provided style to. If ``None`` all columns will be styled.
   :height=None: (None | int | float) If provided, height for rows whose indexes are in indexes_to_style.
   :complement_style=None: (None | :ref:`styler-class`) `Styler` object that contains the style which will be applied to indexes not in `indexes_to_style`
   :complement_height=None: (None | int | float)  Height for rows whose indexes are not in indexes_to_style. If not provided then
            `height` will be used (if provided).
   :overwrite_default_style=True: (bool) If `True`, the default style (the style used when initializing StyleFrame)
            will be overwritten. If `False` then the default style and the provided style wil be combined using
            Styler.combine method.
:returns: self

apply_column_style
""""""""""""""""""

:arguments:
   :cols_to_style: (str | list | tuple | set) The column names to style.
   :styler_obj: (:ref:`styler-class`) A `Styler` object.
   :style_header=False: (bool) If `True`, the column(s) header will also be styled.
   :use_default_formats=True: (bool) If `True`, the default formats for date and times will be used.
   :width=None: (None | int | float) If provided, the new width for the specified columns.
   :overwrite_default_style=True: (bool) If `True`, the default style (the style used when initializing StyleFrame)
            will be overwritten. If `False` then the default style and the provided style wil be combined using
            Styler.combine method.
:returns: self

apply_headers_style
"""""""""""""""""""

:arguments:
   :styler_obj: (:ref:`styler-class`) A `Styler` object.
   :style_index_header: (bool) If True then the style will also be applied to the header of the index column
:returns: self

style_alternate_rows
""""""""""""""""""""

:arguments:
   :styles: (list | tuple | set) List or tuple of :ref:`styler-class` objects to be applied to rows in an alternating manner
:returns: self

rename
""""""

:arguments:
   :columns=None: (dict) A dictionary from old columns names to new columns names.
   :inplace=False: (bool) If `False`, a new StyleFrame object will be returned. If `True`, renames the columns inplace.
:returns: self if inplace is `True`, new StyleFrame object is `False`

set_column_width
""""""""""""""""

:arguments:
    :columns: (str | list| tuple) Column name(s).
    :width: (int | float) The new width for the specified columns.
:returns: self

set_column_width_dict
"""""""""""""""""""""

:arguments:
   :col_width_dict: (dict) A dictionary from column names to width.
:returns: self

set_row_height
""""""""""""""

:arguments:
   :rows: (int | list | tuple | set) Row(s) index.
   :height: (int | float) The new height for the specified indexes.
:returns: self

set_row_height_dict
"""""""""""""""""""

:arguments:
    :row_height_dict: (dict) A dictionary from row indexes to height.
:returns: self

add_color_scale_conditional_formatting
""""""""""""""""""""""""""""""""""""""

:arguments:

    :start_type: (str: one of :ref:`utils.conditional_formatting_types` or any other type Excel supports) The type for the minimum bound
    :start_value: The threshold for the minimum bound
    :start_color: (str: one of :ref:`utils.colors`, hex string or color name ie `'yellow'` Excel supports) The color for the minimum bound
    :end_type: (str: one of :ref:`utils.conditional_formatting_types` or any other type Excel supports) The type for the maximum bound
    :end_value: The threshold for the maximum bound
    :end_color: (str: one of :ref:`utils.colors`, hex string or color name ie `'yellow'` Excel supports) The color for the maximum bound
    :mid_type=None: (None | str: one of :ref:`utils.conditional_formatting_types` or any other type Excel supports) The type for the middle bound
    :mid_value=None: The threshold for the middle bound
    :mid_color=None: (None | str: one of :ref:`utils.colors`, hex string or color name ie `'yellow'` Excel supports) The color for the middle bound
    :columns_range=None: (None | list | tuple) A two-elements list or tuple of columns to which the conditional formatting will be added
            to.
            If not provided at all the conditional formatting will be added to all columns.
            If a single element is provided then the conditional formatting will be added to the provided column.
            If two elements are provided then the conditional formatting will start in the first column and end in the second.
            The provided columns can be a column name, letter or index.
:returns: self

read_excel
""""""""""

A classmethod used to create a StyleFrame object from an existing Excel.

.. note:: ``read_excel`` also accepts all arguments that ``pandas.read_excel`` accepts as kwargs.

:arguments:
   :path: (str) The path to the Excel file to read.
   :sheetname:

      .. deprecated:: 1.6
         Use ``sheet_name`` instead.

   :sheet_name=0: (str | int) The sheet name to read. If an integer is provided then it be used as a zero-based
            sheet index. Default is 0.
   :read_style=False: (bool) If `True` the sheet's style will be loaded to the returned StyleFrame object.
   :use_openpyxl_styles=True: (bool) If `True` (and `read_style` is also `True`) then the styles in the returned
            StyleFrame object will be Openpyxl's style objects. If `False`, the styles will be :ref:`styler-class` objects.
            Defaults to `True` for backward compatibility.

   .. note:: Using ``use_openpyxl_styles=False`` is useful if you are going to filter columns or rows by style, for example:

             ::

                sf = sf[[col for col in sf.columns if col.style.font == utils.fonts.arial]]

   :read_comments=False: (bool) If `True` (and `read_style` is also `True`) cells' comments will be loaded to the returned StyleFrame object. Note
            that reading comments without reading styles is currently not supported.

:returns: StyleFrame object

to_excel
""""""""

.. note:: ``to_excel`` also accepts all arguments that ``pandas.DataFrame.to_excel`` accepts as kwargs.

:arguments:
   :excel_writer='output.xlsx': (str | pandas.ExcelWriter) File path or existing ExcelWriter
   :sheet_name='Sheet1': (str) Name of sheet the StyleFrame will be exported to
   :allow_protection=False: (bool) Allow to protect the cells that specified as protected. If used ``protection=True``
                             in a Styler object this must be set to `True`.
   :right_to_left=False: (bool) Makes the sheet right-to-left.
   :columns_to_hide=None: (None | str | list | tuple | set) Columns names to hide.
   :row_to_add_filters=None: (None | int) Add filters to the given row index, starts from 0 (which will add filters to header row).
   :columns_and_rows_to_freeze=None: (None | str) Column and row string to freeze.
                                     For example "C3" will freeze columns: A, B and rows: 1, 2.
   :best_fit=None: (None | str | list | tuple | set) single column, list, set or tuple of columns names to attempt to best fit the width
                                for.

   .. note:: ``best_fit`` will attempt to calculate the correct column-width based on the longest value in each provided
              column. However this isn't guaranteed to work for all fonts (works best with monospaced fonts). The formula
              used to calculate a column's width is equivalent to

              ::

                (len(longest_value_in_column) + A_FACTOR) * P_FACTOR

              The default values for ``A_FACTOR`` and ``P_FACTOR`` are 13 and 1.3 respectively, and can be modified before
              calling ``StyleFrame.to_excel`` by directly modifying ``StyleFrame.A_FACTOR`` and ``StyleFrame.P_FACTOR``

:returns: self
