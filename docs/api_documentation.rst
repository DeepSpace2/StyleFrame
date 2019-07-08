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
          fill_pattern_type=utils.fill_pattern_types.solid, indent=0, comment_author=None, comment_text=None,
          text_rotation=0)

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
:text_rotation: (int) 1 - 180

Methods
*******

.. py:function:: combine(styles)

    A classmethod used to combine :ref:`styler-class` objects. The right-most object has precedence.
    For example:

    ::

        Styler.combine(Styler(bg_color='yellow', font_size=24), Styler(bg_color='blue'))

    will return

    ::

        Styler(bg_color='blue', font_size=24)

    :param styles: Iterable of Styler objects
    :type styles: list or tuple or set
    :return: self
    :rtype: :ref:`styler-class`

.. py:function:: to_openpyxl_style

    :return: `openpyxl` style object.

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

.. _apply_style_by_indexes_:

.. py:function:: apply_style_by_indexes(indexes_to_style, styler_obj, cols_to_style=None, height=None, complement_style=None, complement_height=None, overwrite_default_style=True)

    :param indexes_to_style: The StyleFrame indexes to style. Usually passed as pandas selecting syntax.
                      For example, ``sf[sf['some_col'] = 20]``
    :type indexes_to_style: list or tuple or int or Container
    :param styler_obj: `Styler` object that contains the style which will be applied to indexes in `indexes_to_style`
    :type styler_obj: :ref:`styler-class`
    :param cols_to_style: The column names to apply the provided style to. If ``None`` all columns will be styled.
    :type cols_to_style: None or str or list[str] or tuple[str] or set[str]
    :param height: If provided, height for rows whose indexes are in indexes_to_style.
    :type height: None or int or float
    :param complement_style: `Styler` object that contains the style which will be applied to indexes not in `indexes_to_style`
    :type complement_style: None or :ref:`styler-class`
    :param complement_height: Height for rows whose indexes are not in indexes_to_style. If not provided then
            `height` will be used (if provided).
    :type complement_height: None or int or float
    :param bool overwrite_default_style: If `True`, the default style (the style used when initializing StyleFrame)
            will be overwritten. If `False` then the default style and the provided style wil be combined using
            Styler.combine method.
    :return: self
    :rtype: StyleFrame

.. py:function:: apply_column_style(cols_to_style, styler_obj, style_header=False, use_default_formats=True, width=None, overwrite_default_style=True)

    :param cols_to_style: The column names to style.
    :type cols_to_style: str or list or tuple or set
    :param styler_obj: A `Styler` object.
    :type styler_obj: (:ref:`styler-class`)
    :param bool style_header: If `True`, the column(s) header will also be styled.
    :param bool use_default_formats: If `True`, the default formats for date and times will be used.
    :param width: If provided, the new width for the specified columns.
    :type width: None or int or float
    :param bool overwrite_default_style: (bool) If `True`, the default style (the style used when initializing StyleFrame)
            will be overwritten. If `False` then the default style and the provided style wil be combined using
            Styler.combine method.
    :return: self
    :rtype: StyleFrame

.. py:function:: apply_headers_style(styler_obj, style_index_header)

    :param styler_obj: A `Styler` object.
    :type styler_obj: :ref:`styler-class`
    :param bool style_index_header: If True then the style will also be applied to the header of the index column
    :return: self
    :rtype: StyleFrame

.. py:function:: style_alternate_rows(styles)

    .. note:: ``style_alternate_rows`` also accepts all arguments that :ref:`StyleFrame.apply_style_by_indexes <apply_style_by_indexes_>` accepts as kwargs.

    :param styles: List, tuple or set of :ref:`styler-class` objects to be applied to rows in an alternating manner
    :type styles: list[:ref:`styler-class`] or tuple[:ref:`styler-class`] or set[:ref:`styler-class`]
    :return: self
    :rtype: StyleFrame

.. py:function:: rename(columns, inplace=False)

    :param dict columns: A dictionary from old columns names to new columns names.
    :param bool inplace: If `False`, a new StyleFrame object will be returned. If `True`, renames the columns inplace.
    :return: self if inplace is `True`, new StyleFrame object is `False`
    :rtype: StyleFrame

.. py:function:: set_column_width(columns, width)

    :param columns: Column name(s).
    :type columns: str or list[str] or tuple[str]
    :param width: The new width for the specified columns.
    :type width: int or float
    :return: self
    :rtype: StyleFrame

.. py:function:: set_column_width_dict(col_width_dict)

    :param col_width_dict: A dictionary from column names to width.
    :type col_width_dict: dict[str, int or float]
    :return: self
    :rtype: StyleFrame

.. py:function:: set_row_height(rows, height)

    :param rows: Row(s) index.
    :type rows: int or list[int] or tuple[int] or set[int]
    :param height: The new height for the specified indexes.
    :type height: int or float
    :return: self
    :rtype: StyleFrame

.. py:function:: set_row_height_dict(row_height_dict)

    :param row_height_dict: A dictionary from row indexes to height.
    :type row_height_dict: dict[int, int or float]
    :return: self
    :rtype: StyleFrame

.. py:function:: add_color_scale_conditional_formatting(start_type, start_value, start_color, end_type, end_value, end_color, mid_type=None, mid_value=None, mid_color=None, columns_range=None)

    :param start_type: The type for the minimum bound
    :type start_type: str: one of :ref:`utils.conditional_formatting_types` or any other type Excel supports
    :param start_value: The threshold for the minimum bound
    :param start_color: The color for the minimum bound
    :type start_color: str: one of :ref:`utils.colors`, hex string or color name ie `'yellow'` Excel supports
    :param end_type: The type for the maximum bound
    :type end_type: str: one of :ref:`utils.conditional_formatting_types` or any other type Excel supports
    :param end_value: The threshold for the maximum bound
    :param end_color: The color for the maximum bound
    :type end_color: str: one of :ref:`utils.colors`, hex string or color name ie `'yellow'` Excel supports
    :param mid_type: The type for the middle bound
    :type mid_type: None or str: one of :ref:`utils.conditional_formatting_types` or any other type Excel supports
    :param mid_value: The threshold for the middle bound
    :param mid_color: The color for the middle bound
    :type mid_color: None or str: one of :ref:`utils.colors`, hex string or color name ie `'yellow'` Excel supports
    :param columns_range: A two-elements list or tuple of columns to which the conditional formatting will be added
            to.
            If not provided at all the conditional formatting will be added to all columns.
            If a single element is provided then the conditional formatting will be added to the provided column.
            If two elements are provided then the conditional formatting will start in the first column and end in the second.
            The provided columns can be a column name, letter or index.
    :type columns_range: None or list[str or int] or tuple[str or int])
    :return: self
    :rtype: StyleFrame

.. py:function:: read_excel(path, sheet_name=0, read_style=False, use_openpyxl_styles=False, read_comments=False)

    A classmethod used to create a StyleFrame object from an existing Excel.

    .. note:: ``read_excel`` also accepts all arguments that ``pandas.read_excel`` accepts as kwargs.

    :param str path: The path to the Excel file to read.
    :param sheetname:
          .. deprecated:: 1.6
             Use ``sheet_name`` instead.
    :param sheet_name: The sheet name to read. If an integer is provided then it be used as a zero-based
            sheet index. Default is 0.
    :type sheet_name: str or int
    :param bool read_style: If `True` the sheet's style will be loaded to the returned StyleFrame object.
    :param bool use_openpyxl_styles: If `True` (and `read_style` is also `True`) then the styles in the returned
        StyleFrame object will be Openpyxl's style objects. If `False`, the styles will be :ref:`styler-class` objects.

        .. note:: Using ``use_openpyxl_styles=False`` is useful if you are going to filter columns or rows by style, for example:

                 ::

                    sf = sf[[col for col in sf.columns if col.style.font == utils.fonts.arial]]

    :param bool read_comments: If `True` (and `read_style` is also `True`) cells' comments will be loaded to the returned StyleFrame object. Note
            that reading comments without reading styles is currently not supported.

    :return: StyleFrame object
    :rtype: StyleFrame

.. py:function:: to_excel(excel_writer='output.xlsx', sheet_name='Sheet1', allow_protection=False, right_to_left=False, columns_to_hide=None, row_to_add_filters=None, columns_and_rows_to_freeze=None, best_fit=None)

    .. note:: ``to_excel`` also accepts all arguments that ``pandas.DataFrame.to_excel`` accepts as kwargs.

    :param excel_writer: File path or existing ExcelWriter
    :type excel_writer: str or pandas.ExcelWriter
    :param str sheet_name: Name of sheet the StyleFrame will be exported to
    :param bool allow_protection: Allow to protect the cells that specified as protected. If used ``protection=True``
        in a Styler object this must be set to `True`.
    :param bool right_to_lef: Makes the sheet right-to-left.
    :param columns_to_hide: Columns names to hide.
    :type columns_to_hide: None or str or list or tuple or set
    :param row_to_add_filters: Add filters to the given row index, starts from 0 (which will add filters to header row).
    :type row_to_add_filters: None or int
    :param columns_and_rows_to_freeze: Column and row string to freeze.
        For example "C3" will freeze columns: A, B and rows: 1, 2.
    :type columns_and_rows_to_freeze: None or str
    :param best_fit: single column, list, set or tuple of columns names to attempt to best fit the width for.

        .. note:: ``best_fit`` will attempt to calculate the correct column-width based on the longest value in each provided
                  column. However this isn't guaranteed to work for all fonts (works best with monospaced fonts). The formula
                  used to calculate a column's width is equivalent to

                  ::

                    (len(longest_value_in_column) + A_FACTOR) * P_FACTOR

                  The default values for ``A_FACTOR`` and ``P_FACTOR`` are 13 and 1.3 respectively, and can be modified before
                  calling ``StyleFrame.to_excel`` by directly modifying ``StyleFrame.A_FACTOR`` and ``StyleFrame.P_FACTOR``

    :type best_fit: None or str or list or tuple or set
    :return: self
    :rtype: StyleFrame
