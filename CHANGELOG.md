#### 4.0
* **Removed Python 3.4 support**
* **Removed Python 3.5 support**
* **Added Python 3.9 support**
* Allowing `StyleFrame.ExcelWriter` to accept any argument (except for `engine`) that `pandas.ExcelWriter` accepts
* Allowing customizing formats of `date`, `time` and `datetime` objects when creating `Styler` instances
* Fixed rows mis-alignment issue when calling `to_excel` with `header=False` ([GitHub issue #88](https://github.com/DeepSpace2/StyleFrame/issues/88))
* `read_excel` does not accept `sheetname` argument anymore (was deprecated since version 1.6). Use `sheet_name` instead.

#### 3.0.6
* Fixes [GitHub issue #94](https://github.com/DeepSpace2/StyleFrame/issues/94) - Passing `border_type=utils.borders.default_grid` to `Styler`

#### 3.0.5
* Fixes [GitHub issue #81](https://github.com/DeepSpace2/StyleFrame/issues/81) - read excel as template headers height

#### 3.0.4
* Fixed style "shifts" when using `read_style=True` and `header=None` with `StyleFrame.read_excel`. Fixes [GitHub issue #80](https://github.com/DeepSpace2/StyleFrame/issues/80) 

#### 3.0.3
No longer relying on openpyxl's colors definition. Related to [GitHub issue #73](https://github.com/DeepSpace2/StyleFrame/issues/73)

#### 3.0.2
Hotfix release - setting maximum versions for dependencies. Related to [GitHub issue #73](https://github.com/DeepSpace2/StyleFrame/issues/73)

#### 3.0.1
* **Removed Python 2.7 support**
* **Added Python 3.8 support**
* Renamed package name to `styleframe` (all lowercase) in accordance of PEP8
* Added `.style` accessor. This allows for easy selection/indexing based on style, for example:
  `sf.loc[sf['col_name'].style.bg_color == utils.colors.yellow]`
  
  or
  
  `sf.loc[~sf['col_name'].style.bold]`
  
* Added `default_grid` to `utils.borders` to allow usage of the default spreadsheet grid
* Added `read_excel_as_template` method
* Fixed a bug that prevented saving if `read_excel` was used with `use_openpxl_style=True`, see [GitHub issue #67](https://github.com/DeepSpace2/StyleFrame/issues/67)
* Allowing usage of pathlib.Path in `to_excel`, see [GitHub issue #69](https://github.com/DeepSpace2/StyleFrame/issues/69) 
* Added ability to execute the tests from the commandline: `styleframe --test`

#### 2.0.5
* `style_alternate_rows` can accept all arguments that `apply_style_by_indexes` accepts as kwargs.
* Added `cols_to_style` argument to `apply_headers_style`

#### 2.0.4
* Fixed a bug that caused `apply_style_by_indexes` not to work in case the dataframe had non-integer indexes in some cases
* Added support for text rotation

#### 2.0.3
* Fixing pandas dependency for different Python versions, related to #52, #53

#### 2.0.2
* Fixed a "'column' is out of columns range" error when settings a column's width if dataframe has more than 26 columns. 

#### 2.0.1
* Hotfix: Fixed typo in setup.py

#### 2.0
* **Change supported versions of openpyxl to >= 2.5**
* Changed `use_openpyxl_styles` argument default value to `False` in `StyleFrame.read_excel`.

#### 1.6.2
* Fixed an issue that prevented reading files which used certain theme colors.

#### 1.6.1
* Reading columns' widths and rows' heights when passing `read_style=True` to `read_excel`.
* Added support for named indexes.
* Added `style_index_header` argument to `apply_headers_style`.
* Added support for pandas <= 0.23.4

#### 1.6
* **Added Python 3.7 support**
* Added support for pandas <= 0.23.1
* Added dt and str accessors to Series
* Added support for passing an integer (sheet index) as `sheet_name` to `StyleFrame.read_excel`
* `StyleFrame.read_excel` `sheetname` argument changed to `sheet_name`. Using `sheetname` is still allowed but will
  show deprecation warning.
* Added `Styler.combine` method.
* Added `utils.number_formats.decimal_with_num_of_digits` method.
* Added `overwrite_default_style` argument to `StyleFrame` methods `apply_style_by_indexes` and `apply_column_style`


#### 1.5.1
* Fixed a bug where `read_excel` will fail when using `read_style=True` in cases
  where specific themes are used (See [GitHub issue #37](https://github.com/DeepSpace2/StyleFrame/issues/37)).

#### 1.5
* Added `complement_style` and `complement_height` arguments to `StyleFrame.apply_style_by_indexes`
* Added support for comments
* Renamed `Styler.create_style` method to `to_openpyxl_style` (`create_style` is still available for backward compatibility)
* Fixed a bug not allowing to access StyleFrame columns by dot notation in case a column name is a number

#### 1.4
* **No longer supporting Python 3.3**
* StyleFrame objects no longer expose .ix method as it is deprecated since pandas 0.20. Use .loc or .iloc instead
* Added ability to access StyleFrame columns as attributes (eg `sf.column_a`)
* Added conditional formatting
* Added `best_fit` to `to_excel` method
* Added support for pandas <= 0.22.0
* Added support for theme colors when reading styles from Excel sheets
* Added option to use `Styler` objects when reading styles from Excel sheets
* Using a JSON schema to validate json from command line
* Added command line argument --show-schema

#### 1.3.1
* Improved error message if invalid style arguments are used in JSON through the commandline interface
* Fixed an error importing utils in case there is already a utils module in Python's path (see [GitHub issue #31](https://github.com/DeepSpace2/StyleFrame/issues/31))

#### 1.3
* Added `utils.fill_pattern_types`
* Added `wrap_text`, `shrink_to_fit`, `fill_pattern_type` and `indent` arguments to `Styler.__init__`

#### 1.2
* Fixed an issue when running tests from code  
* Using `.loc` and `.iloc` instead of `.ix` since `.ix` is deprecated in pandas >0.20 
* Added `horizontal_alignment` and `vertical_alignment` arguments to `Styler.__init__`
* Added `style_alternate_rows` method to `StyleFrame`.

#### 1.1.1
* Added option to pass a json string through cli
* Added `cells` option to `default_styles` when using JSON

#### 1.1
* Added commandline interface that supports json to xlsx
* Added `utils.fonts`
* Added `width` and `height` arguments to relevant styling methods

#### 1.0
* Removed support for individual style specifiers when calling styling methods

#### 0.2.1 
* Added 2 general styles to `utils.number_formats`: '0' as `utils.number_formats.general_integer`
  and '0.00' as `utils.number_formats.general_float`
* Fixed a bug with a `DeprecationWarning` unnecessarily showing

#### 0.2
* Added ability to change font
* Added ability to change cell border type
* Added `style_obj` argument to all styling methods which accepts a `Styler` object so styles can be reused.
  The ability to directly pass style specifiers is deprecated and may break in a future version
* Added basic ability to read stylized excel into a stylized StyleFrame by pass `read_style=True` to `StyleFrame.read_excel`.
  Currently does not support reading style from a subset of sheet (ie using `startrow`, `startcol` and the such)
* Added ability to provide a default `Styler` object to `StyleFrame.__init__`, and added deprecation message when not
  passing a `Styler` object to styling methods

#### 0.1.8
* Added ability to run tests by code:
  ```python
  from StyleFrame import tests
  tests.run()
  ```

#### 0.1.7
* Fixed a bug when adding an underline to a style
* More extensive tests

#### 0.1.6
* Fixed a bug when passing `header=False` to `to_excel`

#### 0.1.5
* Changed dependencies, now requires pandas 0.16.2 - 0.18.1      
* Changed Python support: 2.7, 3.3, 3.4, 3.5
* Fixed a bug when trying to filter the first row
* Transitioning to x.y.z version numbers

#### 0.1.4.2    
* `right_to_left` is now set to `False` as default in `to_excel()`
* Most of the methods now return `self` to allow method chaining (eg `StyleFrame(..).rename(..).to_excel()`)

#### 0.1.4.1
* Supports passing a path to required output file to `to_excel` method, much like pandas's `to_excel`.

#### 0.1.3.5    
* Basic support for Python 3.

#### 0.1.3.2
* Internal changes in `apply_style_by_indexes` method in order to keep number formats of dates and times.

#### 0.1.3.1     
* Fixed a bug when creating a `StyleFrame` from an empty `DataFrame`

#### 0.1.3   
* Some bugs fixes
* Added ability to change width and height of several columns/rows at once (`set_column_width_dict` and `set_row_height` methods)

#### 0.1.2       
* Added ability to create excel with the dataframe's  (and style) indexes

#### 0.1.1    
* Added ability add filter to rows
* Added ability to protect cells and sheets from editing

#### 0.0.9       
* Some bugs fixes

#### 0.0.8   
* Added default style for cells with `'=HYPERLINK(..)'` values (blue color, underlined)       
* Improved unicode support

#### 0.0.7       
* Added ability to set rows height and columns width

#### 0.0.6 
* Added ability to style only the columns headers.         
* Added ability to rename columns while keeping the style of the headers
* Added ability to change font color
* Added `utils.number_formats`       
* Added support for 'direct' item assignment, ie `sf['column_c'] = 5`

#### 0.0.5.5 
* Added ability to hide certain columns when exporting to excel.        
* Changed parameters names. See the documentation.

#### 0.0.5.2   
* Added `ExcelWriter` to `StyleFrame` class        
* Supports initializing `StyleFrame` with containers

#### 0.0.5      
* Initial release