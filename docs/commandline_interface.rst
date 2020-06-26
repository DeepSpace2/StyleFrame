Commandline Interface
=====================

General Information
-------------------

Starting with version 1.1 styleframe offers a commandline interface
that lets you create an xlsx file from a json file.

Usage
-----

====================================   =========================================================================
Flag                                   Explanation
====================================   =========================================================================
``-v``                                 Displays the installed versions of styleframe and its dependencies
``--json_path`` or ``--json-path``     Path to the json file
``--json``                             The json string which defines the Excel file, see example below
``--output_path`` or ``output-path``   Path to the output xlsx file. If not provided defaults to ``output.xlsx``
``--test``                             Execute the tests
====================================   =========================================================================


Usage Examples
^^^^^^^^^^^^^^

``$ styleframe --json_path data.json --output_path data.xlsx``

``$ styleframe --json "[{\"sheet_name\": \"sheet_1\", \"columns\": [{\"col_name\": \"col_a\", \"cells\": [{\"value\": 1}]}]}]"``

.. note:: You may need to use different syntax to pass a JSON string depending on your OS and terminal application.

JSON Format
-----------

The input JSON should be thought of as an hierarchy of predefined entities,
some of which correspond to a Python class used by StyleFrame.
The top-most level should be a list of ``sheet`` entities (see below).

The provided JSON is validated against the following schema:

::

   {
       "$schema": "http://json-schema.org/draft-04/schema#",
       "title": "sheets",
       "definitions": {
           "Sheet": {
               "$id": "#sheet",
               "title": "sheet",
               "type": "object",
               "properties": {
                   "sheet_name": {
                       "type": "string"
                   },
                   "columns": {
                       "type": "array",
                       "items": {
                           "$ref": "#/definitions/Column"
                       },
                       "minItems": 1
                   },
                   "row_heights": {
                       "type": "object"
                   },
                   "extra_features": {
                       "type": "object"
                   },
                   "default_styles": {
                       "type": "object",
                       "properties": {
                           "headers": {
                               "$ref": "#/definitions/Style"
                           },
                           "cells": {
                               "$ref": "#/definitions/Style"
                           }
                       },
                       "additionalProperties": false
                   }
               },
               "required": [
                   "sheet_name",
                   "columns"
               ]
           },
           "Column": {
               "$id": "#column",
               "title": "column",
               "type": "object",
               "properties": {
                   "col_name": {
                       "type": "string"
                   },
                   "style": {
                       "$ref": "#/definitions/Style"
                   },
                   "width": {
                       "type": "number"
                   },
                   "cells": {
                       "type": "array",
                       "items": {
                           "$ref": "#/definitions/Cell"
                       }
                   }
               },
               "required": [
                   "col_name",
                   "cells"
               ]
           },
           "Cell": {
               "$id": "#cell",
               "title": "cell",
               "type": "object",
               "properties": {
                   "value": {},
                   "style": {
                       "$ref": "#/definitions/Style"
                   }
               },
               "required": [
                   "value"
               ],
               "additionalProperties": false
           },
           "Style": {
               "$id": "#style",
               "title": "style",
               "type": "object",
               "properties": {
                   "bg_color": {
                       "type": "string"
                   },
                   "bold": {
                       "type": "boolean"
                   },
                   "font": {
                       "type": "string"
                   },
                   "font_size": {
                       "type": "number"
                   },
                   "font_color": {
                       "type": "string"
                   },
                   "number_format": {
                       "type": "string"
                   },
                   "protection": {
                       "type": "boolean"
                   },
                   "underline": {
                       "type": "string"
                   },
                   "border_type": {
                       "type": "string"
                   },
                   "horizontal_alignment": {
                       "type": "string"
                   },
                   "vertical_alignment": {
                       "type": "string"
                   },
                   "wrap_text": {
                       "type": "boolean"
                   },
                   "shrink_to_fit": {
                       "type": "boolean"
                   },
                   "fill_pattern_type": {
                       "type": "string"
                   },
                   "indent": {
                       "type": "number"
                   }
               },
               "additionalProperties": false
           }
       },
       "type": "array",
       "items": {
           "$ref": "#/definitions/Sheet"
       },
       "minItems": 1
   }

An example JSON:

::

   [
     {
       "sheet_name": "Sheet1",
       "default_styles": {
         "headers": {
           "font_size": 17,
           "bg_color": "yellow"
         },
         "cells": {
           "bg_color": "red"
         }
       },
       "columns": [
         {
           "col_name": "col_a",
           "style": {"bg_color": "blue", "font_color": "yellow"},
           "width": 30,
           "cells": [
             {
               "value": 1
             },
             {
               "value": 2,
               "style": {
                 "bold": true,
                 "font": "Arial",
                 "font_size": 30,
                 "font_color": "green",
                 "border_type": "double"
               }
             }
           ]
         },
         {
           "col_name": "col_b",
           "cells": [
             {
               "value": 3
             },
             {
               "value": 4,
               "style": {
                 "bold": true,
                 "font": "Arial",
                 "font_size": 16
               }
             }
           ]
         }
       ],
       "row_heights": {
         "3": 40
       },
       "extra_features": {
         "row_to_add_filters": 0,
         "columns_and_rows_to_freeze": "A7",
         "startrow": 5
       }
     }
   ]

style
^^^^^

Corresponds to :ref:`Styler <styler-class>` class.

This entity uses the arguments of ``Styler.__init__()`` as keys.
Any missing keys in the JSON will be given the same default values.

``"style": {"bg_color": "yellow", "bold": true}``

cell
^^^^

This entity represents a single cell in the sheet.

Required keys:

``"value"`` - The cell's value.

Optional keys:

``"style"`` - The ``style`` entity for this cell. 
If not provided, the ``style`` provided to the ``coloumn`` entity will be used.
If that was not provided as well, the default ``Styler.__init__()`` values will be used.  

``{"value": 42, "style": {"border": "double"}}``

column
^^^^^^

This entity represents a column in the sheet.

Required keys:

``"col_name"`` - The column name.

``"cells"`` - A list of ``cell`` entities.

Optional keys:

``"style"`` - A style used for the entire column. If not provided the default ``Styler.__init__()`` values will be used. 

``"width"`` - The column's width. If not provided Excel's default column width will be used.

sheet
^^^^^

This entity represents the entire sheet.

Required keys:

``"sheet_name"`` - The sheet's name.

``"columns"`` - A list of ``column`` entities.

Optional keys:

``"default_styles"`` - A JSON object with items as keys and ``style`` entities as values.
Currently supported items: ``headers`` and ``cells``.

``"default_styles": {"headers": {"bg_color": "blue"}}``
 
``"row_heights"`` - A JSON object with rows indexes as keys and heights as value.

``"extra_features"`` - A JSON that contains the same arguments as the
``to_excel`` method, such as ``"row_to_add_filters"``, ``"columns_and_rows_to_freeze"``,
``"columns_to_hide"``, ``"right_to_left"`` and ``"allow_protection"``. 
You can also use other arguments that Pandas' ``"to_excel"`` accepts.
