commandline_json_schema = {
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
                    "additionalProperties": False
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
            "additionalProperties": False
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
            "additionalProperties": False
        }
    },
    "type": "array",
    "items": {
        "$ref": "#/definitions/Sheet"
    },
    "minItems": 1
}
