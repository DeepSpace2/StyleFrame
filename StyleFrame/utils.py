import re
from openpyxl.styles import colors as op_colors


def is_string_is_hex_color_code(hex_string):
    return re.search(r'[a-fA-F0-9]{6}$', hex_string)


class BaseDefClass:
    @classmethod
    def get(cls, key, default=None):
        return cls.__dict__.get(key, default)


# The following classes names violate PEP8 for the sake of keeping backwards compatibility, at least for the meantime

class number_formats(BaseDefClass):
    general = 'General'
    general_integer = '0'
    general_float = '0.00'
    date = 'DD/MM/YY'
    percent = '0.0%'
    time_24_hours = 'HH:MM'
    time_12_hours = 'h:MM AM/PM'
    date_time = 'DD/MM/YY HH:MM'
    thousands_comma_sep = '#,##0'


class colors(BaseDefClass):
    white = 'FFFFFF'
    blue = op_colors.BLUE
    yellow = op_colors.YELLOW
    green = op_colors.GREEN
    black = op_colors.BLACK
    red = op_colors.RED
    purple = '800080'
    grey = 'D3D3D3'


class borders(BaseDefClass):
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


class underline(BaseDefClass):
    single = 'single'
    double = 'double'
