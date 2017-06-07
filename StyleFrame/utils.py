import re
from openpyxl.styles import colors as op_colors


def is_hex_color_string(hex_string):
    return re.match(r'^([A-Fa-f0-9]{8}|[A-Fa-f0-9]{6})$', hex_string) if hex_string else False


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


class fonts(BaseDefClass):
    arial = 'Arial'
    calibri = 'Calibri'
    aharoni = 'Aharoni CLM'
    david = 'David CLM'
    ellinia = 'Ellinia CLM'
    dejavu_sans = 'DejaVu Sans'
    aegean = 'Aegean'
    aegyptus = 'Aegyptus'
    anaktoria = 'Anaktoria'
    analecta = 'Analecta'
    anatolian = 'Anatolian'


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
