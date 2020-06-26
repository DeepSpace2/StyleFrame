import re

HEX_REGEX = re.compile(r'^([A-Fa-f0-9]{8}|[A-Fa-f0-9]{6})$')


def is_hex_color_string(hex_string):
    return HEX_REGEX.match(hex_string) if hex_string else False


class BaseDefClass:
    @classmethod
    def get(cls, key, default=None):
        return cls.__dict__.get(key, default)


# The following classes names violate PEP8 for the sake of keeping backwards compatibility, at least for the meantime

class number_formats(BaseDefClass):
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

    default_date_format = date
    default_time_format = time_24_hours
    default_date_time_format = date_time

    @staticmethod
    def decimal_with_num_of_digits(num_of_digits):
        return '0.{}'.format('0' * num_of_digits)


class colors(BaseDefClass):
    white = '00FFFFFF'
    blue = '000000FF'
    dark_blue = '00000080'
    yellow = '00FFFF00'
    dark_yellow = '00808000'
    green = '0000FF00'
    dark_green = '00008000'
    black = '00000000'
    red = '00FF0000'
    dark_red = '00800000'
    purple = '800080'
    grey = 'D3D3D3'


class fonts(BaseDefClass):
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


class borders(BaseDefClass):
    dash_dot = 'dashDot'
    dash_dot_dot = 'dashDotDot'
    dashed = 'dashed'
    default_grid = 'default_grid'
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


class horizontal_alignments(BaseDefClass):
    general = 'general'
    left = 'left'
    center = 'center'
    right = 'right'
    fill = 'fill'
    justify = 'justify'
    center_continuous = 'centerContinuous'
    distributed = 'distributed'


class vertical_alignments(BaseDefClass):
    top = 'top'
    center = 'center'
    bottom = 'bottom'
    justify = 'justify'
    distributed = 'distributed'


class underline(BaseDefClass):
    single = 'single'
    double = 'double'


class fill_pattern_types(BaseDefClass):
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


class conditional_formatting_types(BaseDefClass):
    num = 'num'
    percent = 'percent'
    max = 'max'
    min = 'min'
    formula = 'formula'
    percentile = 'percentile'
