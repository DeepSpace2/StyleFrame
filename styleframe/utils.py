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
    """
    :cvar str general: 'General'
    :cvar str general_integer: '0'
    :cvar str general_float: '0.00'
    :cvar str percent: '0.0%'
    :cvar str thousands_comma_sep: '#,##0'
    :cvar str date: 'DD/MM/YY'
    :cvar str time_24_hours: 'HH:MM'
    :cvar str time_24_hours_with_seconds: 'HH:MM:SS'
    :cvar str time_12_hours: 'h:MM AM/PM'
    :cvar str time_12_hours_with_seconds: 'h:MM:SS AM/PM'
    :cvar str date_time: 'DD/MM/YY HH:MM'
    :cvar str date_time_with_seconds: 'DD/MM/YY HH:MM:SS'
    """

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
        """
        .. versionadded:: 1.6

        :param int num_of_digits: Number of digits after the decimal point
        :return: A format string that represents a floating point number with the provided number of digits after the
            decimal point.

            For example, ``utils.number_formats.decimal_with_num_of_digits(2)`` will return ``'0.00'``
        :rtype: str
        """
        return '0.{}'.format('0' * num_of_digits)


class colors(BaseDefClass):
    """
    :cvar str white: '00FFFFFF'
    :cvar str blue: '000000FF'
    :cvar str dark_blue: '00000080'
    :cvar str yellow: '00FFFF00'
    :cvar str dark_yellow: '00808000'
    :cvar str green: '0000FF00'
    :cvar str dark_green: '00008000'
    :cvar str black: '00000000'
    :cvar str red: '00FF0000'
    :cvar str dark_red: '00800000'
    :cvar str purple: '800080'
    :cvar str grey: 'D3D3D3'
    """
    
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
    """
    .. versionadded:: 1.1

    :cvar str aegean: 'Aegean'
    :cvar str aegyptus: 'Aegyptus'
    :cvar str aharoni: 'Aharoni CLM'
    :cvar str anaktoria: 'Anaktoria'
    :cvar str analecta: 'Analecta'
    :cvar str anatolian: 'Anatolian'
    :cvar str arial: 'Arial'
    :cvar str calibri: 'Calibri'
    :cvar str david: 'David CLM'
    :cvar str dejavu_sans: 'DejaVu Sans'
    :cvar str ellinia: 'Ellinia CLM'
    """

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
    """
    :cvar str dash_dot: 'dashDot'
    :cvar str dash_dot_dot: 'dashDotDot'
    :cvar str dashed: 'dashed'
    :cvar str default_grid: 'default_grid'
    :cvar str dotted: 'dotted'
    :cvar str double: 'double'
    :cvar str hair: 'hair'
    :cvar str medium: 'medium'
    :cvar str medium_dash_dot: 'mediumDashDot'
    :cvar str medium_dash_dot_dot: 'mediumDashDotDot'
    :cvar str medium_dashed: 'mediumDashed'
    :cvar str slant_dash_dot: 'slantDashDot'
    :cvar str thick: 'thick'
    :cvar str thin: 'thin'
    """

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
    """
    :cvar str general: 'general'
    :cvar str left: 'left'
    :cvar str center: 'center'
    :cvar str right: 'right'
    :cvar str fill: 'fill'
    :cvar str justify: 'justify'
    :cvar str center_continuous: 'centerContinuous'
    :cvar str distributed: 'distributed'
    """
    
    general = 'general'
    left = 'left'
    center = 'center'
    right = 'right'
    fill = 'fill'
    justify = 'justify'
    center_continuous = 'centerContinuous'
    distributed = 'distributed'


class vertical_alignments(BaseDefClass):
    """
    :cvar str top: 'top'
    :cvar str center: 'center'
    :cvar str bottom: 'bottom'
    :cvar str justify: 'justify'
    :cvar str distributed: 'distributed'
    """
    
    top = 'top'
    center = 'center'
    bottom = 'bottom'
    justify = 'justify'
    distributed = 'distributed'


class underline(BaseDefClass):
    """
    :cvar str single: 'single'
    :cvar str double: 'double'
    """
    
    single = 'single'
    double = 'double'


class fill_pattern_types(BaseDefClass):
    """
    .. versionadded:: 1.2

    :cvar str solid: 'solid'
    :cvar str dark_down: 'darkDown'
    :cvar str dark_gray: 'darkGray'
    :cvar str dark_grid: 'darkGrid'
    :cvar str dark_horizontal: 'darkHorizontal'
    :cvar str dark_trellis: 'darkTrellis'
    :cvar str dark_up: 'darkUp'
    :cvar str dark_vertical: 'darkVertical'
    :cvar str gray0625: 'gray0625'
    :cvar str gray125: 'gray125'
    :cvar str light_down: 'lightDown'
    :cvar str light_gray: 'lightGray'
    :cvar str light_grid: 'lightGrid'
    :cvar str light_horizontal: 'lightHorizontal'
    :cvar str light_trellis: 'lightTrellis'
    :cvar str light_up: 'lightUp'
    :cvar str light_vertical: 'lightVertical'
    :cvar str medium_gray: 'mediumGray'
    """
    
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
    """
    :cvar str num: 'num'
    :cvar str percent: 'percent'
    :cvar str max: 'max'
    :cvar str min: 'min'
    :cvar str formula: 'formula'
    :cvar str percentile: 'percentile'
    """
    
    num = 'num'
    percent = 'percent'
    max = 'max'
    min = 'min'
    formula = 'formula'
    percentile = 'percentile'
