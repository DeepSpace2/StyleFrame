# coding:utf-8
from openpyxl.styles import PatternFill, Style, Color, Border, Side, Font, Alignment, Protection

try:  # python2
    import utils
except ImportError:  # python3
    from StyleFrame import utils


class Styler(object):
    """
    Creates openpyxl Style to be applied
    """
    def __init__(self, bg_color=None, bold=False, font="Arial", font_size=12, font_color=None,
                 number_format=utils.number_formats.general, protection=False, underline=None,
                 border_type=utils.borders.thin):

        def get_color_from_string(color_str, default_color=None):
            if color_str and color_str.startswith('#'):
                color_str = color_str[1:]
            if not utils.is_hex_color_string(hex_string=color_str):
                color_str = utils.colors.get(color_str, default_color)
            return color_str

        self.bold = bold
        self.font = font
        self.font_size = font_size
        self.number_format = number_format
        self.protection = protection
        self.underline = underline
        self.border_type = border_type
        self.bg_color = get_color_from_string(bg_color, default_color=utils.colors.white)
        self.font_color = get_color_from_string(font_color, default_color=utils.colors.black)

    def create_style(self):
        side = Side(border_style=self.border_type, color=utils.colors.black)
        border = Border(left=side, right=side, top=side, bottom=side)
        return Style(font=Font(name=self.font, size=self.font_size, color=Color(self.font_color),
                               bold=self.bold, underline=self.underline),
                     fill=PatternFill(patternType='solid', fgColor=self.bg_color),
                     alignment=Alignment(horizontal='center', vertical='center', wrap_text=True, shrink_to_fit=True, indent=0),
                     border=border,
                     number_format=self.number_format,
                     protection=Protection(locked=self.protection))
