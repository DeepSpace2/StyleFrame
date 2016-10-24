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
    def __init__(self, bg_color=utils.colors.white, bold=False, font="Arial", font_size=12, font_color=utils.colors.black,
                 number_format=utils.number_formats.general, protection=False, underline=None,
                 border_type=utils.borders.thin):
        self.bold = bold
        self.font = font
        self.font_size = font_size
        self.font_color = font_color
        self.number_format = number_format
        self.protection = protection
        self.underline = underline
        self.border_type = border_type

        if bg_color.startswith('#'):
            bg_color = bg_color[1:]
        if utils.is_string_is_hex_color_code(hex_string=bg_color):
            self.bg_color = bg_color
        else:
            self.bg_color = utils.colors.get(bg_color, utils.colors.white)

        if font_color.startswith('#'):
            font_color = font_color[1:]
        if utils.is_string_is_hex_color_code(hex_string=font_color):
            self.font_color = font_color
        else:
            self.font_color = utils.colors.get(self.font_color, utils.colors.black)

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
