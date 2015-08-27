# coding:utf-8
from openpyxl.styles import PatternFill, Style, Color, Border, Side, Font, Alignment, colors

name_to_hex_dict = {'white': 'FFFFFF',
                     'blue':  colors.BLUE,
                     'yellow': colors.YELLOW,
                     'green': colors.GREEN,
                     'black': colors.BLACK,
                     'red': colors.RED,
                     'purple': '800080',
                     '00FFFFFF': '00FFFFFF',
                     '000000FF': '000000FF',
                     '0000FF00': '0000FF00',
                     '00FFFF00': '00FFFF00'}


class Styler(object):
    """
    Creates openpyxl Style to be applied
    """
    def __init__(self, bg_color='white', bold=False, font_size=12, number_format='General'):
        self.bg_color = bg_color
        self.bold = bold
        self.font_size = font_size
        self.number_format = number_format

    def create_style(self):
        side = Side(border_style='thin', color=colors.BLACK)
        border = Border(left=side, right=side, top=side, bottom=side)
        return Style(font=Font(name="Arial", size=self.font_size, color=colors.BLACK, bold=self.bold),
                     fill=PatternFill(patternType='solid', fgColor=Color(name_to_hex_dict[self.bg_color])),
                     alignment=Alignment(horizontal='center', vertical='center', wrap_text=True, shrink_to_fit=True, indent=0),
                     border=border,
                     number_format=self.number_format)
