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
                     '0000FF00': '0000FF00'}

class Styler(object):
    """
    Creates openpyxl Style to be applied
    """
    def __init__(self, color='white', bold=False, size=12, number_format='General'):
        self.color = color
        self.bold = bold
        self.size = size
        self.number_format = number_format

    def create_style(self):
        side = Side(border_style='thin', color=colors.BLACK)
        border = Border(left=side, right=side, top=side, bottom=side)
        return Style(font=Font(name="Arial", size=self.size, color=colors.BLACK, bold=self.bold),
                     fill=PatternFill(patternType='solid', fgColor=Color(name_to_hex_dict[self.color])),
                     alignment=Alignment(horizontal='center', vertical='center', wrap_text=True, shrink_to_fit=True, indent=0),
                     border=border,
                     number_format=self.number_format)
