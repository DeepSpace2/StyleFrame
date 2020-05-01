from . import utils
from colour import Color
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import PatternFill, NamedStyle, Color as OpenPyColor, Border, Side, Font, Alignment, Protection
from openpyxl.comments import Comment
from pprint import pformat


class Styler(object):
    """
    Creates openpyxl Style to be applied
    """

    cache = {}

    def __init__(self, bg_color=None, bold=False, font=utils.fonts.arial, font_size=12.0, font_color=None,
                 number_format=utils.number_formats.general, protection=False, underline=None,
                 border_type=utils.borders.thin, horizontal_alignment=utils.horizontal_alignments.center,
                 vertical_alignment=utils.vertical_alignments.center, wrap_text=True, shrink_to_fit=True,
                 fill_pattern_type=utils.fill_pattern_types.solid, indent=0.0, comment_author=None, comment_text=None,
                 text_rotation=0):

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
        self.horizontal_alignment = horizontal_alignment
        self.vertical_alignment = vertical_alignment
        self.bg_color = get_color_from_string(bg_color, default_color=utils.colors.white)
        self.font_color = get_color_from_string(font_color, default_color=utils.colors.black)
        self.shrink_to_fit = shrink_to_fit
        self.wrap_text = wrap_text
        self.fill_pattern_type = fill_pattern_type
        self.indent = indent
        self.comment_author = comment_author
        self.comment_text = comment_text
        self.text_rotation = text_rotation

        if self.border_type == utils.borders.default_grid:
            if self.bg_color is not None:
                raise ValueError('bg_color and border_type={} can not be used together'.format(utils.borders.default_grid))
            self.border_type = None
            self.fill_pattern_type = None

    def __eq__(self, other):
        if not isinstance(other, self.__class__):
            return False
        return self.__dict__ == other.__dict__

    def __hash__(self):
        return hash(tuple((k, v) for k, v in sorted(self.__dict__.items())))

    def __add__(self, other):
        default = Styler().__dict__
        d = dict(self.__dict__)
        for k, v in other.__dict__.items():
            if v != default[k]:
                d[k] = v
        return Styler(**d)

    def __repr__(self):
        return pformat(self.__dict__)

    def generate_comment(self):
        if any((self.comment_author, self.comment_text)):
            return Comment(self.comment_text, self.comment_author)
        return None

    @classmethod
    def default_header_style(cls):
        return cls(bold=True)

    def to_openpyxl_style(self):
        try:
            openpyxl_style = self.cache[self]
        except KeyError:
            side = Side(border_style=self.border_type, color=utils.colors.black)
            border = Border(left=side, right=side, top=side, bottom=side)
            openpyxl_style = self.cache[self] = NamedStyle(
                name=str(hash(self)),
                font=Font(name=self.font, size=self.font_size, color=OpenPyColor(self.font_color),
                          bold=self.bold, underline=self.underline),
                fill=PatternFill(patternType=self.fill_pattern_type, fgColor=self.bg_color),
                alignment=Alignment(horizontal=self.horizontal_alignment, vertical=self.vertical_alignment,
                                    wrap_text=self.wrap_text, shrink_to_fit=self.shrink_to_fit,
                                    indent=self.indent, text_rotation=self.text_rotation),
                border=border,
                number_format=self.number_format,
                protection=Protection(locked=self.protection)
            )
        return openpyxl_style

    @classmethod
    def from_openpyxl_style(cls, openpyxl_style, theme_colors, openpyxl_comment=None):
        def _calc_new_hex_from_theme_hex_and_tint(theme_hex, color_tint):
            if not theme_hex.startswith('#'):
                theme_hex = '#' + theme_hex
            color_obj = Color(theme_hex)
            color_obj.luminance = _calc_lum_from_tint(color_tint, color_obj.luminance)
            return color_obj.hex_l[1:]

        def _calc_lum_from_tint(color_tint, current_lum):
            # based on http://ciintelligence.blogspot.co.il/2012/02/converting-excel-theme-color-and-tint.html
            if not color_tint:
                return current_lum
            return current_lum * (1.0 + color_tint)

        bg_color = openpyxl_style.fill.fgColor.rgb

        # in case we are dealing with a "theme color"
        if not isinstance(bg_color, str):
            try:
                bg_color = theme_colors[openpyxl_style.fill.fgColor.theme]
            except (AttributeError, IndexError, TypeError):
                bg_color = utils.colors.white[:6]
            tint = openpyxl_style.fill.fgColor.tint
            bg_color = _calc_new_hex_from_theme_hex_and_tint(bg_color, tint)

        bold = openpyxl_style.font.bold
        font = openpyxl_style.font.name
        font_size = openpyxl_style.font.size
        try:
            font_color = openpyxl_style.font.color.rgb
        except AttributeError:
            font_color = utils.colors.black

        # in case we are dealing with a "theme color"
        if not isinstance(font_color, str):
            try:
                font_color = theme_colors[openpyxl_style.font.color.theme]
            except (AttributeError, IndexError, TypeError):
                font_color = utils.colors.black[:6]
            tint = openpyxl_style.font.color.tint
            font_color = _calc_new_hex_from_theme_hex_and_tint(font_color, tint)

        number_format = openpyxl_style.number_format
        protection = openpyxl_style.protection.locked
        underline = openpyxl_style.font.underline
        border_type = openpyxl_style.border.bottom.border_style
        horizontal_alignment = openpyxl_style.alignment.horizontal
        vertical_alignment = openpyxl_style.alignment.vertical
        wrap_text = openpyxl_style.alignment.wrap_text or False
        shrink_to_fit = openpyxl_style.alignment.shrink_to_fit
        fill_pattern_type = openpyxl_style.fill.patternType
        indent = openpyxl_style.alignment.indent
        text_rotation = openpyxl_style.alignment.text_rotation

        if openpyxl_comment:
            comment_author = openpyxl_comment.author
            comment_text = openpyxl_comment.text
        else:
            comment_author = None
            comment_text = None

        return cls(bg_color, bold, font, font_size, font_color,
                   number_format, protection, underline,
                   border_type, horizontal_alignment,
                   vertical_alignment, wrap_text, shrink_to_fit,
                   fill_pattern_type, indent, comment_author, comment_text, text_rotation)

    @classmethod
    def combine(cls, *styles):
        return sum(styles, cls())

    create_style = to_openpyxl_style


class ColorScaleConditionalFormatRule(object):
    """Creates a color scale conditional format rule. Wraps openpyxl's ColorScaleRule.
    Mostly should not be used directly, but through StyleFrame.add_color_scale_conditional_formatting
    """

    def __init__(self, start_type, start_value, start_color, end_type, end_value, end_color,
                 mid_type=None, mid_value=None, mid_color=None, columns_range=None):

        self.columns = columns_range

        # checking against None explicitly since mid_value may be 0
        if all(val is not None for val in (mid_type, mid_value, mid_color)):
            self.rule = ColorScaleRule(start_type=start_type, start_value=start_value,
                                       start_color=OpenPyColor(start_color),
                                       mid_type=mid_type, mid_value=mid_value,
                                       mid_color=OpenPyColor(mid_color),
                                       end_type=end_type, end_value=end_value,
                                       end_color=OpenPyColor(end_color))
        else:
            self.rule = ColorScaleRule(start_type=start_type, start_value=start_value,
                                       start_color=OpenPyColor(start_color),
                                       end_type=end_type, end_value=end_value,
                                       end_color=OpenPyColor(end_color))
