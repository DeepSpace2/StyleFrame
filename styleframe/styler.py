from openpyxl.cell import Cell

from . import utils
from colour import Color
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import PatternFill, NamedStyle, Color as OpenPyColor, Border, Side, Font, Alignment, Protection
from openpyxl.comments import Comment
from pprint import pformat

from typing import Dict, List, Optional, Union


class Styler:
    """
    Used to represent a style

    :param bg_color: The background color
    :type bg_color: str: one of :class:`.utils.colors`, hex string or color name ie `'yellow'` Excel supports
    :param bool bold: If ``True``, a bold typeface is used
    :param font: The font to use
    :type font: str: one of :class:`.utils.fonts` or other font name Excel supports
    :param int font_size: The font size
    :param font_color: The font color
    :type font_color: str: one of :class:`.utils.colors`, hex string or color name ie `'yellow'` Excel supports
    :param number_format: The format of the cell's value
    :type number_format: str: one of :class:`.utils.number_formats` or any other format Excel supports
    :param bool protection: If ``True``, the cell/column will be write-protected
    :param underline: The underline type
    :type underline: str: one of :class:`.utils.underline` or any other underline Excel supports
    :param border_type: The border type
    :type border_type: str: one of :class:`.utils.borders` or any other border type Excel supports

    .. versionadded:: 1.2

    :param horizontal_alignment: Text's horizontal alignment
    :type horizontal_alignment: str: one of :class:`.utils.horizontal_alignments` or any other horizontal alignment Excel supports
    :param vertical_alignment: Text's vertical alignment
    :type vertical_alignment: str: one of :class:`.utils.vertical_alignments` or any other vertical alignment Excel supports

    .. versionadded:: 1.3

    :param bool wrap_text:
    :param bool shrink_to_fit:
    :param fill_pattern_type: Cells's fill pattern type
    :type fill_pattern_type: str: one of :class:`.utils.fill_pattern_types` or any other fill pattern type Excel supports
    :param int indent:
    :param str comment_author:
    :param str comment_text:
    :param int text_rotation: Integer in the range 0 - 180

    .. versionadded:: 4.0

    :param date_format:
    :type date_format: str: one of :class:`.utils.number_formats` or any other format Excel supports
    :param time_format:
    :type time_format: str: one of :class:`.utils.number_formats` or any other format Excel supports
    :param date_time_format:
    :type date_time_format: str: one of :class:`.utils.number_formats` or any other format Excel supports

    .. note:: For any of ``date_format``, ``time_format`` and ``date_time_format`` to take effect, the value being
              styled must be an actual ``date``/``time``/``datetime`` object.

    .. versionadded:: 4.1

    :param bool strikethrough:
    :param bool italic:
    """

    cache: Dict['Styler', NamedStyle] = {}

    def __init__(self,
                 bg_color: Optional[str] = None,
                 bold: bool = False,
                 font: str = utils.fonts.arial,
                 font_size: Union[int, float] = 12.0,
                 font_color: Optional[str] = None,
                 number_format: str = utils.number_formats.general,
                 protection: bool = False,
                 underline: Optional[str] = None,
                 border_type: str = utils.borders.thin,
                 horizontal_alignment: str = utils.horizontal_alignments.center,
                 vertical_alignment: str = utils.vertical_alignments.center,
                 wrap_text: bool = True,
                 shrink_to_fit: bool = True,
                 fill_pattern_type: str = utils.fill_pattern_types.solid,
                 indent: Union[int, float] = 0.0,
                 comment_author: Optional[str] = None,
                 comment_text: Optional[str] = None,
                 text_rotation: int = 0,
                 date_format: str = utils.number_formats.date,
                 time_format: str = utils.number_formats.time_24_hours,
                 date_time_format: str = utils.number_formats.date_time,
                 strikethrough: bool = False,
                 italic: bool = False):

        def get_color_from_string(color_str: str, default_color: Optional[str] = None) -> str:
            if color_str and color_str.startswith('#'):
                color_str = color_str[1:]
            if not utils.is_hex_color_string(hex_string=color_str):
                color_str = utils.colors.get(color_str, default_color)
            return color_str

        if border_type == utils.borders.default_grid:
            if bg_color is not None or fill_pattern_type != utils.fill_pattern_types.solid:
                raise ValueError('`bg_color`or `fill_pattern_type` conflict with border_type={}'.format(utils.borders.default_grid))
            self.border_type = None
            self.fill_pattern_type = None
        else:
            self.border_type = border_type
            self.fill_pattern_type = fill_pattern_type

        self.bold = bold
        self.font = font
        self.font_size = font_size
        self.number_format = number_format
        self.protection = protection
        self.underline = underline
        self.horizontal_alignment = horizontal_alignment
        self.vertical_alignment = vertical_alignment
        self.bg_color = get_color_from_string(bg_color, default_color=utils.colors.white)
        self.font_color = get_color_from_string(font_color, default_color=utils.colors.black)
        self.shrink_to_fit = shrink_to_fit
        self.wrap_text = wrap_text
        self.indent = indent
        self.comment_author = comment_author
        self.comment_text = comment_text
        self.text_rotation = text_rotation
        self.date_format = date_format
        self.time_format = time_format
        self.date_time_format = date_time_format
        self.strikethrough = strikethrough
        self.italic = italic

    def __eq__(self, other):
        return isinstance(other, self.__class__) and self.__dict__ == other.__dict__

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
                          bold=self.bold, underline=self.underline, strikethrough=self.strikethrough,
                          italic=self.italic),
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
    def from_openpyxl_style(cls, openpyxl_style: Cell, theme_colors: List[str],
                            openpyxl_comment: Optional[Comment] = None):
        def _calc_new_hex_from_theme_hex_and_tint(theme_hex, color_tint):
            if not theme_hex.startswith('#'):
                theme_hex = '#' + theme_hex
            color_obj = Color(theme_hex)
            color_obj.luminance = _calc_lum_from_tint(color_tint, color_obj.luminance)
            return color_obj.hex_l[1:]

        def _calc_lum_from_tint(color_tint: Optional[float], current_lum: float) -> float:
            """"
            Based on https://ciintelligence.blogspot.co.il/2012/02/converting-excel-theme-color-and-tint.html
            """
            if color_tint is None:
                return current_lum

            current_lum *= 255

            if color_tint < 0:
                return current_lum * (1.0 + color_tint) / 255

            return (current_lum * (1.0 - color_tint) + (255 - 255 * (1.0 - color_tint))) / 255

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
        strikethrough = openpyxl_style.font.strikethrough
        italic = openpyxl_style.font.italic
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
                   fill_pattern_type, indent, comment_author, comment_text, text_rotation,
                   strikethrough=strikethrough, italic=italic)

    @classmethod
    def combine(cls, *styles: 'Styler'):
        """
        .. versionadded:: 1.6

        Used to combine :class:`Styler` objects. The right-most object has precedence.
        For example:

        ::

            Styler.combine(Styler(bg_color='yellow', font_size=24), Styler(bg_color='blue'))

        will return

        ::

            Styler(bg_color='blue', font_size=24)

        :param styles: Iterable of Styler objects
        :type styles: list or tuple or set
        :return: self
        :rtype: :class:`Styler`
        """

        return sum(styles, cls())

    create_style = to_openpyxl_style


class ColorScaleConditionalFormatRule:
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
