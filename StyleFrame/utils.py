import re


def is_string_is_hex_color_code(hex_string):
    return re.search(r'[a-fA-F0-9]{6}$', hex_string)
