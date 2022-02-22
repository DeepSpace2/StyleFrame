def get_python_version():
    import sys
    return 'Python {}'.format(sys.version)


def get_pandas_version():
    from pandas import __version__ as pd_version
    return 'pandas {}'.format(pd_version)


def get_openpyxl_version():
    from openpyxl import __version__ as openpyxl_version
    return 'openpyxl {}'.format(openpyxl_version)


def get_all_versions():
    return _versions_


_version_ = '4.0.0'
_python_version_ = get_python_version()
_pandas_version_ = get_pandas_version()
_openpyxl_version_ = get_openpyxl_version()
_versions_ = '{}\n{}\n{}\nStyleFrame {}'.format(_python_version_, _pandas_version_, _openpyxl_version_, _version_)
