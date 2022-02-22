import warnings

from functools import wraps

funcs_to_deprecated_kwargs = {'read_excel': {'sheetname': 'sheet_name'}}


# noinspection PyUnusedLocal
def formatwarning(message, category, filename, lineno, file=None, line=None):
    return '{}:{}: {}:{}\n'.format(filename, lineno, category.__name__, message)


def deprecated_kwargs(deprecated_kwargs):
    def wrapper(func):
        @wraps(func)
        def inner(*args, **kwargs):
            for deprecated_kwarg in deprecated_kwargs:
                if deprecated_kwarg in kwargs:
                    new_kwarg = funcs_to_deprecated_kwargs[func.__name__][deprecated_kwarg]
                    warnings.warn('{} kwarg is deprecated, use {} instead'.format(deprecated_kwarg,
                                                                                             new_kwarg),
                                  DeprecationWarning, stacklevel=2)

            return func(*args, **kwargs)
        return inner
    return wrapper


warnings.formatwarning = formatwarning
warnings.simplefilter('always', DeprecationWarning)
