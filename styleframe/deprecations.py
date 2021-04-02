import warnings

from functools import wraps

funcs_to_deprecated_kwargs = {
    'data_df': {
        'self': ''
    }
}

properties_to_deprecated_warnings = {
    'data_df': 'data_df is deprecated to emphasize that all data-mangling operations\n'
               'should be performed on a DataFrame object before creating a StyleFrame.\n'
               'If you know what you are doing change your code to use _data_df directly.'
}

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


def deprecated_prop(prop):
    @wraps(prop)
    def inner(*args, **kwargs):
        warnings.warn('{} is deprecated. {}'.format(prop.__name__, properties_to_deprecated_warnings[prop.__name__]),
                      DeprecationWarning, stacklevel=2)
        return prop(*args, **kwargs)
    return inner


warnings.formatwarning = formatwarning
warnings.simplefilter('always', DeprecationWarning)
