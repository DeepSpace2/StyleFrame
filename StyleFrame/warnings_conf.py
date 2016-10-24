import warnings


# noinspection PyUnusedLocal
def formatwarning(message, category, filename, lineno, file=None, line=None):
    return '{}:{}: {}:{}\n'.format(filename, lineno, category.__name__, message)

warnings.formatwarning = formatwarning
warnings.simplefilter('always', DeprecationWarning)
