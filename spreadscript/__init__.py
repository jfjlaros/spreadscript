"""SpreadScript: Use a spreadsheet as a function.


Copyright (c) 2018 Jeroen F.J. Laros <jlaros@fixedpoint.nl>

Licensed under the MIT license, see the LICENSE file.
"""
from .spreadscript import SpreadScript


__version_info__ = ('0', '0', '2')

__version__ = '.'.join(__version_info__)
__author__ = 'Jeroen F.J. Laros'
__contact__ = 'jlaros@fixedpoint.nl'
__homepage__ = 'https://github.com/jfjlaros/spreadscript'

usage = __doc__.split('\n\n\n')


def doc_split(func):
    return func.__doc__.split('\n\n')[0]


def version(name):
    return '{} version {}\n\nAuthor   : {} <{}>\nHomepage : {}'.format(
        name, __version__, __author__, __contact__, __homepage__)
