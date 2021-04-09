import codecs
import os
import re

from setuptools import setup, find_packages

DESCRIPTION = 'A library that wraps pandas and openpyxl and allows easy styling of dataframes in excel. Documentation can be found at http://styleframe.readthedocs.org'

here = os.path.abspath(os.path.dirname(__file__)).lower()


def read(*parts):
    # intentionally *not* adding an encoding option to open
    return codecs.open(os.path.join(here, *parts), 'r').read()


def find_version(*file_paths):
    version_file = read(*file_paths)
    version_match = re.search(r"^_version_ = ['\"]([^'\"]*)['\"]",
                              version_file, re.M)
    if version_match:
        return version_match.group(1)
    raise RuntimeError("Unable to find version string.")


setup(
    name='styleframe',

    # Versions should comply with PEP440.  For a discussion on single-sourcing
    # the version across setup.py and the project code, see
    # https://packaging.python.org/en/latest/single_source_version.html
    version=find_version('styleframe', 'version.py'),

    description=DESCRIPTION,
    long_description=DESCRIPTION,

    # The project's main homepage.
    url='https://github.com/DeepSpace2/StyleFrame',

    # Author details
    author='DeepSpace',
    author_email='deepspace2@gmail.com',

    entry_points={
        'console_scripts': ['styleframe = styleframe.command_line.commandline:execute_from_command_line']
    },

    # Choose your license
    # license='MIT',
    # See https://pypi.python.org/pypi?%3Aaction=list_classifiers
    classifiers=[
        # How mature is this project? Common values are
        #   3 - Alpha
        #   4 - Beta
        #   5 - Production/Stable
        'Development Status :: 5 - Production/Stable',

        # Indicate who your project is intended for
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Build Tools',

        # Pick your license as you wish (should match "license" above)
        'License :: OSI Approved :: MIT License',

        # Specify the Python versions you support here. In particular, ensure
        # that you indicate whether you support Python 2, Python 3 or both.
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9'
    ],

    # What does your project relate to?
    keywords=('pandas', 'openpyxl', 'excel'),

    # You can just specify the packages manually here if your project is
    # simple. Or you can use find_packages().
    packages=find_packages(exclude=['contrib', 'docs', 'tests*']),
    include_package_data=True,
    package_data={'': ['*.json']},
    # List run-time dependencies here.  These will be installed by pip when
    # your project is installed. For an analysis of "install_requires" vs pip's
    # requirements files see:
    # https://packaging.python.org/en/latest/requirements.html
    install_requires=['openpyxl>=2.5,<4', 'colour>=0.1.5,<0.2', 'jsonschema', 'xlrd>=1.0.0,<1.3.0', 'pandas<2']
)
