from setuptools import setup, find_packages
from StyleFrame.version import _version_

setup(
    name='StyleFrame',

    # Versions should comply with PEP440.  For a discussion on single-sourcing
    # the version across setup.py and the project code, see
    # https://packaging.python.org/en/latest/single_source_version.html
    version=_version_,

    description='A library that wraps pandas and openpyxl and allows easy styling of dataframes in excel. Documentation can be found at http://styleframe.readthedocs.org',
    # long_description=long_description,

    # The project's main homepage.
    url='https://github.com/DeepSpace2/StyleFrame',

    # Author details
    author='DeepSpace',
    author_email='deepspace2@gmail.com',

    entry_points={
        'console_scripts': ['styleframe = StyleFrame.commandline:execute_from_command_line']
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
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6'
    ],

    # What does your project relate to?
    keywords='pandas',

    # You can just specify the packages manually here if your project is
    # simple. Or you can use find_packages().
    packages=find_packages(exclude=['contrib', 'docs', 'tests*']),
    include_package_data=True,
    package_data={'': ['*.json']},
    # List run-time dependencies here.  These will be installed by pip when
    # your project is installed. For an analysis of "install_requires" vs pip's
    # requirements files see:
    # https://packaging.python.org/en/latest/requirements.html
    install_requires=['pandas>=0.16.2,<=0.20.3', 'openpyxl<=2.2.5', 'xlrd==1.0.0'],
)
