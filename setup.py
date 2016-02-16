#-------------------------------------------------------------------------------
# Name:        setup.py
# Purpose:     Standard module installation script
# Licence:     MIT License
#              This file is subject to the terms and conditions of the MIT License.
#              For further details, please refer to LICENSE.txt
#-------------------------------------------------------------------------------
"""
This script will install the PyZOS library into your Lib/site-packages directory
as a standard Python module. It can then be imported like any other module package.
"""

from setuptools import setup, find_packages

with open('README.rst') as fh:
    long_description = fh.read()

setup(
    name='PyZOS',
    version='0.0.1',
    description='Zemax / OpticStudio standalone ZOS-API using Python COM',
    long_description=long_description,
    author='Indranil Sinharoy',
    author_email='indranil_leo@yahoo.com',
    license='MIT',
    keywords='zemax opticstudio extensions COM optics',
    url=' ',
    packages=find_packages(),
    include_package_data=True,
    classifiers=[
        'Intended Audience :: Science/Research',
        'Natural Language :: English',
        'Environment :: Win32 (MS Windows)',
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows :: Windows XP',
        'Operating System :: Microsoft :: Windows :: Windows 7',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Topic :: Scientific/Engineering',
    ],
)
