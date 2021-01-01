# -*- coding: utf-8 -*-
"""
CoreRPAHive is a Robotic Process Automation library (RPA) for RobotFramework that allow the
developer create RPA script easier and reduce complexity under robot script layer.
"""
import re
from os.path import abspath, dirname, join
from setuptools import setup, find_packages


CURDIR = dirname(abspath(__file__))

with open("README.rst", "r", encoding='utf-8') as fh:
    LONG_DESCRIPTION = fh.read()

with open(join(CURDIR, 'ExcelDataDriver', '__init__.py'), encoding='utf-8') as f:
    VERSION = re.search("\n__version__ = '(.*)'", f.read()).group(1)

setup(
    name="robotframework-exceldatadriver",
    version=VERSION,
    author="QA Hive Co.,Ltd",
    author_email="support@qahive.com",
    description="ExcelDataDriver is a Excel Data-Driven Testing library for Robot Framework.",
    long_description=LONG_DESCRIPTION,
    license="Apache License 2.0",
    url='https://github.com/qahive/robotframework-ExcelDataDriver',
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "License :: OSI Approved :: Apache Software License",
        "Operating System :: OS Independent",
        "Topic :: Software Development :: Testing",
        "Topic :: Software Development :: Testing :: Acceptance",
        "Framework :: Robot Framework",
    ],
    keywords='robotframework testing automation data-driven qahive',
    platforms='any',
    install_requires=[
        'inject==3.3.2',
        'openpyxl==3.0.5',
        'setuptools==41.0.1',
        'Pillow>=6.2.1',
        'robotframework-seleniumlibrary',
        'robotframework-puppeteerlibrary',
    ],
    python_requires='>3.6',
    test_suite='nose.collector',
    tests_require=['nose', 'parameterized'],
    zip_safe=False,
)
