#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from setuptools import find_packages, setup
import os


URL = 'https://github.com/seoktaehyeon/hrobot'
NAME = 'hrobot'
VERSION = '0.1'
DESCRIPTION = 'Hybrid Robot'
if os.path.exists('README.md'):
    with open('README.md', encoding='utf-8') as f:
        LONG_DESCRIPTION = f.read()
else:
    LONG_DESCRIPTION = DESCRIPTION
AUTHOR = 'Jing&Will'
AUTHOR_EMAIL = 'v.stone@163.com'
LICENSE = 'Apache'
PLATFORMS = [
    'any',
]
REQUIRES = [
    'xlrd==1.2.0',
    'xlwt==1.3.0',
    'robotframework>=4.0.0',
    'allure-robotframework>=2.9.0',
    'selenium>=3.14.0',
    'paramiko>=2.7.0',
    'PyYAML>=5.4.1',
]
CONSOLE_SCRIPTS = 'hrobot = hrobot.hcmd:main'

setup(
    name=NAME,
    version=VERSION,
    description=(
        DESCRIPTION
    ),
    long_description=LONG_DESCRIPTION,
    long_description_content_type='text/markdown',
    author=AUTHOR,
    author_email=AUTHOR_EMAIL,
    maintainer=AUTHOR,
    maintainer_email=AUTHOR_EMAIL,
    license=LICENSE,
    packages=find_packages(),
    platforms=PLATFORMS,
    url=URL,
    install_requires=REQUIRES,
    entry_points={
        'console_scripts': [CONSOLE_SCRIPTS],
    }
)
