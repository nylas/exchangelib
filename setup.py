#!/usr/bin/env python
"""
Release notes:
* Bump version in exchangelib/__init__.py
* Bump version in CHANGELOG.md
* Commit and push changes
* Build package: rm dist/* && python setup.py sdist bdist_wheel
* Push to PyPI: twine upload dist/*
* Create release on GitHub
"""
import io
import os

from setuptools import setup


__version__ = None
with io.open(os.path.join(os.path.dirname(__file__), 'exchangelib/__init__.py'), encoding='utf-8') as f:
    for l in f:
        if not l.startswith('__version__'):
            continue
        __version__ = l.split('=')[1].strip(' "\'\n')
        break


def read(file_name):
    with io.open(os.path.join(os.path.dirname(__file__), file_name), encoding='utf-8') as f:
        return f.read()


setup(
    name='exchangelib',
    version=__version__,
    author='Erik Cederstrand',
    author_email='erik@cederstrand.dk',
    description='Client for Microsoft Exchange Web Services (EWS)',
    long_description=read('README.md'),
    long_description_content_type='text/markdown',
    license='BSD',
    keywords='Exchange EWS autodiscover',
    install_requires=['requests>=2.7', 'requests_ntlm>=0.2.0', 'dnspython>=1.14.0', 'pytz', 'lxml>3.0',
                      'cached_property', 'future', 'six', 'tzlocal', 'python-dateutil', 'pygments', 'defusedxml',
                      'isodate', 'requests_kerberos', "typing; python_version < '3.5'"],
    packages=['exchangelib'],
    tests_require=['PyYAML', 'requests_mock', 'psutil'],
    python_requires=">=2.7, !=3.0.*, !=3.1.*, !=3.2.*, !=3.3.*",
    test_suite='tests',
    zip_safe=False,
    url='https://github.com/ecederstrand/exchangelib',
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Topic :: Communications',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 3',
    ],
)
