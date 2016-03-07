#!/usr/bin/env python3
import os

from setuptools import setup


def read(fname):
    return open(os.path.join(os.path.dirname(__file__), fname)).read()

setup(
    name='exchangelib',
    version='1.2',
    author='Erik Cederstrand',
    author_email='erik@cederstrand.dk',
    description='Client for Microsoft Exchange Web Services (EWS)',
    long_description=read('README'),
    license='BSD',
    keywords='Exchange EWS autodiscover',
    install_requires=['requests>=2.7', 'requests-ntlm>=0.2.0', 'dnspython3>=1.12.0', 'pytz', 'lxml'],
    packages=['exchangelib'],
    test_requires=['PyYAML'],
    test_suite='tests',
    zip_safe=False,
    url='https://github.com/ecederstrand/exchangelib.git',
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Topic :: Communications',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python :: 3 :: Only',
    ],
)
