"""
Script for building the example.

Usage:
    python setup.py py2app
"""

from setuptools import setup

OPTIONS = {
	'iconfile':'gorilla_xlsx.icns'
}

setup(
    app=['jim_tracker.pyw'],
    name='Jim Tracker',
    options={'py2app':OPTIONS},
    setup_requires=["py2app"],
)