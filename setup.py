"""
Script for building the example.

Usage:
    python setup.py py2app
"""

from setuptools import setup

setup(
    app=['notebook_tracker.pyw'],
    setup_requires=["py2app"],
)