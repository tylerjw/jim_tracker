"""
Script for building the example.

Usage:
    python setup.py py2app
"""
from distutils.core import setup
import py2exe

setup(
    windows=['notebook_tracker.pyw'],
)