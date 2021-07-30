from setuptools import setup, find_packages
import os.path
import re

# reading package's version (same way sqlalchemy does)

setup(
    name='async_requet',
    author='Mohamad Khajezade',
    author_email='khajezade.mohamad@gmail.com',
    description='A package to read information from excell and write them in a docx',
    packages=find_packages(),
)
