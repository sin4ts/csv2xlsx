import os

from setuptools import setup

requirement_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'requirements.txt')
install_requires = []
if os.path.isfile(requirement_path):
    with open(requirement_path) as f:
        install_requires = f.read().splitlines()

setup(name = 'csv2xlsx',
    version = '1.0',
    description = 'Convert CSV file to XLSX',
    author = 'sin4ts',
    license= 'MIT License',
    platforms= 'any',
    scripts=['csv2xlsx'],
    author_email = 'stan102@hotmail.fr',
    url = 'https://github.com/sin4ts/csv2xlsx',
    install_requires = install_requires
)
