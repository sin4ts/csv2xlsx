import os
import shutil

from setuptools import setup

requirement_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'requirements.txt')
install_requires = []
if os.path.isfile(requirement_path):
    with open(requirement_path) as f:
        install_requires = f.read().splitlines()

script_list = []
if os.name == 'posix':
    shutil.copy('csv2xlsx.py', 'csv2xlsx')
    script_list = ['csv2xlsx']


setup(name = 'csv2xlsx',
    version = '1.1',
    description = 'Convert CSV file to XLSX',
    author = 'sin4ts',
    license = 'MIT License',
    platforms = 'any',
    scripts = script_list,
    author_email = 'stan102@hotmail.fr',
    url = 'https://github.com/sin4ts/csv2xlsx',
    install_requires = install_requires
)

if os.name == 'posix':
    os.unlink('csv2xlsx')
