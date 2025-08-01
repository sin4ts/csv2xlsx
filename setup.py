import os
import shutil

from setuptools import setup

def get_version(filepath):
    with open(filepath, 'r') as fd:
        for line in fd.readlines():
            if line.startswith('__version__'):
                return line.split('=')[1].strip().replace('\'', '').replace('"', '')

requirement_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'requirements.txt')
install_requires = []
if os.path.isfile(requirement_path):
    with open(requirement_path) as f:
        install_requires = f.read().splitlines()

source_script_name = 'csv2xlsx.py'
target_script_name = 'csv2xlsx'
version = get_version(source_script_name)

script_list = []
if os.name == 'posix':
    shutil.copy(source_script_name, target_script_name)
    script_list = [target_script_name]
else:
    raise Exception('Not Yet Implemented')


setup(name = 'csv2xlsx',
    version = version,
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
    os.unlink(target_script_name)
