# setup.py placed at root directory
from setuptools import setup
setup(
    name='my-excel-edire',
    version='0.0.1',
    author='Eric Di Re',
    description='Custom package for manipulating Microsoft Excel documents.',
    url='https://github.com/edire/my_excel.git',
    python_requires='>=3.6',
    packages=['my_excel'],
    install_requires=['pywin32']
)