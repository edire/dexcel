# setup.py placed at root directory
from setuptools import setup
setup(
    name='my-excel-edire',
    version='0.0.2',
    author='Eric Di Re',
    description='Custom package for manipulating Microsoft Excel documents.',
    url='https://github.com/edire/my_excel.git',
    python_requires='>=3.6',
    packages=['my_excel'],
    package_data={'my_excel': ['MiscFiles/SharedVBA.bas']},
    install_requires=['pywin32']
)