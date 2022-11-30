# setup.py placed at root directory
from setuptools import setup
setup(
    name='dexcel',
    version='0.0.1',
    author='Eric Di Re',
    description='Custom package for manipulating Microsoft Excel documents.',
    url='https://github.com/edire/dexcel.git',
    python_requires='>=3.9',
    packages=['dexcel'],
    package_data={'dexcel': ['MiscFiles/SharedVBA.bas']},
    install_requires=['pywin32']
)