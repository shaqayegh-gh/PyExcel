from setuptools import setup, find_packages
setup(
    name='pyexcel',
    version='0.0.1',
    author='Shaghayegh Ghorbanpoor',
    author_email='ghorbanpoor.shaghayegh@gmail.com',
    packages=find_packages(include=['pyexcel', 'pyexcel.*']),
    description='This is a package for creating or reading excel',
    requires=['xlwt==1.3.0', 'pytest', 'openpyxl==3.1.1']
)