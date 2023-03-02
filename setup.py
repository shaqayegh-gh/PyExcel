from setuptools import setup, find_packages
setup(
    name='PyExcel',
    version='0.0.1',
    author='Shaghayegh Ghorbanpoor',
    author_email='ghorbanpoor.shaghayegh@gmail.com',
    packages=find_packages(include=['PyExcel', 'PyExcel.*']),
    description='This is a package for creating or reading excel',
)