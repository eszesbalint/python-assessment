from setuptools import setup, find_packages

setup(
    name='slidepy',
    version='1.0',
    author='Eszes Bálint',
    author_email='eszes.balint@pm.me',
    url='https://github.com/eszesbalint/python-assessment',
    packages=find_packages(),
    package_dir={'slidepy': 'slidepy'},
    install_requires=[
        'docstring-parser',
        'numpy',
        'typeguard',
        'matplotlib',
        'python-pptx'
    ],
)