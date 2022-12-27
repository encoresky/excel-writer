from setuptools import setup
from distutils.util import convert_path

main_ns = {}
ver_path = convert_path('excel_writer/version')
with open(ver_path) as ver_file:
    exec(ver_file.read(), main_ns)

setup(
    name="excel_writer",
    version=main_ns['__version__'],
    license='MIT',
    description="Simple library to write excel files from dictionary or pandas datafra",
    author="EncoreSky Technologies",
    author_email='coontact@encoresky.com',
    url='https://github.com/encoresky/excel-writer',
    packages=[
        "excel_writer"
    ],
    include_package_data=True,
    install_requires=[
        "XlsxWriter==3.0.3",
        "pandas==1.5.2",
        "openpyxl==3.0.10",
        "Jinja2==3.1.2"
    ],
    dependency_links=[

    ],
    entry_points={

    }
)