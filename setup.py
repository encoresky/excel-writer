from setuptools import setup

from pathlib import Path
this_directory = Path(__file__).parent
long_description = (this_directory / "README_PyPi.md").read_text()

setup(
    name="excel_writer",
    version="1.0.5",
    license='MIT',
    description="Simple library to write excel files from Python Dictionary or Pandas DataFrame.",
    long_description=long_description,
    long_description_content_type='text/markdown',
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
