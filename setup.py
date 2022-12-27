from setuptools import setup

setup(
    name="excel_writer",
    version="1.0.3",
    license='MIT',
    description="Simple library to write excel files from Python Dictionary or Pandas DataFrame.",
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
