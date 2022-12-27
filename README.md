Copyright (c) 2023-Present EncoreSky Technologies Pvt. Ltd. All rights reserved.
DeepCognito EvolutionAI

# Excel Writer

![version](https://img.shields.io/badge/version-1.0.0-blue)

*excel_writer* is a package containing the methods for creating excel file from python dictionary or pandas dataframe.

## Prerequisites

In order to clone the repository, install the framework and the dependencies you need network access.

You'll also need the following:

- [Git](https://git-scm.com)
- [Python](https://www.python.org/downloads/release/python-3811/), _excel_writer_ is compatible with version >=3.8.11.

## Getting the sources

First clone the repository:

```bash
git clone https://github.com/encoresky/excel-writer.git
```

## Installation

To install the package (in either a system-wide or a virtual environment), navigate to the *excel_writer* root folder in a Terminal, and type:

```bash
python3 setup.py install
```

*excel_writer* will be installed as a package in your Python distribution, along with it's dependencies if necessary.

*N.B.* - installing in a virtual environment is recommended.

## Usage

### Write a file using dictionary


```python
from excel_writer import ExcelWriter

input_data = {
    "Name": ["Aarav", "Jayesh", "Vineet", "Rahul", "Mayank", "Deepak"],
    "Age": [24, 28, 27, 25, 28, 35],
    "Emp Id": ["A-001", "A-002", "A-003", "A-004", "A-005", "A-006"],
    "City": ["Indore", "Bhopal", "Gwalior", "Pune", "Kolkata", "Udaipur"],
    "Has Bike": ["Y", "Y", "Y", "Y", "Y", "Y"]
}

excel_writer = ExcelWriter(file_name="test.xlsx", 
                            sheet_name="sheet_1")

excel_writer.write_data(data_dict = input_data)
```

### Write a file using pandas dataframe


```python
import pandas as pd
from excel_writer import ExcelWriter

input_data = {
    "Name": ["Aarav", "Jayesh", "Vineet", "Rahul", "Mayank", "Deepak"],
    "Age": [24, 28, 27, 25, 28, 35],
    "Emp Id": ["A-001", "A-002", "A-003", "A-004", "A-005", "A-006"],
    "City": ["Indore", "Bhopal", "Gwalior", "Pune", "Kolkata", "Udaipur"],
    "Has Bike": ["Y", "Y", "Y", "Y", "Y", "Y"]
}

data_df = pd.DataFrame(input_data)
excel_writer = ExcelWriter(file_name="test.xlsx", 
                            sheet_name="sheet_1")

excel_writer.write_data(data_df = data_df)

```
