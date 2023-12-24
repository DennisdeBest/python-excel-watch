# Python Data Processing and Visualization Project

This project focuses on reading data from multiple `.xlsx` files, processing and visualizing it using Python.

## Description

- The script continuously watches a specific directory for any modifications in `.xlsx` files.
- When a modification is detected, it reads data from these Excel files, each one containing specific columns with different data types.
- It renames the columns to avoid any confusion due to conflicted column names.
- It then combines all the data from these files into a single DataFrame for further processing.
- The processed data is saved into an Excel file with the timestamp in the filename.
- A column chart is also created within the final Excel file itself using the `xlsxwriter` library.

## Libraries Used

- pandas: For data manipulation and analysis.
- xlsxwriter: An Excel file writing library.

## Installation

- Make sure Python 3.11.6 is installed.
- Additional packages needed are:
    - pandas
    - xlsxwriter
    - openpyxl

To setup a virtual environment and install these packages, you can follow these steps:

### Create a new virtual environment

```shell
python3 -m venv venv
```

### Activate the virtual environment

```shell
source venv/bin/activate
```

### Install the necessary packages

```shell
pip install pandas xlsxwriter openpyxl
```

## Usage

- Update the `files_to_watch` variable in the script to match your dataset (.xlsx files you want to watch for modifications).
- Run the script using `python3 script.py` while the virtual environment is activated.

## Output

- The final output will be an Excel file named `<timestamp>.xlsx` containing the combined data. The `<timestamp>` in the filename will be the date and time when the files were processed.
- The Excel file will have a column chart embedded into it, visualizing the data.