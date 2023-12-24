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

The example take 3 input files, the first with just a column of dates and the other 2 with columns of integers.
The final output amount of rows will be limited to the amount of rows in the first Excel file.
The final output has 3 sheets:
- Data: 3 columns with the date values from file 1 and the values from file 2 and 3
- Charts: A line chart based on the dates of the file 1 and the values of the other files and a pie chart showing the difference of the totals of file 2 and 3.
- Summary: A Sheet with the sums of the values from the data chart used to generated the pie chart

To change these values you can : 

- Set the Excel files and the columns from which data needs to be gathered
  - ```python
    df1 = pd.read_excel('input/1.xlsx', usecols=[0], parse_dates=[0])
    df2 = pd.read_excel('input/2.xlsx', usecols=[1], dtype={'B': int})
    df3 = pd.read_excel('input/3.xlsx', usecols=[2], dtype={'C': int})
  
- Name the output columns and set the amount of rows to those of one of the dataframes
  - ```python
    df1.columns = ['Date']
    df2.columns = ['Value1']
    df3.columns = ['Value2']

    # Make sure the number of rows in df2 and df3 match df1 by truncating excess rows
    df2 = df2[:len(df1)]
    df3 = df3[:len(df1)]

- Concatenate the columns into a single DataFrame, set the name of the output file and write the final file
  - ```python
    # Concatenate the three DataFrames into a single one along the column axis
    final_df = pd.concat([df1, df2, df3], axis=1)

    # Define the path where the output file will be saved. The filename is the current timestamp
    filename = 'output/{}.xlsx'.format(datetime.datetime.now().strftime("%Y%m%d%H%M%S"))

    # Convert the 'Date' column to string type
    final_df['Date'] = final_df['Date'].astype(str)

    # Create a pandas ExcelWriter object with the previously defined filename,
    # and specify the engine as xlsxwriter to enable the addition of charts
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    # Write the DataFrame to an excel file, on a sheet named 'Data', and without row index
    final_df.to_excel(writer, sheet_name='Data', index=False)

- Run the script using `python3 script.py` while the virtual environment is activated.

## Output

- The final output will be an Excel file named `<timestamp>.xlsx` containing the combined data. The `<timestamp>` in the filename will be the date and time when the files were processed.
- The Excel file will have a column chart embedded into it, visualizing the data.