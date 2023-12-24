# Import necessary libraries
import datetime
import glob
import os
import time

import pandas as pd

# Define a list of file paths to monitor for changes
files_to_watch = glob.glob('input/*.xlsx')

# Create a dictionary to record the timestamp of each file
file_timestamps = {f: os.path.getmtime(f) for f in files_to_watch}


# Define a procedure to process these files
def process_files():
    # Read the first column from each excel file as a different pandas DataFrame
    # We make use of the pandas read_excel function to accomplish this task
    df1 = pd.read_excel('input/1.xlsx', usecols=[0], parse_dates=[0])
    df2 = pd.read_excel('input/2.xlsx', usecols=[1], dtype={'B': int})
    df3 = pd.read_excel('input/3.xlsx', usecols=[2], dtype={'C': int})

    # Rename the columns of each DataFrame to a more recognizable name
    df1.columns = ['Date']
    df2.columns = ['Value1']
    df3.columns = ['Value2']

    # Make sure the number of rows in df2 and df3 match df1 by truncating excess rows
    df2 = df2[:len(df1)]
    df3 = df3[:len(df1)]

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

    # Create a workbook and chartsheet for adding charts
    workbook = writer.book
    chartsheet = workbook.add_worksheet('Charts')

    # Create a line chart for the data in 'Value1' and 'Value2' columns
    chart = workbook.add_chart({'type': 'line'})

    # Add a series for each of the 'Value1' and 'Value2' columns to the line chart
    for i, col_name in enumerate(['Value1', 'Value2']):
        chart.add_series({
            'name': col_name,
            'categories': f'Data!$A$2:$A${len(final_df) + 1}',
            'values': f'Data!${chr(66 + i)}$2:${chr(66 + i)}${len(final_df) + 1}',
        })

    # Set the size of the created chart
    chart.set_size({'width': 1080, 'height': 576})

    # Insert the line chart into the created chartsheet at cell 'B2'
    chartsheet.insert_chart('B2', chart)

    # Create a pie chart for the summary data
    pie_chart = workbook.add_chart({'type': 'pie'})

    # Aggregate the values of each column and save the result into a new DataFrame
    sum_df = final_df[['Value1', 'Value2']].sum().to_frame()

    # Rename the column of the created DataFrame to 'Value'
    sum_df.columns = ['Value']

    # Write this DataFrame into the same excel file, on a sheet named 'Summary'
    sum_df.to_excel(writer, sheet_name='Summary', index_label='Label', index=True)

    # Add a series to the pie chart using the summary data
    pie_chart.add_series({
        'name': 'Pie Data',
        'categories': 'Summary!$A$2:$A$3',
        'values': 'Summary!$B$2:$B$3',
    })

    # Insert the pie chart into the created chartsheet at cell 'B32'
    chartsheet.insert_chart('B32', pie_chart)

    # Save and close the ExcelWriter object
    writer.close()

    # Print the name of the generated file
    print('Output file generated: {}'.format(filename))


# Process the files for the first time
process_files()

# Create a loop that keeps watching for changes in the specified files
while True:
    # Check if any file has been modified
    any_modified = any(os.path.getmtime(f) != file_timestamps[f] for f in files_to_watch)

    # If any file has been modified
    if any_modified:
        # Print a message indicating that files have been modified
        print('Files modified.')

        # Re-process the files
        process_files()

        # Update the timestamps of the files
        file_timestamps = {f: os.path.getmtime(f) for f in files_to_watch}

    # Wait for 1 second before checking again
    time.sleep(1)
