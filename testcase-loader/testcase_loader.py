import time
import xlwings as xw
import pandas as pd
import gc
import sys

start_time = time.time() 

# function to copy from one workbook and paste to another given a shared table name
def copy_table_data(input_wb, tool_wb, input_sheet_name, tool_sheet_name, table_name):
    input_sheet = input_wb.sheets[input_sheet_name]
    tool_sheet = tool_wb.sheets[tool_sheet_name]

    input_table_range = input_sheet.tables[table_name].range
    tool_table_range = tool_sheet.tables[table_name].range

    # extract data excluding headers
    input_data = input_table_range.offset(1, 0).resize(
        input_table_range.rows.count - 1,
        input_table_range.columns.count
    ).value

    # destination start (below headers)
    tool_data_start = tool_table_range.offset(1, 0)

    # clear existing data
    tool_data_start.resize(
        tool_table_range.rows.count - 1,
        tool_table_range.columns.count
    ).clear_contents()
    
    # set new data
    tool_data_start.resize(len(input_data), len(input_data[0])).value = input_data

# function to open workbook, not sure if I need this it might be best practice to just assume/have the workbooks are closed
def get_workbook(app, file_name, file_path):
    try:
        return xw.books[file_name]
    except (KeyError, xw.XlwingsError):
        return app.books.open(file_path)

tool_file = 'Commissions ALIP 2.0 FIA v35.xlsm'
tool_path = r'C:\Users\AZL68MG\source\repos\PythonApplication1\Commissions ALIP 2.0 FIA v35.xlsm'

input_file = 'TestCasesToLoad.xlsm'
input_path = r'C:\Users\AZL68MG\source\repos\PythonApplication1\TestCasesToLoad.xlsm'

try:

    app = xw.App(visible=False)  

    wb_input = get_workbook(app, input_file, input_path)
    print('Opened input file')

    wb_tool = get_workbook(app, tool_file, tool_path)
    print('Opened output file')

    #input_sheet, tool_sheet, shared table_name
    tables_to_copy = [
        ('Producer Info', 'Producer Info', 'Hierarchies'),
        ('Producer Info', 'Producer Info', 'SupplementalComp'),
        ('Producer Info', 'Producer Info', 'ProdPrefs'),
        ('(Q) NG ALIP', '(Q) NG ALIP', 'NG_Transactions')
    ]


    for input_sheet, tool_sheet, table_name in tables_to_copy:
        copy_table_data(wb_input, wb_tool, input_sheet, tool_sheet, table_name)

    sheet = wb_input.sheets['(Q) NG ALIP']
    table_range = sheet.tables['NG_Transactions'].range

    df = table_range.options(pd.DataFrame, header=True, index=False).value
    pol_nums = df['POLICY_NBR'].unique()

    tool_sheet = wb_tool.sheets['Input']

    tool_sheet.range('K2:K21').value = [None] * 20

    for i, value in enumerate(pol_nums, start=2): # Start at row 2
        tool_sheet.range(f'K{i}').value = value

    print(pol_nums)
    

    # adding a comment here
    # more comments
    # even more comments

    # RUN THE TOOL!!!
    wb_tool.macro('Clear_Output')()
    wb_tool.macro('RunTool')()
    print ('Ran the Tool!')

    output_sheet = 'Commissions Output'
    output_table = 'Output_table'

    df = wb_tool.sheets[output_sheet].tables[output_table].range.options(pd.DataFrame, header=True, index=False).value
    print(df)

    # not the best way to quit as workbooks are still being referenced in memory:
    # app.quit()

finally:

    wb_tool.save()

    wb_input.close()
    wb_tool.close()

    # delete the objects
    del wb_input, wb_tool

    # force garbage collection
    gc.collect()

    elapsed_time = time.time() - start_time
    print(f"\n Script completed in {elapsed_time: .2f} seconds.")

