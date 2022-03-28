# MIT License
#
# Copyright (c) 2022 bertrandboudaud
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import argparse
import pathlib
from numpy import NaN
import pandas as pd
import xlsxwriter

parser = argparse.ArgumentParser(description='bio_txt_toxls, script to ease analisys from a csv file to excel file.',
                                 formatter_class=argparse.ArgumentDefaultsHelpFormatter)

# Required arguments
parser.add_argument('Input',
                    type=pathlib.Path,
                    help='Input csv file')
parser.add_argument('Output',
                    type=pathlib.Path,
                    help='.xlsx file')
#parser.add_argument('Features',
#                    type=str,
#                    default="detect",
#                    help='List of features to expose')

column_mapping = [
#   Excel column names              Txt file column names
    ('Index',                       'Index'), 
    ('Plate position',              '-'),
    ('Sample Name',                 'Sample Name'),
    ('Sample Type',                 'Sample Type'),
    ('Component Name',              'Component Name'),
    ('Component Group Name',        'Component Group Name'),
    ('Component Type',              'Component Type'),
    ('Dilution Factor',             'Dilution Factor'),
    ('Expected RT',                 'Expected RT'),
    ('Actual RT',                   'Retention Time'),
    ('RT Delta (min)',              'Retention Time Delta (min)'),
    ('Area',                        'Area'),
    ('Height',                      'Height'),
    ('Height Ratio',                'Height Ratio'),
    ('Area / Height',               'Area / Height'),
    ('Height Ratio',                'Height Ratio'),
    ('Calculated Concentration',    'Calculated Concentration'),
    ('Concentration acceptance',    'Concentration Acceptance'),
    ('-',                           'Used'),
    ('-',                           'Accuracy'),
    ('-',                           'Accuracy Acceptance')
]

args = parser.parse_args()

csv_data = pd.read_csv(args.Input, sep='\t', decimal=',')

# rename columns
renaming_mapping = {}
for (excel_column_name, txt_column_name) in column_mapping:
    if (excel_column_name != "-"):
        renaming_mapping[txt_column_name] = excel_column_name
csv_data.rename(columns=renaming_mapping,  inplace=True)
print(csv_data)

# add empty colums
for (excel_column_name, txt_column_name) in column_mapping:
    if (txt_column_name == "-"):
        csv_data[excel_column_name] = " "

dataTypeSeries = csv_data.dtypes
print('Data type of each column of Dataframe :')
print(dataTypeSeries)

# write excel file
workbook = xlsxwriter.Workbook(args.Output, {'nan_inf_to_errors': True})

cell_format_table = workbook.add_format()
cell_format_table.set_align('center')
cell_format_table.set_align('top')
cell_format_table.set_center_across()

cell_format_line = workbook.add_format()

# 1st sheet
current_sheet = workbook.add_worksheet("SciexOS")
current_column = 0
for (excel_column_name, txt_column_name) in column_mapping:
    if excel_column_name != "-":
        current_sheet.write(0, current_column, excel_column_name, cell_format_table)
        current_column += 1 
current_line = 1
for index in csv_data.index:
    current_column = 0
    for (excel_column_name, txt_column_name) in column_mapping:
        if excel_column_name != "-":
            value = csv_data[excel_column_name][index]
            if pd.isna(value):
                current_sheet.write(current_line, current_column, "N/A", cell_format_line)
            else:
                current_sheet.write(current_line, current_column, value, cell_format_line)
            current_column += 1 
    current_line += 1
current_sheet.freeze_panes(1, 0)
current_sheet.autofilter(0, 0, current_line,  current_column)

# other sheets
current_sheet = workbook.add_worksheet("Test")
current_column = 0
current_sheet.write(1, current_column, "Plate position", cell_format_table)
current_column += 1
current_sheet.write(1, current_column, "Sample Name", cell_format_table)
current_column += 1
current_sheet.write(1, current_column, "Dilution Factor", cell_format_table)
current_column += 1
sample_groups = csv_data['Component Group Name'].drop_duplicates().sort_values();
for sample_group in sample_groups:
    current_sheet.write(0, current_column, sample_group, cell_format_table)
    values = csv_data.loc[csv_data['Component Group Name'] == sample_group]
    current_sheet.write(1, current_column, "IS | Heavy", cell_format_table)
    values_heavy = values.loc[csv_data['Component Name'].str.endswith("Heavy")]
    current_line = 2
    for index in values_heavy.index:
        area = values_heavy["Area"][index]
        current_sheet.write(current_line, current_column, area, cell_format_line)
        current_line += 1
    current_column += 1 
    current_sheet.write(1, current_column, "Light", cell_format_table)
    values_light = values.loc[csv_data['Component Name'].str.endswith("Light")]
    current_line = 2
    for index in values_light.index:
        area = values_light["Area"][index]
        current_sheet.write(current_line, current_column, area, cell_format_line)
        current_line += 1
    current_column += 1 

# end
workbook.close()
print("End of script")