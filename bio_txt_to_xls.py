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
import copy

parser = argparse.ArgumentParser(description='bio_txt_toxls, script to ease analisys from a csv file to excel file.',
                                 formatter_class=argparse.ArgumentDefaultsHelpFormatter)

# Required arguments
parser.add_argument('Input',
                    type=pathlib.Path,
                    help='Input csv file')
parser.add_argument('Output',
                    type=pathlib.Path,
                    help='.xlsx file')

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

format_header = workbook.add_format()
format_header.set_align('center')
format_header.set_bg_color('silver')
format_header.set_center_across()

format_header_right_border = workbook.add_format()
format_header_right_border.set_align('center')
format_header_right_border.set_bg_color('silver')
format_header_right_border.set_right()

format_value = workbook.add_format()

format_value_right_border = workbook.add_format()
format_value_right_border.set_right()


# 1st sheet
current_sheet = workbook.add_worksheet("SciexOS")
current_column = 0
for (excel_column_name, txt_column_name) in column_mapping:
    if excel_column_name != "-":
        current_sheet.write(0, current_column, excel_column_name, format_header)
        current_column += 1 
current_line = 1
for index in csv_data.index:
    current_column = 0
    for (excel_column_name, txt_column_name) in column_mapping:
        if excel_column_name != "-":
            value = csv_data[excel_column_name][index]
            if pd.isna(value):
                current_sheet.write(current_line, current_column, "N/A", format_value)
            else:
                current_sheet.write(current_line, current_column, value, format_value)
            current_column += 1 
    current_line += 1
current_sheet.freeze_panes(1, 0)
current_sheet.autofilter(0, 0, current_line,  current_column)

# other sheets
current_sheet = workbook.add_worksheet("Test")
current_column = 0
current_sheet.write(1, current_column, "Plate position", format_header_right_border)
current_column += 1
current_sheet.write(1, current_column, "Sample Name", format_header_right_border)
current_column += 1
current_sheet.write(1, current_column, "Dilution Factor", format_header_right_border)
current_column += 1
sample_groups = csv_data['Component Group Name'].drop_duplicates().sort_values()
sample_names = None
for sample_group in sample_groups:
    current_sheet.merge_range(0, current_column, 0, current_column +1, sample_group, format_header_right_border)
    values = csv_data.loc[csv_data['Component Group Name'] == sample_group]
    current_sheet.write(1, current_column, "IS | Heavy", format_header)
    values_heavy = values.loc[csv_data['Component Name'].str.endswith("Heavy")]
    if sample_names is None:
        sample_names = values_heavy["Sample Name"]
    else:
        if not (sample_names.values == values_heavy["Sample Name"].values).all():
            raise Exception("Sample Name inconstency! aborting.")
    current_line = 2
    for index in values_heavy.index:
        value = values_heavy["Area"][index]
        if pd.isna(value):
            current_sheet.write(current_line, current_column, "N/A", format_value)
        else:
            current_sheet.write(current_line, current_column, value, format_value)
        current_line += 1
    current_column += 1 
    current_sheet.write(1, current_column, "Light", format_header_right_border)
    values_light = values.loc[csv_data['Component Name'].str.endswith("Light")]
    if sample_names is None:
        sample_names = values_heavy["Sample Name"]
    else:
        if not (sample_names.values == values_heavy["Sample Name"].values).all():
            raise Exception("Sample Name inconstency! aborting.")
    current_line = 2
    for index in values_light.index:
        value = values_light["Area"][index]
        if pd.isna(value):
            current_sheet.write(current_line, current_column, "N/A", format_value_right_border)
        else:
            current_sheet.write(current_line, current_column, value, format_value_right_border)
        current_line += 1
    current_column += 1
current_line = 2
for sample_name in sample_names:
    current_sheet.write(current_line, 1, sample_name, format_value_right_border)
    dilutions = csv_data.loc[csv_data['Sample Name'] == sample_name]["Dilution Factor"]
    if not (dilutions == dilutions.iloc[0]).all():
        raise Exception("Dilution Factor inconstency! aborting.")
    current_sheet.write(current_line, 2, dilutions.iloc[0], format_value_right_border)
    current_line += 1
current_sheet.freeze_panes(2, 3)


# end
workbook.close()
print("End of script")