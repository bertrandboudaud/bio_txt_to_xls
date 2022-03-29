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

column_mapping = [
#   Excel column names           Txt file column names
    ('Index',                    'Index'), 
    ('Plate position',           '-'),
    ('Sample Name',              'Sample Name'),
    ('Sample Type',              'Sample Type'),
    ('Component Name',           'Component Name'),
    ('Component Group Name',     'Component Group Name'),
    ('Component Type',           'Component Type'),
    ('Dilution Factor',          'Dilution Factor'),
    ('Expected RT',              'Expected RT'),
    ('Actual RT',                'Retention Time'),
    ('RT Delta (min)',           'Retention Time Delta (min)'),
    ('Area',                     'Area'),
    ('Height',                   'Height'),
    ('Height Ratio',             'Height Ratio'),
    ('Area / Height',            'Area / Height'),
    ('Height Ratio',             'Height Ratio'),
    ('Calculated Concentration', 'Calculated Concentration'),
    ('Concentration acceptance', 'Concentration Acceptance'),
    ('-',                        'Used'),
    ('-',                        'Accuracy'),
    ('-',                        'Accuracy Acceptance'),
    ('Peak Width Confidence',    'AutoPeak Peak Width Confidence' ),
    ('Peak Saturated',           'AutoPeak Saturated')
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

# create excel file
workbook = xlsxwriter.Workbook(args.Output, {'nan_inf_to_errors': True})

# cell styles
#  headers
format_header = workbook.add_format()
format_header.set_align('center')
format_header.set_bg_color('silver')
format_header.set_center_across()
#  headers with one border on the right
format_header_right_border = workbook.add_format()
format_header_right_border.set_align('center')
format_header_right_border.set_bg_color('silver')
format_header_right_border.set_right()
#  standard cell
format_value = workbook.add_format()
#  standard cell with one border on the right
format_value_right_border = workbook.add_format()
format_value_right_border.set_right()

# 1st sheet
current_sheet = workbook.add_worksheet("SciexOS")
current_column = 0
for (excel_column_name, txt_column_name) in column_mapping:
    if excel_column_name != "-" and excel_column_name in csv_data:
        current_sheet.write(0, current_column, excel_column_name, format_header)
        current_column += 1 
current_line = 1
for index in csv_data.index:
    current_column = 0
    for (excel_column_name, txt_column_name) in column_mapping:
        test = excel_column_name in csv_data
        if excel_column_name != "-" and excel_column_name in csv_data:
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

sheets_description = {
    "Height" : [
        ("IS | Heavy", "Height", "Heavy", False),
        ("Light", "Height", "Light", True),
        ("Height Ratio", "Height Ratio", "Light", True),
        ("Conc. µM", "Calculated Concentration", "Light", False),
        ("Conc. Acceptance", "Concentration acceptance", "Light", True)
        ],
    "Area" : [
        ("IS | Heavy", "Area", "Heavy", False),
        ("Light", "Area", "Light", True)
        ]
    }

def writeFeature(current_sheet, current_column, title, separator, values, feature, light_or_heavy, sample_names):
    current_sheet.write(1, current_column, title, format_header_right_border)
    values_filtered = values.loc[csv_data['Component Name'].str.endswith(light_or_heavy)]
    if sample_names is None:
        sample_names = values_filtered["Sample Name"]
    else:
        if not (sample_names.values == values_filtered["Sample Name"].values).all():
            raise Exception("Sample Name inconstency! aborting.")
    current_line = 2
    for index in values_filtered.index:
        value = values_filtered[feature][index]
        if separator:
            cell_format = format_value_right_border
        else:
            cell_format = format_value
        if pd.isna(value):
            current_sheet.write(current_line, current_column, "N/A", cell_format)
        else:
            current_sheet.write(current_line, current_column, value, cell_format)
        current_line += 1
    return sample_names

for sheet_title in sheets_description:
    current_sheet = workbook.add_worksheet(sheet_title)
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
        starting_group_column = current_column
        values = csv_data.loc[csv_data['Component Group Name'] == sample_group]
        for (title, feature, light_or_heavy, separator) in sheets_description[sheet_title]:
            sample_names = writeFeature(current_sheet, current_column, title, separator, values, feature, light_or_heavy, sample_names)
            current_column += 1 
        current_sheet.merge_range(0, starting_group_column, 0, current_column - 1, sample_group, format_header_right_border)
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
print("File written to " + str(args.Output))