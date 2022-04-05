# bio_txt_to_xls

bio_txt_to_xls, script to ease analisys by exporting csv file to Excel file.

## Requirement

You need to use Python V3.

dependencies:
  * pandas
  * xlsxwriter

To install requirements:

pandas:
```pip install pandas```

XlsWriter:
```pip install XlsxWriter```

## How to use


```bash
usage: bio_txt_to_xls.py [-h] [--Separator SEPARATOR] [--Decimal DECIMAL]
               Input Output {quantitative,qualitative}

```
## Optional Arguments

|short|long|default|help|
| :--- | :--- | :--- | :--- |
|`-h`|`--help`||show this help message and exit|
||`--Separator`|`tab`|Separator used in the input file to separate each column (tab for tabulation).|
||`--Decimal`|`,`|Character used as decimal sign in the input file|


## Example of use:

```python.exe c:\Users\bertr\bio_txt_to_xls\bio_txt_to_xls.py C:/Users/bertr/example/QualitativeAnalysis_export.txt C:/Users/bertr/example/Qualitative_output.xlsx qualitative```

## Using Docker:

```docker run -v c:\Users\bertr\my_files:/tmp bio_txt_to_xls /tmp/QuantitativeAnalysis_export.txt /tmp/QuantitativeAnalysis_output_docker.xlsx quantitative```
