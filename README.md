# Description
This project is a wrapper of the xslxwriter project, in order to easily convert a single CSV file, or mutliple CSV files into XSLX files. It can be used both as a python library, or as a standalone script.

# Requirements
Only the xlsxwriter python library is required: https://github.com/jmcnamara/XlsxWriter

# Installation
## Windows
```
python3.exe setup.py install
```
## Linux
```
sudo make install
```

# Usage
## Windows
```
python3.exe csv2xlsx.py DIRECTORY ADDITIONAL_FILE.csv -d '|' -q none
```

## Linux
```
python3 csv2xlsx.py DIRECTORY ADDITIONAL_FILE.csv -d '|' -q none
csv2xlsx DIRECTORY ADDITIONAL_FILE.csv -d '|' -q none
```
## Python
Use as a library:
```
import csv2xlsx
csv2xslx.process_file('dir/file.csv')
csv2xslx.process_directory('.')
```

# Doc
## csv2xlsx [python standalone script]
```
usage: csv2xlsx.py [-h] [-o OUTPUT] [-d DELIMITER] [-q QUOTECHAR] [-e ENCODING] [--no-header]
                   [--no-verify] [--no-merge] [-f FILTER] [-O] [-r]
                   input [input ...]

positional arguments:
  input                 Input CSV file or directory

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Output XLSX filepath (default: csv2xlsx-output.xlsx)
  -d DELIMITER, --delimiter DELIMITER
                        CSV delimiter character. Default is ",". Set to TAB to use tabulation as
                        delimiter
  -q QUOTECHAR, --quotechar QUOTECHAR
                        CSV quotechar character. Default: " . Set to NONE to disable quotechar
  -e ENCODING, --encoding ENCODING
                        File encoding. Default is UTF-8
  --no-header           Don't process first row as header
  --no-verify           Don't verify CSV file consistency
  --no-merge            Don't merge files into a single XLSX file
  -f FILTER, --filter FILTER
                        Filename filter with regex. Default is '.*\.[C^c][S^s][V^v]$')
  -O, --overwrite       Overwrite existing file
  -r, --recurse         Process directories recursively
```

## process_csv [python function]
* row_list [list] [mandatory]: List of list to import in a worksheet
* output_path [string] [optional]: Path of the output workbook to use
* workbook [xlswwrtier.Workbook] [optional]: Existing workbook to use
* title [string] [optional]: Title of the worksheet to create
* auto_size [bool] [optional]: Auto size every columns
* filter_first_row [bool] [optional]: Filter the first row, if header flag is enabled (Default = True)
* freeze_first_row [bool] [optional]: Freeze the first row, if header flag is enabled (Default = True)
* header [bool] [optional]: Process the first row as headers (Default = True)
* close_workbook [bool] [optional]: Close the workbook object and return the created filename. If set to False, then the workbook object is returned (Default = True)
* overwrite [bool] [optional]: Overwrite the output workbook in case the file already exists (Default = False)

## process_file [python function]
* input_path [string] [mandatory]: List of list to import in a worksheet
* output_path [string] [optional]: Path of the output workbook to use
* workbook [xlswwrtier.Workbook] [optional]: Existing workbook to use
* title [string] [optional]: Title of the worksheet to create
* encoding [string] [optional]: Encoding to use for the provide filepath (Default = UTF-8)
* auto_size [bool] [optional]: Auto size every columns (Default = True)
* filter_first_row [bool] [optional]: Filter the first row, if header flag is enabled (Default = True)
* freeze_first_row [bool] [optional]: Freeze the first row, if header flag is enabled (Default = True)
* header [bool] [optional]: Process the first row as headers (Default = True)
* close_workbook [bool] [optional]: Close the workbook object and return the created filename. If set to False, then the workbook object is returned (Default = True)
* overwrite [bool] [optional]: Overwrite the output workbook in case the file already exists (Default = False)
* delimiter [string] [optional]: delimiter character to use for reading the CSV file (Default = ',')
* quotechar [string] [optional]: quotechar character to use for reading the CSV file  (Default = '"')
* verify_csv_data [bool] [optional]: Verify if delimiter and quotechar are found in the file (Default is False)

## process_directory [python function]
* input_path [string] [mandatory]: List of list to import in a worksheet
* output_path [string] [optional]: Path of the output workbook to use
* workbook [xlswwrtier.Workbook] [optional]: Existing workbook to use
* encoding [string] [optional]: Encoding to use for the provide filepath (Default = UTF-8)
* auto_size [bool] [optional]: Auto size every columns (Default = True)
* filter_first_row [bool] [optional]: Filter the first row, if header flag is enabled (Default = True)
* freeze_first_row [bool] [optional]: Freeze the first row, if header flag is enabled (Default = True)
* header [bool] [optional]: Process the first row as headers (Default = True)
* close_workbook [bool] [optional]: Close the workbook object and return the created filename. If set to False, then the workbook object is returned (Default = True)
* overwrite [bool] [optional]: Overwrite the output workbook in case the file already exists (Default = False)
* delimiter [string] [optional]: delimiter character to use for reading the CSV files (Default = ',')
* quotechar [string] [optional]: quotechar character to use for reading the CSV files  (Default = '"')
* verify_csv_data [bool] [optional]: Verify if delimiter and quotechar are found in the files (Default = False)
* merge [bool] [optional]: Merge all CSV files into a single workbook (Default = True)
* file_regex [string] [optional]: Regex to filter file to import, based on the file name (Default = '.*\\.[C^c][S^s][V^v]$')

# TODO
* Feature to add a CSV file to an existing XLSX file
