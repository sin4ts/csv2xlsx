# Requirements
Only the xlsxwriter library is required.

# Usage
Use as a standalone script:
 python csv2xlsx.py FILE.csv -d ',' -q '"'

Use as a library:
 import csv2xlsx

 csv2xslx.process_file('')
 csv2xslx.process_directory()

# TODO
* Detect if delimter and quotechar are used and suggest new value if not
* Installation with setup.py
* Test on Windows
* Feature to add a CSV file to an existing XLSX file
