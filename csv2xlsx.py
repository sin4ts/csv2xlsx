#!/usr/bin/python3
# -*- coding: utf-8 -*-

import re
import os
import csv
import sys
import argparse
import xlsxwriter

class FileExistsException(Exception):
    def __init__(self, filepath):
        self.filepath = filepath

    def __str__(self):
        return 'File already exists: {}'.format(self.filepath)

DEFAULT_OUTPUT_FILENAME = 'csv2xlsx-output.xlsx'
MAX_COLUMN_WIDTH = 150
DEFAULT_DELIMITER=','
DEFAULT_QUOTECHAR='"'
DEFAULT_FILE_REGEX=r'.*\.[C^c][S^s][V^v]$'

def verify_csv_file(input_path, delimiter, quotechar):
    common_delimiter_list = [',', ';', '|', '\t']
    common_quotechar_list = ['"', '\'']
    suggested_delimiter = delimiter
    suggested_quotechar = quotechar
    with open(input_path, 'r') as f:
        first_line = f.readline()

        # Detect delimiter
        max_count = len(first_line.split(suggested_delimiter))
        for current_delimiter in common_delimiter_list:
            count = len(first_line.split(current_delimiter))
            if count > max_count:
                max_count = count
                suggested_delimiter = current_delimiter

        # Detect quotechar
        max_count = len(first_line.split(suggested_quotechar))
        for current_quotechar in common_quotechar_list:
            count = len(first_line.split(current_quotechar))
            if count > max_count:
                max_count = count
                suggested_quotechar = current_quotechar

    if suggested_delimiter != delimiter and suggested_quotechar != quotechar:
        print('It looks like the delimiter and quotechar you have specificied are not valid for {}. The following were detected:'.format(input_path))
    elif suggested_delimiter != delimiter:
        print('It looks like the delimiter you have specificied are not valid for {}. The following was detected:'.format(input_path))
    elif suggested_quotechar != quotechar:
        print('It looks like the delimiter you have specificied are not valid for {}. The following was detected:'.format(input_path))

    if suggested_delimiter != delimiter:
        res = 'unset'
        while res.lower().strip() not in ['', 'y', 'n', 'yes', 'no']:
            res = input('Delimiter: {} ? [Y/n] '.format(suggested_delimiter.replace('\t', 'TAB')))
        if res.lower().strip() in ['y', 'yes', '']:
            delimiter = suggested_delimiter

    if suggested_quotechar != quotechar:
        res = 'unset'
        while res.lower().strip() not in ['', 'y', 'n', 'yes', 'no']:
            res = input('Quotchar: {} ? [Y/n] '.format(suggested_quotechar))
        if res.lower().strip() in ['y', 'yes', '']:
            quotechar = suggested_quotechar

    return delimiter, quotechar

def process_csv(row_list, output_path=DEFAULT_OUTPUT_FILENAME, workbook=None, title=None, auto_size=True, filter_first_row=True, freeze_first_row=True, header=True, close_workbook=True, overwrite=False):
    if workbook is None:
        if os.path.exists(output_path):
            if not overwrite:
                raise FileExistsException(output_path)
            else:
                os.unlink(output_path)
        workbook = xlsxwriter.Workbook(output_path)

    if title is not None:
        #Worksheet title is limiter to 31 characters
        title = title[:31]

    worksheet_name_list = [X.name for X in workbook.worksheets()]
    index = 1
    tmp_title = title
    while tmp_title in worksheet_name_list:
        tmp_title = '{t}{i:03d}'.format(t=title[:28], i=index)
        index += 1
    title = tmp_title

    worksheet = workbook.add_worksheet(title)
    row_count = len(row_list)

    if header:
        bold = workbook.add_format({'bold': True})

    max_col_count = 0
    column_width_list = []
    for row_index in range(row_count):
        current_row = row_list[row_index]
        col_count = len(current_row)

        if auto_size:
            # Get size of every cell of the current row for auto sizing later on
            current_column_width_list = [len(X) for X in current_row]
            new_column_width_list = []
            for i in range(max(max_col_count, col_count)):
                # col_count is equal to len(current_row)
                if i >= col_count:
                    new_column_width_list.append(column_width_list[i])
                # max_col_count is equal to len(column_width_list) as it has not yet been updated
                elif i >= max_col_count:
                    new_column_width_list.append(current_column_width_list[i])
                else:
                    new_column_width_list.append(max(column_width_list[i], current_column_width_list[i]))
            column_width_list = new_column_width_list

        max_col_count = max(col_count, max_col_count)

        for col_index in range(col_count):
            if row_index == 0 and header:
                worksheet.write(row_index, col_index, current_row[col_index], bold)
            else:
                value = current_row[col_index]
                try:
                    value = int(value)
                except Exception:
                    pass
                worksheet.write(row_index, col_index, value)

    if header:
        if filter_first_row:
            worksheet.autofilter(0, 0, row_count-1, max_col_count-1)
        if freeze_first_row:
            worksheet.freeze_panes(1, 0)

    if auto_size:
        for index in range(len(column_width_list)):
            column_width = min(column_width_list[index], MAX_COLUMN_WIDTH)
            worksheet.set_column(index, index, column_width)

    if workbook is not None and close_workbook:
        workbook.close()
        return workbook.filename
    else:
        return workbook

def process_file(input_path, output_path=None, workbook=None, title=None, encoding='utf-8', auto_size=True, filter_first_row=True, freeze_first_row=True, header=True, delimiter=DEFAULT_DELIMITER, quotechar=DEFAULT_QUOTECHAR, verify_csv_data=False, overwrite=False, close_workbook=True):
    input_filename = os.path.basename(input_path)
    input_dirname = os.path.dirname(input_path)
    if input_filename.lower().endswith('.csv'):
        output_filename = '{}.xlsx'.format(input_filename[:-4])
        if title is None:
            title = input_filename[:-4]
    else:
        output_filename = '{}.xlsx'.format(input_filename)
    if output_path is not None:
        output_dirname = os.path.dirname(output_path)
        if output_dirname != '' and not os.path.exists(output_dirname):
            os.makedirs(output_dirname)
        if os.path.isdir(output_path):
            output_path = os.path.join(output_path, output_filename)
    else:
        output_path = os.path.join(input_dirname, output_filename)
    if not output_path.lower().endswith('.xlsx'):
        output_path = '{}.xlsx'.format(output_path)
    if title is None:
        title = input_filename

    if verify_csv_data:
        delimiter, quotechar = verify_csv_file(input_path, delimiter, quotechar)


    with open(input_path, 'r', encoding=encoding, errors='ignore') as csvfile:
        csv_reader = csv.reader(csvfile, delimiter=delimiter, quotechar=quotechar)
        row_list = list(csv_reader)

    workbook = process_csv(row_list, output_path=output_path, workbook=workbook, filter_first_row=filter_first_row, freeze_first_row=freeze_first_row, title=title, header=header, close_workbook=False, overwrite=overwrite)
    print('Imported : {}'.format(input_path))

    if workbook is not None and close_workbook:
        workbook.close()
        return workbook.filename
    else:
        return workbook

def process_directory(input_path, output_path=None, encoding='utf-8', header=True, delimiter=DEFAULT_DELIMITER, auto_size=True, filter_first_row=True, freeze_first_row=True, quotechar=DEFAULT_QUOTECHAR, merge=True, verify_csv_data=False, file_regex=DEFAULT_FILE_REGEX, workbook=None, recurse=False, overwrite=False, close_workbook=True):
    for entry in os.listdir(input_path):
        path = os.path.join(input_path, entry)
        if os.path.isdir(path) and recurse:
            workbook = process_directory(path, output_path, encoding=encoding, header=header, auto_size=auto_size, filter_first_row=filter_first_row, freeze_first_row=freeze_first_row, delimiter=delimiter, quotechar=quotechar, verify_csv_data=verify_csv_data, merge=merge, file_regex=file_regex, workbook=workbook, recurse=recurse, close_workbook=False)
            if workbook is not None and not merge:
                print('Data written to {}'.format(workbook.filename))
                workbook.close()
                workbook=None
        elif os.path.isfile(path) and (file_regex is None or file_regex.strip() == '' or re.match(file_regex, entry)):
            workbook = process_file(path, output_path, encoding=encoding, workbook=workbook, delimiter=delimiter, quotechar=quotechar, filter_first_row=filter_first_row, freeze_first_row=freeze_first_row, verify_csv_data=verify_csv_data, header=header, overwrite=overwrite, close_workbook=False)
            if workbook is not None and not merge:
                print('Data written to {}'.format(workbook.filename))
                workbook.close()
                workbook=None

    if workbook is not None and close_workbook:
        workbook.close()
        return workbook.filename
    else:
        return workbook

def run(input_path_list, output_path=None, merge=True, encoding='utf-8', delimiter=DEFAULT_DELIMITER, quotechar=DEFAULT_QUOTECHAR, verify_csv_data=True, header=True, file_regex=None, recurse=False, overwrite=False):
    workbook = None
    for input_path in input_path_list:
        if os.path.isdir(input_path):
            workbook = process_directory(input_path, output_path, header=header, workbook=workbook, encoding=encoding, delimiter=delimiter, quotechar=quotechar, verify_csv_data=verify_csv_data, merge=merge, file_regex=file_regex, recurse=recurse, overwrite=overwrite, close_workbook=False)
        else:
            workbook = process_file(input_path, output_path, workbook=workbook, encoding=encoding, delimiter=delimiter, quotechar=quotechar, header=header, verify_csv_data=verify_csv_data, overwrite=overwrite, close_workbook=False)
        if not merge and workbook is not None:
            print('Data written to {}'.format(workbook.filename))
            workbook.close()
            workbook = None

    if workbook is not None:
        print('Data written to {}'.format(workbook.filename))
        workbook.close()

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    #output_group = parser.add_mutually_exclusive_group()
    #output_group.add_argument('-o', '--output', help='Output XLSX filepath (default: {})'.format(DEFAULT_OUTPUT_FILENAME))
    #output_group.add_argument('-a', '--add', help='Add new sheet to an existing  XLSX file')
    parser.add_argument('-o', '--output', help='Output XLSX filepath (default: {})'.format(DEFAULT_OUTPUT_FILENAME))
    parser.add_argument('-d', '--delimiter', default=DEFAULT_DELIMITER, help='CSV delimiter character. Default is  "{}". Set to TAB to use tabulation as delimiter'.format(DEFAULT_DELIMITER))
    parser.add_argument('-q', '--quotechar', default=DEFAULT_QUOTECHAR, help='CSV quotechar character. Default: {} . Set to NONE to disable quotechar'.format(DEFAULT_QUOTECHAR))
    parser.add_argument('-e', '--encoding', default='utf-8', help='File encoding. Default is UTF-8')
    parser.add_argument('--no-header', action='store_true', help='Don\'t process first row as header')
    parser.add_argument('--no-verify', action='store_true', help='Don\'t verify CSV file consistency')
    parser.add_argument('--no-merge', action='store_true', help='Don\'t merge files into a single XLSX file')
    parser.add_argument('-f', '--filter', default=DEFAULT_FILE_REGEX, help='Filename filter with regex. Default is \'{}\')'.format(DEFAULT_FILE_REGEX))
    parser.add_argument('-O', '--overwrite', action='store_true', help='Overwrite existing XLSX output file')
    parser.add_argument('-r', '--recurse', action='store_true', help='Process directories recursively')
    parser.add_argument('input', nargs='+', help='Input CSV file or directory')

    args = parser.parse_args()
    if args.delimiter.strip().lower() == 'tab':
        args.delimiter = '\t'
    if args.quotechar.strip().lower() in ['', 'none']:
        args.quotechar = None
    if not args.no_merge and args.output is None and (len(args.input) > 1 or os.path.isdir(args.input[0])):
        args.output = DEFAULT_OUTPUT_FILENAME
    if args.no_merge and args.output is not None and os.path.basename(args.output) != '':
        print('You have a provided a single file output path and disable merging.\nPlease enable merging or provide a directory as output path')
        sys.exit(1)

    try:
        run(args.input, output_path=args.output, header=(not args.no_header), merge=(not args.no_merge), delimiter=args.delimiter, quotechar=args.quotechar, encoding=args.encoding, file_regex=args.filter, recurse=args.recurse, overwrite=args.overwrite)
    except FileExistsException as e:
        print(e)
        print('Export aborted')
