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

def process_csv(row_list, output_path=DEFAULT_OUTPUT_FILENAME, workbook=None, title=None, auto_size=True, filter_first_row=True, freeze_first_row=True, header=True, close_workbook=True, force=False):
    if workbook is None:
        if os.path.exists(output_path):
            if not force:
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

    if close_workbook:
        workbook.close()
        return workbook.filename
    else:
        return workbook

def process_file(input_path, output_path=None, workbook=None, title=None, encoding='utf-8', header=True, delimiter='\t', quotechar=None, force=False):
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


    with open(input_path, 'r', encoding=encoding, errors='ignore') as csvfile:
        csv_reader = csv.reader(csvfile, delimiter=delimiter, quotechar=quotechar)
        row_list = list(csv_reader)

    workbook = process_csv(row_list, output_path=output_path, workbook=workbook, title=title, header=header, close_workbook=False, force=force)
    print('Imported : {}'.format(input_path))
    return workbook

def process_directory(input_path, output_path=None, encoding='utf-8', header=True, delimiter='\t', quotechar=None, merge=True, file_regex=None, workbook=None, recurse=False, force=False):
    for entry in os.listdir(input_path):
        path = os.path.join(input_path, entry)
        if os.path.isdir(path) and recurse:
            workbook = process_directory(path, output_path, encoding=encoding, header=header, delimiter=delimiter, quotechar=quotechar, merge=merge, file_regex=file_regex, workbook=workbook, recurse=recurse)
            if workbook is not None and not merge:
                print('Data written to {}'.format(workbook.filename))
                workbook.close()
                workbook=None
        elif os.path.isfile(path) and (file_regex is None or file_regex.strip() == '' or re.match(file_regex, entry)):
            workbook = process_file(path, output_path, encoding=encoding, workbook=workbook, delimiter=delimiter, quotechar=quotechar, header=header, force=force)
            if workbook is not None and not merge:
                print('Data written to {}'.format(workbook.filename))
                workbook.close()
                workbook=None
    return workbook

def run(input_path_list, output_path=None, merge=True, encoding='utf-8', delimiter='\t', quotechar=None, header=True, file_regex=None, recurse=False, force=False):
    workbook = None
    for input_path in input_path_list:
        if os.path.isdir(input_path):
            workbook = process_directory(input_path, output_path, header=header, workbook=workbook, encoding=encoding, delimiter=delimiter, quotechar=quotechar, merge=merge, file_regex=file_regex, recurse=recurse, force=force)
        else:
            workbook = process_file(input_path, output_path, workbook=workbook, encoding=encoding, delimiter=delimiter, quotechar=quotechar, header=header, force=force)
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
    parser.add_argument('-d', '--delimiter', default='\t', help='CSV delimiter character (default: TAB)')
    parser.add_argument('-q', '--quotechar', default=None, help='CSV quotechar character (default: None)')
    parser.add_argument('-e', '--encoding', default='utf-8', help='File encoding (default: UTF-8)')
    parser.add_argument('--no-header', action='store_true', help='Don\'t process first row as header')
    parser.add_argument('--no-merge', action='store_true', help='Don\'t merge files into a single XLSX file')
    parser.add_argument('-f', '--filter', default='.*\\.[C^c][S^s][V^v]$', help='Filename filter with regex (default: \'.*\\.[C^c][S^s][V^v]$\')')
    parser.add_argument('-F', '--force', action='store_true', help='Overwrite existing file')
    parser.add_argument('-r', '--recurse', action='store_true', help='Process directories recursively')
    parser.add_argument('input', nargs='+', help='Input CSV file or directory')

    args = parser.parse_args()
    if not args.no_merge and args.output is None and (len(args.input) > 1 or os.path.isdir(args.input[0])):
        args.output = DEFAULT_OUTPUT_FILENAME
    if args.no_merge and args.output is not None and os.path.basename(args.output) != '':
        print('You have a provided a single file output path and disable merging.\nPlease enable merging or provide a directory as output path')
        sys.exit(1)

    try:
        run(args.input, output_path=args.output, header=(not args.no_header), merge=(not args.no_merge), delimiter=args.delimiter, quotechar=args.quotechar, encoding=args.encoding, file_regex=args.filter, recurse=args.recurse, force=args.force)
    except FileExistsException as e:
        print(e)
        print('Export aborted')
