#!/usr/bin/python
# -*- coding: UTF-8 -*-
import getopt
import json
import math
import os
import sys

import pandas as pd


def read_xlsx(file_name, header=0):
    df = pd.read_excel(file_name, sheet_name=None, header=header)
    result = {}
    for i in df.keys():
        result[i] = df[i].to_dict(orient='records')
    return result


def is_excel(file_name):
    ext = os.path.splitext(file_name)[1]
    return ext == '.xlsx' or ext == '.xls'


def list_xlsx(file_dir):
    result = []
    if os.path.isdir(file_dir):
        for root, dirs, files in os.walk(file_dir):
            for file in files:
                if is_excel(file):
                    result.append(os.path.join(root, file))
    elif is_excel(file_dir):
        result.append(file_dir)
    return result


def clean_empty(d):
    if not isinstance(d, (dict, list)):
        return d
    if isinstance(d, list):
        return [v for v in (clean_empty(v) for v in d) if v]
    return {k: v for k, v in ((k, clean_empty(v)) for k, v in d.items()) if
            v and (type(v) != float or not math.isnan(v))}


def execute(input_path, output_path=None, header=1, compress=False):
    output_path = output_path if output_path is not None else input_path
    result = {}
    excels = list_xlsx(input_path)
    for excel in excels:
        print(excel)
        sheets = read_xlsx(excel, header)
        for name in sheets.keys():
            print('\t', name)
            result[name] = sheets[name]
    result = clean_empty(result)
    indent = None if compress else 2
    to_json = json.dumps(result, ensure_ascii=False, indent=indent)
    output_path = output_path if os.path.isfile(output_path) else output_path + '/metadata.json'
    with open(output_path, 'wb') as out:
        out.write(bytes(to_json, encoding="utf-8"))


def main(argv):
    input_path = None
    output_path = None
    header = 1
    compress = False
    show_help = '[*] -i <input_path> -o <output_path> -n <name> -c <compress>'
    try:
        opts, args = getopt.getopt(argv, "hi:o:n:c:", ["input=", "output=", "name=", "compress=", "help="])
    except getopt.GetoptError:
        print(show_help)
        sys.exit(2)
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(show_help)
            sys.exit()
        elif opt in ("-i", "--input"):
            input_path = arg
        elif opt in ("-o", "--output"):
            output_path = arg
        elif opt in ("-n", "--name"):
            header = int(arg)
        elif opt in ("-c", "--compress"):
            compress = bool(arg)
    try:
        execute(input_path, output_path=output_path, header=header, compress=compress)
    except TypeError:
        print(show_help)


if __name__ == '__main__':
    main(sys.argv[1:])
