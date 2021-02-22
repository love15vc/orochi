#!/usr/bin/python
# -*- coding: UTF-8 -*-
import json
import sys
import math
import pandas as pd


def read_xlsx(file_name, header=0):
    df = pd.read_excel(file_name, header=header)
    result = []
    for i in df.keys():
        result.append(df[i].to_dict())
    return result[0], result[1:]


def gen_json(keys, values):
    i18n = values.pop(0)
    if type(i18n) == float:
        return
    result = {}
    for i in values:
        if keys[i] and not str(keys[i]).isspace() and (type(keys[i]) != float or not math.isnan(keys[i])):
            result[keys[i]] = values[i] if type(values[i]) != float or not math.isnan(values[i]) else keys[i]
    output = 'translation.' + i18n + '.json'
    to_json = json.dumps(result, ensure_ascii=False, indent=2)
    with open(output, 'wb') as out:
        out.write(bytes(to_json, encoding="utf-8"))
    print(output)


def main(file_name):
    keys, values_all = read_xlsx(file_name)
    for values in values_all:
        gen_json(keys, values)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1])
