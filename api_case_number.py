# -*- coding:utf-8 -*-
import xlrd
import os
from pathlib import Path
import re
import logging

LOG_FORMAT = "%(asctime)s %(levelname)s %(message)s"
DATE_FORMAT = '%Y-%m-%d  %H:%M:%S %a'
logging.basicConfig(level=logging.DEBUG,
                    format=LOG_FORMAT,
                    datefmt = DATE_FORMAT ,
                    filename=r"D:\UP3\DSPi\AI\coding\test.log")

crt_list_colum = {'TestSet':0, 'CaseName':1, 'CasePath':2, 'Priority_19A':3, 'Priority_18B':4,
                  'Duration':5, 'CaseOwner':6, 'SquardTeam':7, 'CRTowner':8, 'Release':9,
                  'TestEntity':10, 'TestEnviroment':11, 'Note':12}

api_common_check = re.compile(r'api_common_chk')

def commoncheck_exist(file_name, check_patten):
    patten_exist_number = 0
    with open(file_name,'r', errors="ignore") as case:
        try:
            lines = case.readlines()
        except UnicodeDecodeError:
            logging.warning("illegal character used in file{}:".format(file_name))
            print("illegal character used in file:", file_name)
            return patten_exist_number
        for i, line in enumerate(lines):
            if line.startswith('#'):
                continue
            if line.startswith('...'):
                continue
            if check_patten.findall(line):
                patten_exist_number += 1
    return patten_exist_number
    #    print(case_ute_name, "has the API common check enabled")
### counter the case number for each test_set
#test_set = {}
#for set_name in sheet1.col_values(0):
#    if set_name in test_set:
#        test_set[set_name] += 1
#    else:
#        test_set[set_name] = 0
#print(test_set)

#all_set_name = sorted(test_set.keys())
#for key, value in enumerate(test_set):
#    print(key, value)

if __name__ == '__main__':
    file_name = "CRT_case_list_UP_20190319.xlsx"
    filePath = os.path.join(os.getcwd(), file_name)
    ute_work_space = Path("C:/robotlte_trunk")


    x1 = xlrd.open_workbook(filePath)
    sheet1 = x1.sheet_by_name("CRT_list")

    total_row = sheet1.nrows
    total_col = sheet1.ncols

# print("row number is ", total_row, "column number is ", total_col)
# print(sheet1.row_values(1, 2, 4))
# case_name = sheet1.row_values(1,2,3)
    test_set = {}
    file_handled = 0
    for row in range(1, total_row - 1):
        if row % 10 == 0:
            print("processed file:", row)
        case_name = sheet1.cell(row, crt_list_colum['CasePath']).value
        if len(case_name) < 4:  ### bypass the deleted cases
            continue
        testset_name = sheet1.cell(row, crt_list_colum['TestSet']).value
        case_priority = sheet1.cell(row, crt_list_colum['Priority_19A']).value
        case_ute_name = ute_work_space/case_name
        if os.path.exists(case_ute_name):
            api_check_exist =  commoncheck_exist(case_ute_name, api_common_check)
            file_handled += 1
            #print("file real name:", os.path.split(case_ute_name)[1])
            #print("file name without path", Path(case_ute_name).parts[-1])
        else:
            print("can't find the file in row",row + 1, case_ute_name)
            logging.info("can't find the file in row {0:d}:{1}".format(row + 1, case_ute_name))
            continue
        if testset_name in test_set:
            test_set[testset_name].append([case_name,case_priority,api_check_exist])
        else:
            test_set[testset_name] = [[case_name,case_priority,api_check_exist]]
    # print("Congratdulations, file process completed, total file is", file_handled)
    for each_set in test_set.keys():
        api_common_enabled_number = 0
        logging.info("================= {} ===============".format(each_set))
        for case in test_set[each_set]:
            if case[2] > 1:
                api_common_enabled_number += 1
            logging.info("  {0}  {1} {2}".format(case[1],case[2], case[0]))
        logging.info("API common check enabled cases in {0} is: {1}\n".format(each_set,api_common_enabled_number))
    #print(test_set)
#print(case_name)
#print(case_ute_name)

