#!/usr/bin/python3

import json
import xlrd
import os


def other_sheet(sheet, idIndex):
    fields = sheet.row_values(1)

    count = 2
    list = []
    while count < sheet.col_values(0).__len__() - 2:
        dic = {}
        flag = 0
        row_data = sheet.row_values(count)
        for i in range(sheet.ncols):  # 列数
            d = row_data[i]
            if int(d) == idIndex:
                dic[fields[i]] = int(d)
                flag = flag + 1
            elif flag > 0:
                if isinstance(d, (int, float)):
                    if d == int(d):
                        dic[fields[i]] = int(d)
                    else:
                        dic[fields[i]] = d
                else:
                    dic[fields[i]] = d
                list.append(dic)
        count = count + 1
    return list


def begin():
    print("=================================")
    table = input("请输入表名：")
    path = os.path.abspath('.') + "\\" + table + ".xlsx"

    data = xlrd.open_workbook(path)
    sheet = data.sheet_by_index(0)

    name = sheet.row_values(0)[1]
    fields = sheet.row_values(1)

    count = 2
    list = []
    while count <= sheet.col_values(0).__len__()-1:  # 有多少行
        dic = {}
        row_data = sheet.row_values(count)  # 每一行的数据
        for i in range(sheet.ncols):
            d = row_data[i]
            if "index" in str(d):
                lindex = d.find("=")
                rindex = d.rfind("=")
                sheetIndex = d[lindex + 1:lindex + 2]
                sheet2 = data.sheet_by_index(int(sheetIndex) - 1)
                idIndex = d[rindex + 1:len(d)]
                dic[fields[i]] = other_sheet(sheet2, int(idIndex))
            else:
                if isinstance(d, (int, float)):
                    if d == int(d):
                        dic[fields[i]] = int(d)
                    else:
                        dic[fields[i]] = d
                else:
                    dic[fields[i]] = d
        list.append(dic)
        count = count + 1

    j = json.dumps(list, sort_keys=True, indent=4, ensure_ascii=False)

    document = open("..//" + name + ".json", "w+", encoding='gb18030')
    document.write(str(j))
    document.close()

    print("文件导出成功：" + name + ".json")
    print()


def begin2():
    print("=================================")
    table = input("请输入表名：")
    path = os.path.abspath('.') + "\\" + table + ".xlsx"

    data = xlrd.open_workbook(path)
    sheet = data.sheet_by_index(0)

    name = sheet.row_values(0)[1]
    fields = sheet.row_values(1)

    count = 2
    list = []
    while count < sheet.col_values(0).__len__() - 2:  # 有多少行
        dic = {}
        row_data = sheet.row_values(count)  # 每一行的数据
        playerId = row_data[0]
        gold = row_data[1]
        dic[int(playerId)] = int(gold)
        list.append(dic)
        count = count + 1

    document = open("./" + name + ".txt", "w+")
    document.write(str(list))
    document.close()

    print("文件导出成功：" + name + ".json")
    print()


if __name__ == '__main__':
    while True:
        begin2()
