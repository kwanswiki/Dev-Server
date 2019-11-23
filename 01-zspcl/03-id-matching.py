"""
-*- AUTHOR: GODWIN KWAN -*-
-*- CREATED: 2019-11 -*-
-*- CODING: UTF-8 -*-
"""


import os
import pandas
import openpyxl
import re


def filtered_files(dir_path: str):

    all_files = os.listdir(dir_path)
    file_list = os.listdir(dir_path)

    for i in all_files:
        if (i.endswith('.xls')) or (i.endswith('.xlsx')):
            print(i)
        else:
            file_list.remove(i)
    return file_list


def list_sort_keyword(item_list, sort_keyword: str):

    for i in range(0, len(item_list)):
        if sort_keyword in item_list[i]:
            key_item = item_list.pop(i)
            item_list.insert(0, key_item)
            break
    return item_list


def sheet_operation1(data_unit, data_database):

    data_unit.columns = [
        '序号', '服务商属性', '服务商名称', '服务人员', '服务方式', '服务对象/类型', '拜访日期', '终端名称', '省',
        '终端地址', '签到时间', '拜访时长', '被拜访人姓名', '科室', '沟通产品', '拜访小结', '服务编码', '服务记录平台'
    ]

    data_unit_date = data_unit['拜访日期'].dt.strftime('%Y-%m-%d')
    data_unit_time = pandas.to_datetime(data_unit['签到时间'].astype(str)).dt.strftime('%H:%M')
    data_unit['MatchID'] = data_unit_date + '_' + data_unit_time + '_' + data_unit['终端名称'].str.strip()

    data_unit = pandas.merge(data_unit, data_database, on='MatchID', how='left')

    data_unit.drop(['服务编码', 'MatchID'], axis=1, inplace=True)
    matched_col = data_unit.pop('ID')
    data_unit.insert(16, '服务编码', matched_col)

    return data_unit


def sheet_operation2(data_unit, data_database):

    data_unit.columns = [
        '序号', '服务商属性', '服务商名称', '服务人员', '服务对象/类型', '服务方式', '培训时间', '省', '培训地点',
        '签到时间', '受训单位', '受训人数', '培训主题', '培训内容', '培训小结', '服务编码', '服务记录平台'
    ]
    data_unit['受训单位'] = data_unit['受训单位'].str.replace('广州市', '广州')

    data_unit['MatchID_1'] = data_unit['培训时间'].dt.strftime('%Y-%m-%d')
    data_unit['MatchID'] = None

    for j in range(0, data_unit.shape[0] - 1):
        for k in range(0, data_database.shape[0]):
            '''
            正则匹配：
            - `.`是匹配除换行符`\n`外的任意字符，`*`表示匹配前一个字符0次或无限次，+或*后跟？表示非贪婪匹配，即尽可能少的匹配，如*?重复任意次，但尽可能少重复；
            - 具体参考[该文章](https://blog.csdn.net/qq_37699336/article/details/84981687)
            '''
            regex_pattern = data_unit['MatchID_1'].iloc[j] + '_.*?' + str(data_unit['受训单位'].iloc[j]).strip() + '.*?'
            if re.search(regex_pattern, data_database['MatchID'].iloc[k]):
                data_unit.iloc[j, 18] = data_database.iloc[k, 1]

    data_unit = pandas.merge(data_unit, data_database, on='MatchID', how='left')

    data_unit.drop(['服务编码', 'MatchID_1', 'MatchID'], axis=1, inplace=True)
    matched_col = data_unit.pop('ID')
    data_unit.insert(15, '服务编码', matched_col)

    return data_unit


def id_match(files_path: str):
    extract_files = filtered_files(files_path)
    extract_files = list_sort_keyword(extract_files, '芒哥')
    print('Files Found: ', len(extract_files))

    data_database1 = pandas.DataFrame(pandas.read_excel(os.path.join(files_path, extract_files[0]),
                                                        sheet_name='芒哥零售数据',
                                                        usecols=['ID', '拜访日期', '拜访时间', '药店名称']))
    data_database1_date = data_database1['拜访日期'].dt.strftime('%Y-%m-%d')
    data_database1_time = pandas.to_datetime(data_database1['拜访时间'].astype(str)).dt.strftime('%H:%M')
    data_database1['MatchID'] = data_database1_date + '_' + data_database1_time + '_' + data_database1['药店名称'].str.strip()
    data_database1.drop(['拜访日期', '拜访时间', '药店名称'], axis=1, inplace=True)

    data_database2 = pandas.DataFrame(pandas.read_excel(os.path.join(files_path, extract_files[0]),
                                                        sheet_name='药店活动',
                                                        usecols=['ID', '主题', '开始时间']))
    data_database2['主题'] = data_database2['主题'].str.replace('广州市', '广州')

    data_database2['MatchID'] = data_database2['开始时间'].dt.strftime('%Y-%m-%d') + '_' + data_database2['主题'].str.strip()
    data_database2.drop(['主题', '开始时间'], axis=1, inplace=True)

    for i in range(1, len(extract_files)):

        sheet_name1 = '拜访服务'
        data_unit1 = pandas.DataFrame(pandas.read_excel(os.path.join(files_path, extract_files[i]),
                                                        sheet_name=sheet_name1, skiprows=2, usecols=list(range(18))))
        data_unit1 = sheet_operation1(data_unit1, data_database1)

        sheet_name2 = '店员培训服务'
        data_unit2 = pandas.DataFrame(pandas.read_excel(os.path.join(files_path, extract_files[i]),
                                                        sheet_name=sheet_name2, skiprows=2, usecols=list(range(17))))
        data_unit2 = sheet_operation2(data_unit2, data_database2)

        excel_writer = pandas.ExcelWriter(os.path.join(files_path, extract_files[i]), engine='openpyxl')
        excel_book = openpyxl.load_workbook(excel_writer.path)
        excel_writer.book = excel_book
        excel_writer.sheets = dict((worksheet.title, worksheet) for worksheet in excel_book.worksheets)
        data_unit1.to_excel(excel_writer, sheet_name=sheet_name1, encoding='utf-8', index=False, header=True, startrow=2)
        data_unit2.to_excel(excel_writer, sheet_name=sheet_name2, encoding='utf-8', index=False, header=True, startrow=2)
        print('已完成匹配文档 ', extract_files[i])
        excel_writer.save()
        excel_writer.close()
