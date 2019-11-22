"""
AUTHOR: GODWIN KWAN
DATE: 2019-11
# UNICODE: UTF-8
"""


import os
import pandas


def id_match(files_path: str):
    extract_files = filtered_files(files_path)
    print('Files Found: ', len(extract_files))

    extract_files = list_sort_keyword(extract_files, '芒哥')
    print(extract_files)

    # 匹配零售数据
    data_database = pandas.DataFrame(pandas.read_excel(os.path.join(files_path, extract_files[0]),
                                                       sheet_name='芒哥零售数据',
                                                       usecols=['ID', '拜访日期', '拜访时间', '药店名称']))
    data_database_date = data_database['拜访日期'].dt.strftime('%Y-%m-%d')
    data_database_time = pandas.to_datetime(data_database['拜访时间'].astype(str)).dt.strftime('%H:%M')
    data_database['MatchID'] = data_database_date + '_' + data_database_time + '_' + data_database['药店名称'].str.strip()
    data_database.drop(['拜访日期', '拜访时间', '药店名称'], axis=1, inplace=True)

    for i in range(1, len(extract_files)):
        data_unit = pandas.DataFrame(pandas.read_excel(os.path.join(files_path, extract_files[i]),
                                                       sheet_name='拜访服务', skiprows=2, usecols=list(range(18))))

        data_unit.columns = [
            '序号', '服务商属性', '服务商名称', '服务人员', '服务方式', '服务对象/类型', '拜访日期', '终端名称', '省',
            '终端地址', '签到时间', '拜访时长', '被拜访人姓名', '科室', '沟通产品', '拜访小结', '服务编码', '服务记录平台'
        ]  # 暴力替换表头

        data_unit_date = data_unit['拜访日期'].dt.strftime('%Y-%m-%d')
        data_unit_time = pandas.to_datetime(data_unit['签到时间'].astype(str)).dt.strftime('%H:%M')
        data_unit['MatchID'] = data_unit_date + '_' + data_unit_time + '_' + data_unit['终端名称'].str.strip()

        data_unit = pandas.merge(data_unit, data_database, on='MatchID', how='left')

        data_unit.drop(['服务编码', 'MatchID'], axis=1, inplace=True)  # 只删除列就用`.drop()`，如果要移动列就用`.pop()`再配合`insert()`
        matched_col = data_unit.pop('ID')
        data_unit.insert(16, '服务编码', matched_col)

        print(data_unit)

    # 匹配药店活动
    data_database = pandas.DataFrame(pandas.read_excel(os.path.join(files_path, extract_files[0]),
                                                       sheet_name='药店活动',
                                                       usecols=['ID', '主题', '开始时间']))
    data_database['主题'] = data_database['主题'].str.replace('广州市', '广州')  # 删除特定字符

    data_database['MatchID'] = data_database['开始时间'].dt.strftime('%Y-%m-%d') + '_' + data_database['主题'].str.strip()
    data_database.drop(['主题', '开始时间'], axis=1, inplace=True)

    for i in range(1, len(extract_files)):
        data_unit = pandas.DataFrame(pandas.read_excel(os.path.join(files_path, extract_files[i]),
                                                       sheet_name='店员培训服务', skiprows=2, usecols=list(range(17))))

        data_unit.columns = [
            '序号', '服务商属性', '服务商名称', '服务人员', '服务对象/类型', '服务方式', '培训时间', '省', '培训地点',
            '签到时间', '受训单位', '受训人数', '培训主题', '培训内容', '培训小结', '服务编码', '服务记录平台'
        ]  # 暴力替换表头
        data_unit['受训单位'] = data_unit['受训单位'].str.replace('广州市', '广州')  # 删除特定字符

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
                print(regex_pattern)
                if re.search(regex_pattern, data_database['MatchID'].iloc[k]):
                    data_unit.iloc[j, 18] = data_database.iloc[k, 1]  # 也可以直接把需要匹配的值直接赋给目标列上，就不用了下面的`merge()`操作了

        data_unit = pandas.merge(data_unit, data_database, on='MatchID', how='left')
        print(data_unit)
