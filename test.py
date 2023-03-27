#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2023/3/24 17:32
# software: PyCharm
# import pandas as pd
#
# # 读取Excel表格数据
# df = pd.read_excel('example.xlsx')
#
#
# # 定义函数，根据条件修改第三列的值
# def modify(row):
#     if row['大件件数'] + row['冷链件数'] >= 50 and row['小件件数']<10:
#         return '大于10'
#     else:
#         return row['备注']
#
#
# # 对每一行数据应用函数
# df['Column3'] = df.apply(modify, axis=1)
# # 保存修改后的数据到Excel表格中
# # df.to_excel('new_example.xlsx', index=False)
# import pandas as pd
# # 创建DataFrame
# data = {'姓名': ['张三', '李四'], '成绩': ['90\n95', '80\n85']}
# df = pd.DataFrame(data)
# # 将DataFrame写入Excel文件
# df.to_excel('example.xlsx', index=False)
import pandas as pd
# 创建两个示例 DataFrame
df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
df2 = pd.DataFrame({'X': [7, 8, 9], 'Y': [10, 11, 12]})
# 创建一个 ExcelWriter 对象，指定文件名
with pd.ExcelWriter('example.xlsx') as writer:
    # 将 df1 写入一个名为 Sheet1 的 sheet 中
    df1.to_excel(writer, sheet_name='Sheet1')
    # 将 df2 写入一个名为 Sheet2 的 sheet 中
    df2.to_excel(writer, sheet_name='Sheet2')