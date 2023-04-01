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
# import pandas as pd
# # 创建两个示例 DataFrame
# df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
# df2 = pd.DataFrame({'X': [7, 8, 9], 'Y': [10, 11, 12]})
# # 创建一个 ExcelWriter 对象，指定文件名
# with pd.ExcelWriter('example.xlsx') as writer:
#     # 将 df1 写入一个名为 Sheet1 的 sheet 中
#     df1.to_excel(writer, sheet_name='Sheet1')
#     # 将 df2 写入一个名为 Sheet2 的 sheet 中
#     df2.to_excel(writer, sheet_name='Sheet2')

# class ZhiPi:
#     def __init__(self, file_path):
#         # 获取当前日期
#         today = datetime.date.today()
#         # 将日期格式化为字符串
#         self.today_str = today.strftime('%Y-%m-%d')
#         # 读取Excel文件
#         self.df = pd.read_excel(file_path, sheet_name=0)
#         # 删除第1列和第3列
#         self.df = self.df.drop(self.df.columns[[0, 2]], axis=1)
#         print(self.df)
#
#
# ZhiPi(file_path=r'C:\Users\HeWenjun\Desktop\3.13可诺(1).xlsx')
# num1 = int(input('input1:'))
# num2 = int(input('input2:'))
# if num1 < 10 and num2 < 10:
#     print('small')
# elif num1 < 10 < num2:
#     print('medium')
# elif num1 >= 10 and num2 >= 10:
#     print('big')
# import pandas as pd
# import numpy as np
# import random

# # 创建一个包含随机数据的DataFrame
# df1 = pd.DataFrame(np.random.randint(0, 100, size=(10, 4)), columns=list('ABCD'))
# # 筛选出df1中B列的值大于50的行，并将结果保存到df2中
# df2 = df1[df1['B'] > 50]
#
#
# # 创建一个函数，用于生成样式
# def highlight(value):
#     if value > 50:
#         return 'background-color: yellow'
#     else:
#         return ''
#
#
# # 使用style.applymap()方法将样式应用到df2中
# styled_df = df2.style.applymap(highlight)
# styled_df.to_excel('style.xlsx')
i = 10
if i in [1, 11]:
    print('in it')
