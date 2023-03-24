#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2023/3/24 17:32
# software: PyCharm
import pandas as pd

# 读取Excel表格数据
df = pd.read_excel('example.xlsx')


# 定义函数，根据条件修改第三列的值
def modify(row):
    if row['大件件数'] + row['冷链件数'] >= 50 and row['小件件数']<10:
        return '大于10'
    else:
        return row['备注']


# 对每一行数据应用函数
df['Column3'] = df.apply(modify, axis=1)
# 保存修改后的数据到Excel表格中
df.to_excel('new_example.xlsx', index=False)
