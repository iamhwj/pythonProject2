#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2023/3/10 15:27
# software: PyCharm
import pandas as pd
import os
import datetime


class keNuo:
    def __init__(self):
        # 获取当前日期
        today = datetime.date.today()
        # 将日期格式化为字符串
        self.today_str = today.strftime('%Y-%m-%d')
        # 读取Excel文件
        self.df = pd.read_excel(r'C:\Users\HeWenjun\Desktop\3.13可诺(1).xlsx', sheet_name=0)
        # 删除第1列和第3列
        self.df.drop(self.df.columns[[0, 2]], axis=1)

    def money_modify(self, row):
        if row['大件件数'] + row['冷链件数'] >= 50 and row['冷链件数'] < 10 and row['实付金额'] < 8000:
            pay = (row['大件件数'] + row['冷链件数']) * 7
            return f'送货费40元,运费（{pay}）元'
        elif row['大件件数'] + row['冷链件数'] >= 50 and row['冷链件数'] < 10 and row['实付金额'] >= 8000:
            return '送货费40元'
        elif row['大件件数'] + row['冷链件数'] < 50 and row['实付金额'] < 8000:
            pay = (row['大件件数'] + row['冷链件数']) * 7
            return f'送货费80元,运费（{pay}）元'
        elif row['大件件数'] + row['冷链件数'] < 50 and row['实付金额'] >= 8000:
            return '送货费80元'
        else:
            return row['备注']

        # 筛选符合条件的数据
        # xizang_df = df.loc[df['省份'] == '西藏自治区']
        # 保存更改后的数据回Excel文件
        # xizang_df.to_excel(f'./{self.today_str}/{self.today_str}专线.xlsx', index=False)

    def data_writer(self):
        df = self.df
        df['备注'] = df.apply(self.money_modify, axis=1)
        df.to_excel('new_example.xlsx', index=False)

    def create_folder(self):
        # 创建文件夹
        folder_path = os.path.join(os.getcwd(), self.today_str)
        os.makedirs(folder_path, exist_ok=True)


if __name__ == '__main__':
    kn = keNuo()
    kn.data_writer()
