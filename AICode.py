#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2023/3/10 15:27
# software: PyCharm
import pandas as pd
import os
import datetime


class KeNuo:
    def __init__(self, file_path):
        # 获取当前日期
        today = datetime.date.today()
        # 将日期格式化为字符串
        self.today_str = today.strftime('%Y-%m-%d')
        # 读取Excel文件
        self.df = pd.read_excel(file_path, sheet_name=0)
        # 删除第1列和第3列
        self.df.drop(self.df.columns[[0, 2]], axis=1)

    def money_modify(self, row):
        if row['大件件数'] + row['冷链件数'] >= 50 and row['冷链件数'] < 10 and row['实付金额'] < 8000:
            pay = (row['大件件数'] + row['冷链件数']) * 7
            return f'送货费40元\n运费（{pay}）元'
        elif row['大件件数'] + row['冷链件数'] >= 50 and row['冷链件数'] < 10 and row['实付金额'] >= 8000:
            return '送货费40元'
        elif row['大件件数'] + row['冷链件数'] < 50 and row['实付金额'] < 8000:
            pay = (row['大件件数'] + row['冷链件数']) * 7
            return f'送货费80元\n运费（{pay}）元'
        elif row['大件件数'] + row['冷链件数'] < 50 and row['实付金额'] >= 8000:
            return '送货费80元'
        else:
            return row['备注']

    def xizang_df(self):
        # 日期
        today = self.today_str
        df = self.df
        # 筛选符合西藏条件的数据
        xizang_df = df.loc[df['省份'] == '西藏自治区']
        # 保存更改后的数据回Excel文件
        xizang_df.to_excel(f'./{today}/{today}专线.xlsx', index=False, sheet_name='西藏')

    def sichuan_df(self):
        # 日期
        today = self.today_str
        df = self.df
        # 筛选符合四川条件的数据
        sichuan_df = df.loc[df['省份'] != '西藏自治区']
        print(sichuan_df)
        # 按城市分类
        groups = sichuan_df.groupby('城市')
        categories = list(groups.groups.keys())
        # for i in categories:
        #     sichuan_city_df = sichuan_df.loc[df['城市'] == i]
        #     # 保存更改后的数据回Excel文件
        #     sichuan_city_df.to_excel(f'./{today}/{today}直送.xlsx', index=False, sheet_name=i)

        with pd.ExcelWriter(f'./{today}/{today}直送.xlsx') as writer:
            for i in categories:
                sichuan_city_df = sichuan_df.loc[df['城市'] == i]
                # 将 df1 写入一个名为 Sheet1 的 sheet 中
                sichuan_city_df.to_excel(writer, index=False, sheet_name=i)

    def create_folder(self):
        # 创建文件夹
        folder_path = os.path.join(os.getcwd(), self.today_str)
        os.makedirs(folder_path, exist_ok=True)

    def data_writer(self):
        df = self.df
        # 按照货物计算价格后，写入备注
        df['备注'] = df.apply(self.money_modify, axis=1)
        self.df = df
        self.create_folder()
        self.xizang_df()
        self.sichuan_df()


if __name__ == '__main__':
    kn = KeNuo(file_path=r'C:\Users\HeWenjun\Desktop\3.13可诺(1).xlsx')
    kn.data_writer()
