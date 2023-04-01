#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2023/3/10 15:27
# software: PyCharm
import pandas as pd
import os
import datetime
from styleModify import style_modify


class KeNuo:
    # 格式化日期，作为文档名称，并首次处理Excel源文件内容
    def __init__(self, file_path):
        # 获取当前日期
        today = datetime.date.today()
        # 将日期格式化为字符串
        self.today_str = today.strftime('%Y-%m-%d')
        # 读取Excel文件,第一张sheet表
        self.df = pd.read_excel(file_path, sheet_name=0)
        # 删除第1列和第3列
        self.df = self.df.drop(self.df.columns[[0, 2]], axis=1)
        # print(self.df['货收联系人'])

    # 修改联系人样式
    def set_member_style(self, row):
        if row in ['郑楠', '王发成', '岳依萱', '杨腾军', '梁卜月']:
            return 'background-color: #EF7D7D'

    # 修改金额样式
    def set_money_style(self, row):
        if row < 8000:
            return 'background-color: #C6EFCE'

    # 按照配送件数计算金额，写入备注列
    def money_modify(self, row):
        # 没有冷链的判断
        if row['冷链件数'] == 0 or pd.isnull(row['冷链件数']):
            if row['大件件数'] >= 50 and row['实付金额'] >= 8000:
                return row['备注']
            elif row['大件件数'] >= 50 and row['实付金额'] < 8000:
                return f"运费{7 * row['大件件数']}"
            elif row['大件件数'] < 50 and row['实付金额'] >= 8000:
                return '送货费80'
            elif row['大件件数'] < 50 and row['实付金额'] < 8000:
                return f"送货费80\n运费{7 * row['大件件数']}"
        # 有冷链的判断
        else:
            if row['大件件数'] + row['冷链件数'] >= 50:
                if row['冷链件数'] <= 10 and row['实付金额'] >= 8000:
                    return '送货费40'
                elif row['冷链件数'] <= 10 and row['实付金额'] < 8000:
                    return f"送货费40\n运费{7 * row['大件件数']}"
                elif row['冷链件数'] > 10 and row['实付金额'] >= 8000:
                    return row['备注']
                elif row['冷链件数'] > 10 and row['实付金额'] < 8000:
                    return f"运费{7 * row['大件件数']}"
            else:
                if row['实付金额'] >= 8000:
                    return f"送货费80"
                elif row['实付金额'] < 8000:
                    return f"送货费80\n运费{7 * row['大件件数']}"

    # 写入西藏的数据，并将四川城市创建为新的空sheet
    def xizang_df(self):
        # 日期
        today = self.today_str
        filename = f'./{today}/{today}专线.xlsx'
        df = self.df
        # 筛选符合西藏条件的数据
        xizang_df = df.loc[df['省份'] == '西藏自治区']
        # 新增几个空的表，以四川城市命名
        sichuan_df = df.loc[df['省份'] != '西藏自治区']
        # 按四川城市分类
        groups = sichuan_df.groupby('城市')
        categories = list(groups.groups.keys())
        # 准备空表单
        null_df = pd.DataFrame({'': ['']})
        with pd.ExcelWriter(filename) as writer:
            xizang_df.style.applymap(self.set_member_style, subset='实付金额')
            xizang_df.to_excel(writer, index=False, sheet_name='西藏自治区')
            for i in categories:
                # 按城市名添加新的空表
                null_df.to_excel(writer, sheet_name=i, index=False)
        style_modify(filename)

    # 将四川的城市数据分类写入
    def sichuan_df(self):
        # 日期
        today = self.today_str
        filename = f'./{today}/{today}直送.xlsx'
        df = self.df
        # 筛选符合四川条件的数据
        sichuan_df = df.loc[df['省份'] != '西藏自治区']
        # 按城市分类
        groups = sichuan_df.groupby('城市')
        categories = list(groups.groups.keys())
        with pd.ExcelWriter(filename) as writer:
            for i in categories:
                sichuan_city_df = sichuan_df.loc[df['城市'] == i]
                # 链式添加背景色
                sichuan_city_df = sichuan_city_df.style.applymap(self.set_member_style, subset='货收联系人').applymap(
                    self.set_money_style, subset='实付金额')
                # 将城市数据写入每个对应名的Sheet中
                sichuan_city_df.to_excel(writer, index=False, sheet_name=i)
        style_modify(filename)

    def create_folder(self):
        # 创建文件夹
        folder_path = os.path.join(os.getcwd(), self.today_str)
        os.makedirs(folder_path, exist_ok=True)

    def data_writer(self):
        # 初处理过的df
        df = self.df
        # 按照货物计算价格后，写入备注
        df['备注'] = df.apply(self.money_modify, axis=1)
        # 2次处理的df赋值
        self.df = df
        self.create_folder()
        self.xizang_df()
        self.sichuan_df()


if __name__ == '__main__':
    kn = KeNuo(file_path=r"C:\Users\17647\Desktop\3.13可诺(1).xlsx")
    kn.data_writer()
