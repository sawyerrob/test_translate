#!/usr/bin/python
# -*- coding: UTF-8 -*-
import xlwt
import pandas as pd
import random
import winreg
import os


def read_excel(excel_path,excel_list):
    #默认情况下，pandas 假定第一行为表头 (header)，如果 Excel 不是从第一行开始，
    # header 参数用于指定将哪一行作为表头，表头在 DataFrame 中变成列索引 (column index) ，
    # header 参数从 0 开始，比如第二行作为 header，则header = 0 ,如果没有列名请用None
    data_frame1 = pd.read_excel(excel_path,header = None)
    # pd很好用，但是需要注意读取excel后还需用values属性才能得到值
    #获取具体数据
    datas1 = data_frame1.values
    # 二维数组变一维
    for data1 in datas1:
        for each in data1:
            excel_list.append(each)
    return excel_list



# 从注册表中获得桌面路径
def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                         r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]


# 随机生成一个乱序数组
def random_list(start, stop, length):
    if length >= 0:
        length = int(length)
        start, stop = (int(start), int(stop)) if start <= stop else (int(stop), int(start))
        random_list = []
        for i in range(length):
            random_list.append(random.randint(start, stop))
        return random_list


# 创建一个style模型
def set_style(name, height, bold=False, format_str='', align='center', color=44):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.height = height

    borders = xlwt.Borders()  # 为样式创建边框
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2

    alignment = xlwt.Alignment()  # 设置排列
    if align == 'center':
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER
    else:
        alignment.horz = xlwt.Alignment.HORZ_LEFT
        alignment.vert = xlwt.Alignment.VERT_BOTTOM

    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # 默认为55室灰色，44是绿色
    pattern.pattern_fore_colour = color

    style.font = font
    style.borders = borders
    style.num_format_str = format_str
    style.alignment = alignment
    style.pattern = pattern

    return style


def write_excel(path,list_char,list_num):
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('Worksheet', cell_overwrite_ok=True)
    worksheet.write_merge(0, 0, 0, 9, '欢迎测试', set_style(u"微软雅黑", 700, True))
    # 建立一个画线框,设置行高和列宽
    for col in range(10):
        for row in range(20):
            worksheet.write(row + 1, col, "", set_style(u"微软雅黑", 300, True))
        worksheet.col(col).width = 256 * 12
    # 填充数据
    for col in range(10):
        for row in range(20):
            if col % 2 == 0:
                if row % 2 == 0:
                    worksheet.write(row + 1, col, str(random.choice(list_char)), set_style(u"微软雅黑", 300, True))
            elif row % 2 == 0:
                # 他爸的灯壳子，不转化为字符串会异常raise Exception("Unexpected data type %r" % type(label))
                worksheet.write(row + 1, col, str(random.choice(list_num)), set_style(u"微软雅黑", 300, True))
                # worksheet.write(row + 1, col, "", set_style(u"微软雅黑", 200, True))

    workbook.save(path + '\\test.xls')


if __name__ == "__main__":
    # 汉字请命名为1.xlsx和程序在同一个目录下
    excel1 = os.getcwd() + '\\1.xlsx'
    # 数字请命名为2.xlsx和程序在同一个目录下
    excel2 = os.getcwd() + '\\2.xlsx'
    excel_data1 = []
    excel_data2 = []
    list_num =read_excel(excel1,excel_data1)
    list_char = read_excel(excel2, excel_data2)
    path = get_desktop()
    write_excel(path,list_char,list_num)
    # print(random.choice(list_char))

