# -*- coding: utf-8 -*-
"""
Created on Sun Mar 16 01:49:52 2025

@author: thinkpad
"""

#import os
#from exportFigure1 import ex_Fig
import re
#import tkinter as tk
#import tkinter.ttk as ttk
#from tkinter import messagebox, filedialog

def spilt_name(name):
    numbers = re.findall(r'\d+\.?\d*', name) # 使用正则表达式查找字符串中的数字
    numbers = [float(num) if '.' in num else int(num) for num in numbers]
    return numbers

def sorted_files(filenames, LV_normal, key_word):
    # 创建一个空列表来存储文件名
    file_names_other = []
    file_names_1500V = []  #存储额定电压的文件名
#    PIC_series = 1  #初始图片序号 
    # 遍历文件夹中的所有文件
    for filename in filenames:
        # 检查文件名是否包含"明细"
        if key_word in filename:
            # 将文件名添加到列表中
            if LV_normal in filename:
                file_names_1500V.append(filename)
            else:
                file_names_other.append(filename)
    
#    print(file_names_other)
#    file_names_1500V = [s for s in file_names if LV_normal in s]
    file_names_1500V = sorted(file_names_1500V, key=lambda x: (spilt_name(x)[1], spilt_name(x)[3]), reverse = True)
    
#    file_names_1000V = [s for s in file_names if LV_min in s]
    file_names_other = sorted(file_names_other, key=lambda x: (spilt_name(x)[0], -spilt_name(x)[1], -spilt_name(x)[3]))
    
#    file_names_1800V = [s for s in file_names if LV_max in s]
#    file_names_1800V = sorted(file_names_1800V, key=lambda x: (spilt_name(x)[1], spilt_name(x)[3]), reverse = True)
    
    file_names = file_names_1500V + file_names_other
    return file_names
# 指定文件夹路径
def chooseAll(folder_path, stationFile_path, save_path, file_name, LV_normal, PIC_series, xaxis_flag, yaxis_names, plot_lims, scale_lists, color_lists, plot_names):
    PIC_series = ex_Fig(file_name, folder_path, stationFile_path, save_path, PIC_series, xaxis_flag, yaxis_names, plot_lims, scale_lists, color_lists, plot_names)




if __name__ == '__main__':
    aa = 'C:\\Users\\thinkpad\\Desktop\\20250320\\AW0-无惰行\\'  #打开文件目录
    bb = 'C:\\Users\\thinkpad\\Desktop\\线路仿真绘图\\设施.xls'  #站点文件路径
    cc = 'C:\\Users\\thinkpad\\Desktop\\20250320\\AW0-无惰行\\' #图片保存目录
    dd = '1500V'
    ee = 1
#    ee = {'group0' : 0,
#          'group1' : 0,
#          'group2' : 1,
#          'group3' : 1}  #x轴选择，0——运行里程  1——运行时间
#    ff = {
#            'group0' : ['速度(km/h)', '时间(s)', '网侧电流(A)', '电机电流(A)'], 
#            'group1' : ['累计能耗(kWh)', '网侧输入功率(kW)', '电机输出功率(kW)'],
#            'group2' : ['速度(km/h)', '网侧电流(A)', '电机电流(A)'],
#            'group3' : ['累计能耗(kWh)', '网侧输入功率(kW)', '电机输出功率(kW)']            
#         }  #绘图选择
#    gg = {
#            'group0' : [[0, 100], [0, 300], [-5000, 5000], [-500, 500]], 
#            'group1' : [[0, 500], [-8000, 8000], [-800, 800]],
#            'group2' : [[0, 100], [-5000, 5000], [-500, 500]],
#            'group3' : [[0, 500], [-8000, 8000], [-800, 800]]
#         }  #范围
#    hh = {
#            'group0' : [10, 10, 10, 10], 
#            'group1' : [10, 10, 10],
#            'group2' : [10, 10, 10],
#            'group3' : [10, 10, 10]
#         }  #格数
#    ii = {
#            'group0' : ['red', 'blue', 'green', 'cyan'], 
#            'group1' : ['magenta', 'yellow', 'black'],
#            'group2' : ['red', 'green', 'cyan'],
#            'group3' : ['magenta', 'yellow', 'black']
#         }  #颜色
#    jj = {'group0' : "运行时间、速度、网侧电流、电机电流对运行里程曲线", 
#          'group1' : "累计能耗、网侧输入功率、电机输出功率对运行里程曲线",
#          'group2' : "速度、网侧电流、电机电流对运行时间曲线",
#          'group3' : "累计能耗、网侧输入功率、电机输出功率对运行时间曲线"
#         }  #图名
    ff = {'group0': 0}
    gg = {'group0': ['速度(km/h)', '时间(s)', '网侧电流(A)', '电机电流(A)']}
    hh = {'group0': ['[0, 100]', '[0, 300]', '[-5000, 5000]', '[-500, 500]']}
    ii = {'group0': [10, 10, 10, 10]}
    jj = {'group0': ['red', 'blue', 'green', 'yellow']}
    kk = {'group0': "速度、时间、网侧电流、电机电流对运行里程曲线"}
    ll = ['明细_1000V_2M4T_AW0_上行.xls', '明细_1000V_2M4T_AW0_下行.xls', '明细_1000V_3M3T_AW0_上行.xls', '明细_1000V_3M3T_AW0_下行.xls', '明细_1000V_4M2T_AW0_上行.xls', '明细_1000V_4M2T_AW0_下行.xls', '明细_1500V_2M4T_AW0_上行.xls', '明细_1500V_2M4T_AW0_下行.xls', '明细_1500V_3M3T_AW0_上行.xls', '明细_1500V_3M3T_AW0_下行.xls', '明细_1500V_4M2T_AW0_上行.xls', '明细_1500V_4M2T_AW0_下行.xls', '明细_1800V_2M4T_AW0_上行.xls', '明细_1800V_2M4T_AW0_下行.xls', '明细_1800V_3M3T_AW0_上行.xls', '明细_1800V_3M3T_AW0_下行.xls', '明细_1800V_4M2T_AW0_上行.xls', '明细_1800V_4M2T_AW0_下行.xls']
