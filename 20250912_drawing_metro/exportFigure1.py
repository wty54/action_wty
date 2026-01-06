# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 14:34:42 2025

@author: thinkpad
"""

import os
os.environ['MPLBACKEND'] = 'Agg'
#import matplotlib
#current_backend = matplotlib.get_backend()
#if current_backend != 'Agg':
#    matplotlib.use('Agg')
from matplotlib import pyplot as plt
#import matplotlib.pyplot as plt
#from matplotlib.font_manager import FontProperties
import pandas as pd
#import numpy as np
#import os
import re

#def check_xls(Load_xls, Save_xls):
#    df = pd.read_excel(Load_xls)
# 
## 遍历DataFrame的行，检查特定列（假设是第一列）
#    for column in range(0, 16):
#        for index in range(1, len(df)):  # 从第二行开始，因为第一行是标题行或者你想从第二行数据开始处理
#            if df.iloc[index, column] == '-1.#IO':  # 检查第一列的值是否为'-1.#IO'
#                if index == 1:
#                    df.iloc[index, column] = 0
#                else:
#                    df.iloc[index, column] = df.iloc[index - 1, column]  # 如果是，则替换为上一行的值
#     
#    # 保存修改后的DataFrame到新的Excel文件
#    df.to_excel(Save_xls, index=False)
#列表字符串转换list
def cv_list(string):
    for pp in range(len(string)):
        numbers = re.findall(r'-?\d+\.?\d*', str(string[pp])) # 使用正则表达式查找字符串中的数字
        numbers = [int(num) for num in numbers]
        string[pp] = numbers
    return string

#定义绘制曲线图
def Four_lines(x0, y_list, y_lim, scale_list, color_list, lines_name, ST, ED, Load_data, folder_path, save_path, Station, Station_between, StationBetween_count, PIC_num, work_condition, State_ed):
#    global PIC_num  #先声明图片序号为全局变量
    x_data = x0[ST : ED]
    y_data = {}  #创建字典用于存放y轴数据
    for j in range(0, len(y_list)):
        y_data[f'y_data{j}'] = (y_list[j])[ST : ED]
#        if (bool(re.search("对运行里程曲线", lines_name))) and y_list[j].attribute_b == '时间(s)':
#            y_data[f'y_data{j}'] = [x - (y_data[f'y_data{j}'])[0] for x in y_data[f'y_data{j}']]      
    
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 指定默认字体为黑体
    plt.rcParams['axes.unicode_minus'] = False   # 解决保存图像是负号'-'显示为方块的问题
    plt.rcParams['xtick.labelsize'] = 6
    plt.rcParams['ytick.labelsize'] = 6
    plt.rcParams['axes.titlesize'] = 8    # 标题稍大一些
    plt.rcParams['axes.labelsize'] = 6    # 坐标轴标签
    axs = {}  #创建字典用于存放y轴
    #    lines = []
#    labels = []
    fig, axs[f'ax{0}'] = plt.subplots()
    #  创建x轴（里程）和y1轴（速度）
    axs[f'ax{0}'].plot(x_data, y_data[f'y_data{0}'], color_list[0], label=(y_list[0]).attribute_b)
    axs[f'ax{0}'].set_xlabel(x0.attribute_b, fontsize=8)  # x轴标签
    axs[f'ax{0}'].set_ylabel((y_list[0]).attribute_b, fontsize=8)  # y轴标签
    axs[f'ax{0}'].set_ylim((y_lim[0])[0], (y_lim[0])[1])  # 设置y轴的范围
    axs[f'ax{0}'].set_yticks(range((y_lim[0])[0], (y_lim[0])[1] + 1, ((y_lim[0])[1] - (y_lim[0])[0]) // scale_list[0]))
    axs[f'ax{0}'].tick_params(labelsize=6)  #标签字号设置为6
    line, label = axs[f'ax{0}'].get_legend_handles_labels()
    axis_count = 1  #初始化y轴计数器
    for i in range(1, len(y_list)):
        # 创建y2轴（时间）
        axs[f'ax{i}'] = axs[f'ax{i-1}'].twinx()
        axs[f'ax{i}'].spines['left'].set_position(('outward', 40 * i))  # 向左移动40像素
        axs[f'ax{i}'].yaxis.set_ticks_position('left')
        axs[f'ax{i}'].plot(x_data, y_data[f'y_data{i}'], color_list[i], label=(y_list[i]).attribute_b)
        axs[f'ax{i}'].set_ylabel((y_list[i]).attribute_b, fontsize=8)
        axs[f'ax{i}'].yaxis.set_label_position("left")  # 设置标签位置为左侧（实际上默认就是）
        axs[f'ax{i}'].set_ylim((y_lim[i])[0], (y_lim[i])[1])  # 设置y轴的范围
        axs[f'ax{i}'].set_yticks(range((y_lim[i])[0], (y_lim[i])[1] + 1, ((y_lim[i])[1] - (y_lim[i])[0]) // scale_list[i]))
        axs[f'ax{i}'].tick_params(labelsize=6)  #标签字号设置为6
        line_i, label_i = axs[f'ax{i}'].get_legend_handles_labels()
        line = line + line_i
        label= label + label_i
        axis_count += 1
#    aa = ax[f'ax{2}'].get_legend_handles_labels()
#    print(aa)
    #合并图例
    axs[f'ax{0}'].legend(line, label, loc="upper center", bbox_to_anchor=(0.4, -0.1), ncol=axis_count, fontsize=6)
    # 下行翻转横坐标
    if bool(re.search("下行", Load_data)) and x0.attribute_b == "公里标(km)":
        plt.gca().invert_xaxis()  # 翻转x轴，使得大的值在左侧，小的值在右侧
    plt.tight_layout()  # 自动调整子图参数，使之填充图像区域
    plt.subplots_adjust(bottom=0.15, top=0.95)  # 调整bottom和top参数来控制上下边距
#    if bool(re.search("对运行时间曲线", lines_name)):
    if ST == 0 and ED == len(x0):  #全程绘图设置标题
        plt.title(Station[0] + '——' + Station[-1], fontsize=8)  #线路起点——线路终点标题
        plt.axvline(x_data[0], ls='-.', c='r', lw=0.5)
        x_norm = (x_data[0] - plt.xlim()[0]) / (plt.xlim()[1] - plt.xlim()[0])
        plt.text(x_norm, 0.95, Station[0], transform=plt.gca().transAxes, fontsize=6, ha='left', va='top')
        x_scale = [x_data[0]]
#        plt.xticks([x_data[0]])
        for ss in range(len(State_ed)):
            x_norm = (x_data[State_ed[ss]] - plt.xlim()[0]) / (plt.xlim()[1] - plt.xlim()[0])
            plt.axvline(x_data[State_ed[ss]], ls='-.', c='r', lw=0.5)
            x_scale += [x_data[State_ed[ss]]]
#            plt.xticks(x_data[ss])
#            plt.xticks([x_data[State_ed[ss]]])
            if ss == len(State_ed) - 1:
                plt.text(x_norm, 0.95, Station[ss+1], transform=plt.gca().transAxes, fontsize=6, ha='right', va='top')
            else:
                plt.text(x_norm, 0.95, Station[ss+1], transform=plt.gca().transAxes, fontsize=6, ha='left', va='top')
        plt.xticks(x_scale)
#        plt.tick_params(axis='both', labelsize=6)  # 设置 x 轴刻度字号
#            print(Station[ss])
#        for Stat in Station:
#           plt.axvline(x_data[0], ls='-.', c='r', lw=0.5)  # 标记出起点和终点
#           plt.axvline(x_data[-1], ls='-.', c='r', lw=0.5)
    else:  #每站绘图设置标题
        plt.title(Station_between[StationBetween_count], fontsize=8)  # 图表标题
        plt.axvline(x_data[0], ls='-.', c='r', lw=0.5)  # 标记出起点和终点
        plt.axvline(x_data[-1], ls='-.', c='r', lw=0.5)
        plt.xticks([x_data[0], x_data[-1]])
#    plt.annotate('重要事件', xy=(x_data[0], 0), xytext=(0, 2), textcoords="offset points")  # va="top"表示文本顶部对齐
#    plt.gca().xaxis.set_major_formatter(ticker.NullFormatter())  # 隐藏中间的刻度标签（但不隐藏线）
    plt.rcParams['axes.grid'] = False
    plt.grid(True, linestyle='--', alpha=0.5)  # 显示网格
#    plt.show()  # 显示图表    
#    保存前检查并去除\
#    if bool(re.search("对运行时间曲线", lines_name)):
    if ST == 0 and ED == len(x0):  #全程绘图设置图名（站点）
        name1 = Station[0] + '——' + Station[-1]
    else:
        name1 = Station_between[StationBetween_count]
#    name2 = Speed.attribute_b.replace("/", "")
#     name3 = Kilo.attribute_b.replace("/", "")
#    savepath = folder_path + ' ' + str(StationBetween_count + 1) + name1 + '_' + name2 + '-' + name3 + '.jpg'
    savepath = save_path + '\\' + str(PIC_num) + ' ' + name1 + ' ' + work_condition + '_' + lines_name + '.jpg'
#    plt.savefig('C:\\Users\\thinkpad\\Desktop\\线路仿真绘图\\图片(kmh).jpg', dpi=300)  # dpi参数控制图片的分辨率

    fig.savefig(savepath, dpi=300)
#    plt.cla()
    plt.close()
#    PIC_num += 1

#定义一个含属性的数组类
class ListWithAttribute(list):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.attribute_b = None  # 例如，初始化属性b

def ex_Fig(xls_name, folder_path, Load_station, save_path, PIC_num, xaxisType_flag, xaxis_flag, yaxis_names, plot_lims, scale_lists, color_lists, plot_names):
#    global PIC_num
    # 设置地址
#    folder_path = 'C:\\Users\\thinkpad\\Desktop\\线路仿真绘图\\'
    Load_data = folder_path + '\\' + xls_name 
#    modify_data = Load_data
    name = re.sub(r'\.[^\.]*$', '', xls_name)
    information = re.split(r'[_]+', name)
    work_condition = information[1] + '_' + information[2] + '_' + information[3] + '_' + information[4]
#    work_condition = xls_name[3:20]
#    Load_station = 'C:\\Users\\thinkpad\\Desktop\\线路仿真绘图\\设施.xls'
    # 修正表格
#    check_xls(Load_data, modify_data)
    # 读取数据
    df1 = pd.read_excel(Load_data, sheet_name='全程明细数据')
    df2 = pd.read_excel(Load_station, sheet_name='线路设施')
    #将数据赋给向量
    State = ListWithAttribute(df1['工况'].tolist())
    Kilo = ListWithAttribute(df1['公里标(km)'].tolist())
    Period = ListWithAttribute(df1['时间(s)'].tolist())
    Speed = ListWithAttribute(df1['速度(km/h)'].tolist())
    Anti_force = ListWithAttribute(df1['阻力(kN)'].tolist())
    Motor_current = ListWithAttribute(df1['电机电流(A)'].tolist())
    Motor_power = ListWithAttribute(df1['电机输出功率(kW)'].tolist())
    Line_current = ListWithAttribute(df1['网侧电流(A)'].tolist())
    Line_power = ListWithAttribute(df1['网侧输入功率(kW)'].tolist())
    Energy = ListWithAttribute(df1['累计能耗(kWh)'].tolist())
    Accelerate = ListWithAttribute(df1['加速度(m/s^2)'].tolist())
    Period2 = ListWithAttribute(df1['时间(s)'].tolist())  #按停站修正时间
    #编辑向量属性
    State.attribute_b = '工况'
    Kilo.attribute_b = '公里标(km)'
    Period.attribute_b = '时间(s)'
    Speed.attribute_b = '速度(km/h)'
    Anti_force.attribute_b = '阻力(kN)'
    Motor_current.attribute_b = '电机电流(A)'
    Motor_power.attribute_b = '电机输出功率(kW)'
    Line_current.attribute_b = '网侧电流(A)'
    Line_power.attribute_b = '网侧输入功率(kW)'
    Energy.attribute_b = '累计能耗(kWh)'
    Accelerate.attribute_b = '加速度(m/s^2)'
    Period2.attribute_b = '时间(s)'
#    Period_lim = [0, 300]
#    Speed_lim = [0, 100]
#    Line_current_lim = [-5000, 5000]
#    Motor_current_lim = [-500, 500]
#    Energy_lim = [0, 500]
#    Line_power_lim = [-8000, 8000]
#    Motor_power_lim = [-800, 800]
#    Anti_force_lim = [-100, 100]
    #构造绘图变量数组，用于列表查找与循环
    plot_lists = [Speed, Line_current, Motor_current, Anti_force, Motor_power, Energy, Line_power, Accelerate]
    #修改明细表中的错误值'-1.#IO'
    for aa in range(len(plot_lists)):
        for bb in range(len(plot_lists[aa])):
            if plot_lists[aa][bb] == '-1.#IO':
                if bb == 0:
                    plot_lists[aa][bb] = 0.0
                else:
                    plot_lists[aa][bb] = plot_lists[aa][bb - 1]
            else:
                plot_lists[aa][bb] = float(plot_lists[aa][bb])
                
    #判断仿真明细的名称是否有上下行或者平直道
    if bool(re.search("上行", Load_data)):
        Station = df2['设施名称'].tolist()
#        Sta_kilo = df2['里程'].tolist()
    elif bool(re.search("下行", Load_data)):
        Station = df2['设施名称'].tolist()
#        Sta_kilo = df2['里程'].tolist()
        Station.reverse()
#        Sta_kilo.reverse()
    elif bool(re.search("平直道", Load_data)):
        Station = ["起点", "终点"]
    Station_between = [f'{Station[i]}——{Station[i+1]}' for i in range(len(Station) - 1)]
    
    #定义工况起车和停车计数器数组
    State_st = []
    State_ed = []
    #定义x轴选择字典键数组
    x_zero = []
    x_one = []
    #定义选定y轴数据字典
    yList_all = {}
    #按停车分段区间
    for State_count in range(len(State) - 1):
        if State[State_count] == "停车":
            if State[State_count+1] != "停车":
                State_st += [State_count]
            else:
                continue
        else:
            if State[State_count+1] == "停车":
                State_ed += [State_count + 1]
            else:
                continue
    
    #修正时间为每站计时
    for STED_count in range(len(State_st)):
        x0 = State_st[STED_count]
        x1 = State_ed[STED_count]
        Period2_temp = Period2[x0]
        for x in range(x0, len(Period2)):
            if STED_count+1 < len(State_st):
                if x0 <= x <= x1:
#                    print(Period2[x])
#                    print(Period2[x0])
#                    Period2_temp = Period2[x0]
                    Period2[x] = Period2[x] - Period2_temp
#                    print(Period2[x])
                elif x1 < x < State_st[STED_count+1]:
                    Period2[x] = None
                else:
                    break
            else:
                Period2[x] = Period2[x] - Period2_temp
    #将修正时间加入y轴数组
    plot_lists.insert(1, Period2)            
    #将选定的y轴数据传给yList_all
    for k in range(len(yaxis_names)):
        yList_all[f'group{k}'] = list(filter(lambda y: any(yaxis_name in y.attribute_b for yaxis_name in yaxis_names[f'group{k}']), plot_lists))
        yList_all[f'group{k}'] = sorted(yList_all[f'group{k}'], key = lambda x : yaxis_names[f'group{k}'].index(x.attribute_b))
        plot_lims[f'group{k}'] = cv_list(plot_lims[f'group{k}'])
#        y_list = yList_all[f'group{k}']
    #构建区分x轴每站和全程的数组   
    for key, value in xaxisType_flag.items():
        if value in [0]:
            x_zero.append(key)
        else:
            x_one.append(key)
    
    for StationBetween_count in range(len(Station_between)):
    # 绘制运行时间、速度、网侧电流、电机电流对运行里程曲线    
        START = State_st[StationBetween_count]
        END = State_ed[StationBetween_count] + 1
        for j in range(len(x_zero)):
            y_list = yList_all[x_zero[j]]
            plot_lim = plot_lims[x_zero[j]]
#            plot_lim = cv_list(plot_lim)
            scale_list = scale_lists[x_zero[j]]
            color_list = color_lists[x_zero[j]]
            plot_name = plot_names[x_zero[j]]
            if xaxis_flag[x_zero[j]] == 0:
                Four_lines(Kilo, y_list, plot_lim, scale_list, color_list, plot_name, START, END, Load_data, folder_path, save_path, Station, Station_between, StationBetween_count, PIC_num, work_condition, State_ed)
            else:
                Four_lines(Period2, y_list, plot_lim, scale_list, color_list, plot_name, START, END, Load_data, folder_path, save_path, Station, Station_between, StationBetween_count, PIC_num, work_condition, State_ed)
            PIC_num += 1
    for m in range(len(x_one)):
        y_list = yList_all[x_one[m]]
        plot_lim = plot_lims[x_one[m]]
#        plot_lim = cv_list(plot_lim)
        scale_list = scale_lists[x_one[m]]
        color_list = color_lists[x_one[m]]
        plot_name = plot_names[x_one[m]]
        if xaxis_flag[x_one[m]] == 0:
            Four_lines(Kilo, y_list, plot_lim, scale_list, color_list, plot_name, 0, len(Kilo), Load_data, folder_path, save_path, Station, Station_between, 0, PIC_num, work_condition, State_ed)
        else:
            Four_lines(Period, y_list, plot_lim, scale_list, color_list, plot_name, 0, len(Period), Load_data, folder_path, save_path, Station, Station_between, 0, PIC_num, work_condition, State_ed)
        PIC_num += 1
        
    return PIC_num
if __name__ == '__main__':
    aa = '明细_1500V_4M2T_AW3_上行.xls'  #文件名称
    bb = 'C:\\Users\\thinkpad\\Desktop\\线路仿真绘图'  #打开文件目录
    cc = 'C:\\Users\\thinkpad\\Desktop\\线路仿真绘图\\设施.xls'  #站点文件路径
    dd = 'C:\\Users\\thinkpad\\Desktop\\线路仿真绘图' #图片保存目录
    ee = 1  #初始图片序号PIC_num形参
#    ff = {'group0' : 0,
#          'group1' : 0,
#          'group2' : 1,
#          'group3' : 1}  #x轴站点绘图选择，0——每站，1——全程
#    gg = {'group0' : 0,
#          'group1' : 0,
#          'group2' : 1,
#          'group3' : 1}  #x轴选择，0——运行里程  1——运行时间
#    hh = {
#            'group0' : ['速度(km/h)', '时间(s)', '网侧电流(A)', '电机电流(A)'], 
#            'group1' : ['累计能耗(kWh)', '网侧输入功率(kW)', '电机输出功率(kW)'],
#            'group2' : ['速度(km/h)', '网侧电流(A)', '电机电流(A)'],
#            'group3' : ['累计能耗(kWh)', '网侧输入功率(kW)', '电机输出功率(kW)']            
#         }  #绘图选择
#    ii = {
#            'group0' : ['[0, 100]', '[0, 300]', '[-5000, 5000]', '[-500, 500]'], 
#            'group1' : ['[0, 500]', '[-8000, 8000]', '[-800, 800]'],
#            'group2' : ['[0, 100]', '[-5000, 5000]', '[-500, 500]'],
#            'group3' : ['[0, 500]', '[-8000, 8000]', '[-800, 800]']
#         }  #范围
#    jj = {
#            'group0' : [10, 10, 10, 10], 
#            'group1' : [10, 10, 10],
#            'group2' : [10, 10, 10],
#            'group3' : [10, 10, 10]
#         }  #格数
#    kk = {
#            'group0' : ['red', 'blue', 'green', 'cyan'], 
#            'group1' : ['magenta', 'yellow', 'black'],
#            'group2' : ['red', 'green', 'cyan'],
#            'group3' : ['magenta', 'yellow', 'black']
#         }  #颜色
#    ll = {'group0' : "运行时间、速度、网侧电流、电机电流对运行里程曲线", 
#          'group1' : "累计能耗、网侧输入功率、电机输出功率对运行里程曲线",
#          'group2' : "速度、网侧电流、电机电流对运行时间曲线",
#          'group3' : "累计能耗、网侧输入功率、电机输出功率对运行时间曲线"
#         }  #图名
    

#    ff = ['累计能耗(kWh)', '网侧输入功率(kW)', '电机输出功率(kW)']  #绘图选择
#    gg = [[0, 500], [-8000, 8000], [-800, 800]]  #范围
#    hh = [10, 10, 10]  #格数
#    ii = ['magenta', 'yellow', 'black']  #颜色
    
    ff = {'group0': 1}
    gg = {'group0': 1}
    hh = {'group0': ['速度(km/h)']}
    ii = {'group0': ['[0, 150]']}
    jj = {'group0': [10]}
    kk = {'group0': ['red']}
    ll = {'group0': "速度对运行时间曲线"}
    NO = ex_Fig(aa, bb, cc, dd, ee, ff, gg, hh, ii, jj, kk, ll)  #输出下一个明细第一个图片序号
    
    
    
    