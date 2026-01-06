# -*- coding: utf-8 -*-
"""
Created on Sun Mar 16 01:49:52 2025

@author: thinkpad
"""

import pandas as pd
import numpy as np
import os
#from exportFigure1 import ex_Fig

def time_to_seconds(time_str):
    # 将时间字符串分割为时、分、秒
    hours, minutes, seconds = map(int, time_str.split(':'))
    # 将时分秒转换为总秒数
    total_seconds = hours * 3600 + minutes * 60 + seconds
    return total_seconds

def seconds_to_time(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    remaining_seconds = seconds % 60
 
    # 使用f-string格式化输出
    formatted_time = f"{hours:02d}:{minutes:02d}:{remaining_seconds:02d}"
    return formatted_time

def cal_total(folder_path, save_path, distance, holdTime, file_names):
    # 指定文件夹路径
#    folder_path = 'C:\\Users\\thinkpad\\Desktop\\20250320\\惰行'
    all_data = pd.DataFrame()  # 初始化空的DataFrame来存储所有数据
    data_cache = []  #设置空元祖用于暂存上、下行统计数据
    total_cal = []  #设置空数组计算上下行统计
#    distance = 24.44  #总里程，用于计算旅行速度
#    holdTime = 120  #折返时间120s
    ii = 0  #表格中插入文件名
    jj = 0  #用于计算往返数据的计数 
    try:
        for file in file_names:
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(file_path, header=None)  # 读取不使用header，因为没有列名
            second_row = df.iloc[2]  # 获取第三行数据（索引从0开始，所以是2）
            all_data = all_data.append(second_row, ignore_index=True)  # 添加到总DataFrame中
            all_data.iloc[ii, 0] = file
            data_cache.append(second_row.tolist()[4:])
            ii += 1
            jj += 1
            if jj < 2:  
                continue
            else:
        #        print(data_cache[0][1], data_cache[1][1])
                up_time = time_to_seconds(data_cache[0][0])
                down_time = time_to_seconds(data_cache[1][0])
                total_cal.append(np.nan)
                total_cal.append(np.nan)
                total_cal.append(np.nan)
                total_cal.append(np.nan)
                total_cal += [seconds_to_time(up_time + down_time + holdTime)]  #行驶时间
                total_cal += [round(7200 * distance / (up_time + down_time + holdTime), 1)]  #旅行速度
                total_cal += [round((float(data_cache[0][2]) + float(data_cache[1][2])) / 2, 1)]  #平均速度
                total_cal += [round(((float(data_cache[0][3])**2 * up_time + float(data_cache[1][3])**2 * down_time) / (up_time + down_time + holdTime))**0.5, 0)]  #电机牵引RMS电流
                total_cal += [round(((float(data_cache[0][4])**2 * up_time + float(data_cache[1][4])**2 * down_time) / (up_time + down_time + holdTime))**0.5, 0)]  #电机制动RMS电流（无再生）
                total_cal += [int(data_cache[0][5]) + int(data_cache[1][5])]  #牵引消费电量
                total_cal += [int(data_cache[0][6]) + int(data_cache[1][6])]  #制动消费电量
                total_cal += [int(data_cache[0][7]) + int(data_cache[1][7])]  #总计消费电量（100%再生）
                total_cal += [int(data_cache[0][8]) + int(data_cache[1][8])]  #总计消费电量（50%再生）
                total_cal += [int(data_cache[0][9]) + int(data_cache[1][9])]  #总计消费电量（15%再生）
                total_cal += [round((float(data_cache[0][10]) + float(data_cache[1][10])) / 2, 1)]  #每公里消费电量（100%再生）
                total_cal += [round((float(data_cache[0][11]) + float(data_cache[1][11])) / 2, 1)]  #每公里消费电量（50%再生）
                total_cal += [round((float(data_cache[0][12]) + float(data_cache[1][12])) / 2, 1)]  #每公里消费电量（15%再生）
                total_cal += [round((float(data_cache[0][13]) + float(data_cache[1][13])) / 2, 1)]  #再生率（100%再生）
                total_cal += [round((float(data_cache[0][14]) + float(data_cache[1][14])) / 2, 1)]  #再生率（50%再生）
                total_cal += [round((float(data_cache[0][15]) + float(data_cache[1][15])) / 2, 1)]  #再生率（15%再生）
                total_cal += [round(((float(data_cache[0][16])**2 * up_time + float(data_cache[1][16])**2 * down_time) / (up_time + down_time + holdTime))**0.5, 0)]  #网侧RMS电流
                all_data = all_data.append(pd.DataFrame([total_cal], columns=all_data.columns), ignore_index=True)
                ii += 1
                total_cal = []
                data_cache = []
                jj = 0
        output_path = save_path + '\\全程统计.xlsx'  # 输出文件的路径和名称
        all_data.to_excel(output_path, index=False, header=False)  # 写入Excel文件，不包含索引列
        return 0, None  # 返回状态和错误信息
    except Exception as e:
        return 1, file
        
if __name__ == '__main__':
    folder_path = 'C:\\Users\\thinkpad\\Desktop\\20250912\\无惰行统计'
    save_path = 'C:\\Users\\thinkpad\\Desktop\\20250912'
    distance = 25.41
    holdTime = 120
    file_names = ["统计_1500V_4M2T_AW3_上行.xls", "统计_1500V_4M2T_AW3_下行.xls", "统计_1500V_4M2T_AW2_上行.xls", "统计_1500V_4M2T_AW2_下行.xls", "统计_1500V_4M2T_AW0_上行.xls", "统计_1500V_4M2T_AW0_下行.xls", "统计_1800V_3M3T_AW3_上行.xls", "统计_1800V_3M3T_AW3_下行.xls", "统计_1800V_4M2T_AW2_上行.xls", "统计_1800V_4M2T_AW2_下行.xls", "统计_1800V_4M2T_AW0_上行.xls", "统计_1800V_4M2T_AW0_下行.xls"]
    cal_total(folder_path, save_path, distance, holdTime, file_names)