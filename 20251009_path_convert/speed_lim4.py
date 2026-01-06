# -*- coding: utf-8 -*-
"""
Created on Thu Mar 21 09:13:34 2024

@author: thinkpad
"""

import pandas as pd
#import openpyxl

def speed_limit(Load, Save, limitation, orientation):
    df = pd.read_excel(Load, usecols = [0, 1, 3], names = ["起点里程", "终点里程", "限速"])
#    df.rename(columns = {'曲线半径' : '限速'}, inplace = True)
    if(orientation == 1):  #0——下行递减  1——上行递增
        df = df.sort_values(by = ['起点里程'], ascending = False)
#        df['终点里程', '限速'] = df[['终点里程'], ['限速']].sort_values(by = ['起点里程'])
        df[['起点里程', '终点里程']] = df[['终点里程', '起点里程']]
#    df_sorted = df.sort_values(by = ['起点里程', '终点里程', '曲线半径'], inplace = True)
    df['限速'] = df['限速'] ** 0.5 * 3.91
    df['限速'] = df['限速'].round(decimals = 1)
    df['限速'].loc[df['限速'] > limitation] = limitation
    df.to_excel(Save, sheet_name='限速', index = False)
    return 1

if __name__ == '__main__':
    aa = 'C:\\Users\\thinkpad\\Desktop\\西安6号线线路条件\\曲线（上行）.xlsx'
    bb = 'C:\\Users\\thinkpad\\Desktop\\西安6号线线路条件\\限速（上行）.xlsx'
    cc = 80
    dd = 0
    speed_limit(aa, bb, cc, dd)