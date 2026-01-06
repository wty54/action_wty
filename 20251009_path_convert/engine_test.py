# -*- coding: utf-8 -*-
"""
Created on Fri Jan 19 15:54:44 2024

@author: thinkpad
"""

import tkinter as tk
import tkinter.ttk as ttk
import os
import sys
from tkinter import messagebox, filedialog
#import subprocess
from decline_transform1 import transform
from curve_transform2 import ruler
from station3 import station
from speed_lim4 import speed_limit
#import type_test

class Application(tk.Frame):
    def __init__(self, master = None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()
    def create_widgets(self):
        #坡度坡长组件
        #       打开文件路径按钮
        group1 = ttk.Labelframe(self, text="坡度/坡长")
        group1.grid(row = 1, column = 0, padx = 5, pady = 5)
        
        group2 = ttk.Labelframe(self, text="曲线")
        group2.grid(row = 2, column = 0, padx = 5, pady = 5)
        
        group3 = ttk.Labelframe(self, text="设施")
        group3.grid(row = 3, column = 0, padx = 5, pady = 5)
        
        group4 = ttk.Labelframe(self, text="限速")
        group4.grid(row = 4, column = 0, padx = 5, pady = 5)
        
        self.OpenEntry1 = ttk.Entry(group1, textvariable = filename1)
        self.OpenEntry1.grid(row = 1, column = 0, padx = 5, pady = 5)

        self.btn01 = ttk.Button(group1)
        self.btn01["text"] = "选择文件"
        self.btn01.grid(row = 1, column = 1, padx = 5, pady = 5)
        self.btn01["command"] = self.open_dxf1
#        filename.set(self.open_dxf)
#       保存文件路径按钮        
        
        self.SaveEntry1 = ttk.Entry(group1, textvariable = path_var1)
        self.SaveEntry1.grid(row = 2, column = 0, padx = 5, pady = 5)
        
        self.btn02 = ttk.Button(group1)
        self.btn02["text"] = "保存文件"
        self.btn02.grid(row = 2, column = 1, padx = 5, pady = 5)
        self.btn02["command"] = self.save_xls1
        
#       配置按钮
        self.btn04 = ttk.Button(group1)
        self.btn04["text"] = "配置"
        self.btn04.grid(row = 1, column = 2, padx = 5, pady = 5)
        self.btn04["command"] = self.open_popup1
        
        #曲线组件
        self.OpenEntry2 = ttk.Entry(group2, textvariable = filename2)
        self.OpenEntry2.grid(row = 1, column = 0, padx = 5, pady = 5)

        self.btn05 = ttk.Button(group2)
        self.btn05["text"] = "选择文件"
        self.btn05.grid(row = 1, column = 1, padx = 5, pady = 5)
        self.btn05["command"] = self.open_dxf2
#        filename.set(self.open_dxf)
#       保存文件路径按钮        
        
        self.SaveEntry2 = ttk.Entry(group2, textvariable = path_var2)
        self.SaveEntry2.grid(row = 2, column = 0, padx = 5, pady = 5)
        
        self.btn06 = ttk.Button(group2)
        self.btn06["text"] = "保存文件"
        self.btn06.grid(row = 2, column = 1, padx = 5, pady = 5)
        self.btn06["command"] = self.save_xls2

#       配置按钮
        self.btn08 = ttk.Button(group2)
        self.btn08["text"] = "配置"
        self.btn08.grid(row = 1, column = 2, padx = 5, pady = 5)
        self.btn08["command"] = self.open_popup2
        
        #设施组件
        self.OpenEntry3 = ttk.Entry(group3, textvariable = filename3)
        self.OpenEntry3.grid(row = 1, column = 0, padx = 5, pady = 5)

        self.btn09 = ttk.Button(group3)
        self.btn09["text"] = "选择文件"
        self.btn09.grid(row = 1, column = 1, padx = 5, pady = 5)
        self.btn09["command"] = self.open_dxf3
#        filename.set(self.open_dxf)
#       保存文件路径按钮        
        
        self.SaveEntry3 = ttk.Entry(group3, textvariable = path_var3)
        self.SaveEntry3.grid(row = 2, column = 0, padx = 5, pady = 5)
        
        self.btn10 = ttk.Button(group3)
        self.btn10["text"] = "保存文件"
        self.btn10.grid(row = 2, column = 1, padx = 5, pady = 5)
        self.btn10["command"] = self.save_xls3
        
#       配置按钮
        self.btn11 = ttk.Button(group3)
        self.btn11["text"] = "配置"
        self.btn11.grid(row = 1, column = 2, padx = 5, pady = 5)
        self.btn11["command"] = self.open_popup3
        
        #限速组件
        self.OpenEntry4 = ttk.Entry(group4, textvariable = filename4)
        self.OpenEntry4.grid(row = 1, column = 0, padx = 5, pady = 5)

        self.btn12 = ttk.Button(group4)
        self.btn12["text"] = "选择文件"
        self.btn12.grid(row = 1, column = 1, padx = 5, pady = 5)
        self.btn12["command"] = self.open_xls4
#        filename.set(self.open_dxf)
#       保存文件路径按钮        
        
        self.SaveEntry4 = ttk.Entry(group4, textvariable = path_var4)
        self.SaveEntry4.grid(row = 2, column = 0, padx = 5, pady = 5)
        
        self.btn13 = ttk.Button(group4)
        self.btn13["text"] = "保存文件"
        self.btn13.grid(row = 2, column = 1, padx = 5, pady = 5)
        self.btn13["command"] = self.save_xls4
#        path_var.set(self.save_xls)
#       运行按钮
#        self.btn03 = ttk.Button(self)
#        self.btn03["text"] = "运行"
#        self.btn03.grid(row = 3, column = 1, padx = 5, pady = 5)
#        self.btn03["command"] = self.run_program
        
#       配置按钮
        self.btn14 = ttk.Button(group4)
        self.btn14["text"] = "配置"
        self.btn14.grid(row = 1, column = 2, padx = 5, pady = 5)
        self.btn14["command"] = self.open_popup4
#坡度/坡长打开路径    
    def open_dxf1(self):
        try:
            filetypes = [("DXF", ".dxf")]
            filepath = filedialog.askopenfilename(title = '打开文件', 
                      filetypes = filetypes, defaultextension = '.dxf', 
                      initialdir = 'C:/Users/thinkpad/Desktop')
            filename1.set(filepath)
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')
#曲线打开路径    
    def open_dxf2(self):
        try:
            filetypes = [("DXF", ".dxf")]
            filepath = filedialog.askopenfilename(title = '打开文件', 
                      filetypes = filetypes, defaultextension = '.dxf', 
                      initialdir = 'C:/Users/thinkpad/Desktop')
            filename2.set(filepath)
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')
#设施打开路径    
    def open_dxf3(self):
        try:
            filetypes = [("DXF", ".dxf")]
            filepath = filedialog.askopenfilename(title = '打开文件', 
                      filetypes = filetypes, defaultextension = '.dxf', 
                      initialdir = 'C:/Users/thinkpad/Desktop')
            filename3.set(filepath)
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')
#限速打开路径            
    def open_xls4(self):
        try:
            filetypes = [("XLSX", ".xlsx"), ("XLS", ".xls")]
            filepath = filedialog.askopenfilename(title = '打开文件', 
                      filetypes = filetypes, defaultextension = '.xlsx', 
                      initialdir = 'C:/Users/thinkpad/Desktop')
            filename4.set(filepath)
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')

#坡度/坡长保存路径    
    def save_xls1(self):
        try:
            filetypes = [("XLSX", ".xlsx"), ("XLS", ".xls")]
            filenewpath = filedialog.asksaveasfilename(title = '保存文件', 
                      filetypes = filetypes, defaultextension = '.xlsx', 
                      initialdir = 'C:/Users/thinkpad/Desktop' )
#            self.create_widgets.path_var = filenewpath
            path_var1.set(filenewpath)
        #保存文件

        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')
#曲线保存路径     
    def save_xls2(self):
        try:
            filetypes = [("XLSX", ".xlsx"), ("XLS", ".xls")]
            filenewpath = filedialog.asksaveasfilename(title = '保存文件', 
                      filetypes = filetypes, defaultextension = '.xlsx', 
                      initialdir = 'C:/Users/thinkpad/Desktop' )
#            self.create_widgets.path_var = filenewpath
            path_var2.set(filenewpath)
        #保存文件

        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')
#设施保存路径     
    def save_xls3(self):
        try:
            filetypes = [("XLSX", ".xlsx"), ("XLS", ".xls")]
            filenewpath = filedialog.asksaveasfilename(title = '保存文件', 
                      filetypes = filetypes, defaultextension = '.xlsx', 
                      initialdir = 'C:/Users/thinkpad/Desktop' )
#            self.create_widgets.path_var = filenewpath
            path_var3.set(filenewpath)
        #保存文件

        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件') 
#限速保存路径     
    def save_xls4(self):
        try:
            filetypes = [("XLSX", ".xlsx"), ("XLS", ".xls")]
            filenewpath = filedialog.asksaveasfilename(title = '保存文件', 
                      filetypes = filetypes, defaultextension = '.xlsx', 
                      initialdir = 'C:/Users/thinkpad/Desktop' )
#            self.create_widgets.path_var = filenewpath
            path_var4.set(filenewpath)
        #保存文件

        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')             
            
#坡度/坡长配置子窗口
    def open_popup1(self):
        popup = tk.Toplevel(self)
        center_window(popup)
        popup.geometry("300x251")
        popup.title("坡长/坡度转换")
#        popup.attributes('-topmost', True)    
        
#        x0 = tk.StringVar()
#        x1 = tk.StringVar()
#        y0 = tk.StringVar()
#        y1 = tk.StringVar()
        
#        e1 = ttk.Entry(popup, textvariable = x0)
#        e2 = ttk.Entry(popup, textvariable = x1)
#        e3 = ttk.Entry(popup, textvariable = y0)
#        e4 = ttk.Entry(popup, textvariable = y1)
        
        self.e1 = ttk.Entry(popup)
        self.e2 = ttk.Entry(popup)
        self.e3 = ttk.Entry(popup)
        self.e4 = ttk.Entry(popup)
        self.e12 = ttk.Entry(popup)
        
        self.r1 = ttk.Radiobutton(popup, text = "单行文字", value = 0, variable = text_type1)
        self.r2 = ttk.Radiobutton(popup, text = "多行文字", value = 1, variable = text_type1)
        
        ttk.Label(popup, text = "x坐标最小值").grid(row = 1,column = 2,
          padx = 5, pady = 5)
        self.e1.grid(row = 1,column = 3, padx = 5, pady = 5)
        ttk.Label(popup, text = "x坐标最大值").grid(row = 2,column = 2,
          padx = 5, pady = 5)
        self.e2.grid(row = 2,column = 3, padx = 5, pady = 5)
        ttk.Label(popup, text = "y坐标最小值").grid(row = 3,column = 2,
          padx = 5, pady = 5)
        self.e3.grid(row = 3,column = 3, padx = 5, pady = 5)
        ttk.Label(popup, text = "y坐标最大值").grid(row = 4,column = 2,
          padx = 5, pady = 5)
        self.e4.grid(row = 4,column = 3, padx = 5, pady = 5)   
        ttk.Label(popup, text = "起始公里标（km）").grid(row = 6,column = 2,
          padx = 5, pady = 5)
        self.e12.grid(row = 6,column = 3, padx = 5, pady = 5) 
                
        self.r1.grid(row = 5,column = 2, padx = 5, pady = 5)
        self.r2.grid(row = 5,column = 3, padx = 5, pady = 5)
    
        ttk.Button(popup, text = "运行", command = self.ConfigValue).grid(row = 7,column = 2,
          padx = 5, pady = 5)
        ttk.Button(popup, text = "关闭", command = popup.destroy).grid(row = 7,column = 3,
          padx = 5, pady = 5)
#曲线配置子窗口
    def open_popup2(self):
        popup2 = tk.Toplevel(self)
        center_window(popup2)
        popup2.geometry("650x400")
        popup2.title("曲线转换")
#        popup2.attributes('-topmost', True)    
        
#        x0 = tk.StringVar()
#        x1 = tk.StringVar()
#        y0 = tk.StringVar()
#        y1 = tk.StringVar()
        
#        e1 = ttk.Entry(popup, textvariable = x0)
#        e2 = ttk.Entry(popup, textvariable = x1)
#        e3 = ttk.Entry(popup, textvariable = y0)
#        e4 = ttk.Entry(popup, textvariable = y1)
        group21 = ttk.Labelframe(popup2, text="曲线区域")
        group21.grid(row = 1, column = 0, padx = 5, pady = 5)
        group22 = ttk.Labelframe(popup2, text="公里标区域")
        group22.grid(row = 1, column = 1, padx = 5, pady = 5)
        
        self.e5 = ttk.Entry(group21)
        self.e6 = ttk.Entry(group21)
        self.e7 = ttk.Entry(group21)
        self.e8 = ttk.Entry(group21)
        self.e9 = ttk.Entry(group22)
        self.e10 = ttk.Entry(group22)
        self.e11 = ttk.Entry(group22)
        self.e19 = ttk.Entry(group22)
        self.e20 = ttk.Entry(group21)
        self.e21 = ttk.Entry(group22)
        self.e22 = ttk.Entry(group22)
        self.e23 = ttk.Entry(group21)
        self.e24 = ttk.Entry(group21)       
        
        self.r3 = ttk.Radiobutton(group21, text = "单行文字", value = 1, variable = text_type2)
        self.r4 = ttk.Radiobutton(group21, text = "多行文字", value = 0, variable = text_type2)
        
        self.r5 = ttk.Radiobutton(group21, text = "多段线", value = 1, variable = curve_type2)
        self.r6 = ttk.Radiobutton(group21, text = "直线", value = 0, variable = curve_type2)
        
        ttk.Label(group21, text = "x坐标最小值").grid(row = 1,column = 2,
          padx = 5, pady = 5)
        self.e5.grid(row = 1,column = 3, padx = 5, pady = 5)
        ttk.Label(group21, text = "x坐标最大值").grid(row = 2,column = 2,
          padx = 5, pady = 5)
        self.e6.grid(row = 2,column = 3, padx = 5, pady = 5)
        ttk.Label(group21, text = "y坐标最小值").grid(row = 3,column = 2,
          padx = 5, pady = 5)
        self.e7.grid(row = 3,column = 3, padx = 5, pady = 5)
        ttk.Label(group21, text = "y坐标最大值").grid(row = 4,column = 2,
          padx = 5, pady = 5)
        self.e8.grid(row = 4,column = 3, padx = 5, pady = 5)   
        ttk.Label(group21, text = "曲线位置文字图层").grid(row = 5,column = 2,
          padx = 5, pady = 5)
        self.e21.grid(row = 5,column = 3, padx = 5, pady = 5) 
        ttk.Label(group21, text = "长度关键字").grid(row = 6,column = 2,
          padx = 5, pady = 5)
        self.e23.grid(row = 6,column = 3, padx = 5, pady = 5)
        ttk.Label(group21, text = "半径关键字").grid(row = 7,column = 2,
          padx = 5, pady = 5)
        self.e24.grid(row = 7,column = 3, padx = 5, pady = 5)
        ttk.Label(group21, text = "半径长度文字类型").grid(row = 8,column = 2,
          padx = 5, pady = 5)
        self.r3.grid(row = 8,column = 3, padx = 5, pady = 5)
        self.r4.grid(row = 8,column = 4, padx = 5, pady = 5)
        ttk.Label(group21, text = "曲线绘图线类型").grid(row = 9,column = 2,
          padx = 5, pady = 5)
        self.r5.grid(row = 9,column = 3, padx = 5, pady = 5)
        self.r6.grid(row = 9,column = 4, padx = 5, pady = 5)
        
        ttk.Label(group22, text = "x坐标最小值").grid(row = 1,column = 2,
          padx = 5, pady = 5)
        self.e9.grid(row = 1,column = 3, padx = 5, pady = 5)
        ttk.Label(group22, text = "x坐标最大值").grid(row = 2,column = 2,
          padx = 5, pady = 5)
        self.e10.grid(row = 2,column = 3, padx = 5, pady = 5)
        ttk.Label(group22, text = "y坐标最小值").grid(row = 3,column = 2,
          padx = 5, pady = 5)
        self.e11.grid(row = 3,column = 3, padx = 5, pady = 5)
        ttk.Label(group22, text = "y坐标最大值").grid(row = 4,column = 2,
          padx = 5, pady = 5)
        self.e19.grid(row = 4,column = 3, padx = 5, pady = 5)   
        ttk.Label(group22, text = "公里标尺文字图层").grid(row = 5,column = 2,
          padx = 5, pady = 5)
        self.e20.grid(row = 5,column = 3, padx = 5, pady = 5)
        ttk.Label(group22, text = "最小整数公里标").grid(row = 6,column = 2,
          padx = 5, pady = 5)
        self.e22.grid(row = 6,column = 3, padx = 5, pady = 5)
                
        ttk.Button(popup2, text = "运行", command = self.ConfigValue2).grid(row = 2,column = 0,
          padx = 5, pady = 5)
        ttk.Button(popup2, text = "关闭", command = popup2.destroy).grid(row = 2,column = 1,
          padx = 5, pady = 5)

#设施配置子窗口
    def open_popup3(self):
        popup3 = tk.Toplevel(self)
        center_window(popup3)
        popup3.geometry("300x251")
        popup3.title("设施转换")
#        popup2.attributes('-topmost', True)    
        
#        x0 = tk.StringVar()
#        x1 = tk.StringVar()
#        y0 = tk.StringVar()
#        y1 = tk.StringVar()
        
#        e1 = ttk.Entry(popup, textvariable = x0)
#        e2 = ttk.Entry(popup, textvariable = x1)
#        e3 = ttk.Entry(popup, textvariable = y0)
#        e4 = ttk.Entry(popup, textvariable = y1)
        
        self.e13 = ttk.Entry(popup3)
        self.e14 = ttk.Entry(popup3)
        self.e15 = ttk.Entry(popup3)
        self.e16 = ttk.Entry(popup3)
        self.e17 = ttk.Entry(popup3)
        self.e18 = ttk.Entry(popup3)
        
        ttk.Label(popup3, text = "x坐标最小值").grid(row = 1,column = 2,
          padx = 5, pady = 5)
        self.e13.grid(row = 1,column = 3, padx = 5, pady = 5)
        ttk.Label(popup3, text = "x坐标最大值").grid(row = 2,column = 2,
          padx = 5, pady = 5)
        self.e14.grid(row = 2,column = 3, padx = 5, pady = 5)
        ttk.Label(popup3, text = "y坐标最小值").grid(row = 3,column = 2,
          padx = 5, pady = 5)
        self.e15.grid(row = 3,column = 3, padx = 5, pady = 5)
        ttk.Label(popup3, text = "y坐标最大值").grid(row = 4,column = 2,
          padx = 5, pady = 5)
        self.e16.grid(row = 4,column = 3, padx = 5, pady = 5)   
        ttk.Label(popup3, text = "站点图层").grid(row = 5,column = 2,
          padx = 5, pady = 5)
        self.e17.grid(row = 5,column = 3, padx = 5, pady = 5)
        ttk.Label(popup3, text = "公里标关键字").grid(row = 6,column = 2,
          padx = 5, pady = 5)
        self.e18.grid(row = 6,column = 3, padx = 5, pady = 5)
       
        ttk.Button(popup3, text = "运行", command = self.ConfigValue3).grid(row = 7,column = 2,
          padx = 5, pady = 5)
        ttk.Button(popup3, text = "关闭", command = popup3.destroy).grid(row = 7,column = 3,
          padx = 5, pady = 5)
#限速配置子窗口
    def open_popup4(self):
        popup4 = tk.Toplevel(self)
        center_window(popup4)
        popup4.geometry("300x100")
        popup4.title("限速转换")
#        popup2.attributes('-topmost', True)    
        
#        x0 = tk.StringVar()
#        x1 = tk.StringVar()
#        y0 = tk.StringVar()
#        y1 = tk.StringVar()
        
#        e1 = ttk.Entry(popup, textvariable = x0)
#        e2 = ttk.Entry(popup, textvariable = x1)
#        e3 = ttk.Entry(popup, textvariable = y0)
#        e4 = ttk.Entry(popup, textvariable = y1)
        
        self.e19 = ttk.Entry(popup4)
        self.r7 = ttk.Radiobutton(popup4, text = "上行递增", value = 0, variable = orientation_type4)
        self.r8 = ttk.Radiobutton(popup4, text = "下行递减", value = 1, variable = orientation_type4)
                
        ttk.Label(popup4, text = "最高限速").grid(row = 1,column = 2,
          padx = 5, pady = 5)
        self.e19.grid(row = 1,column = 3, padx = 5, pady = 5)
        self.r7.grid(row = 2,column = 2, padx = 5, pady = 5)
        self.r8.grid(row = 2,column = 3, padx = 5, pady = 5)
       
        ttk.Button(popup4, text = "运行", command = self.ConfigValue4).grid(row = 7,column = 2,
          padx = 5, pady = 5)
        ttk.Button(popup4, text = "关闭", command = popup4.destroy).grid(row = 7,column = 3,
          padx = 5, pady = 5)
        
    
    def ConfigValue(self):
        dwg_file = self.OpenEntry1.get()
        save_path = self.SaveEntry1.get()
#        with open("decline_transform.py", 'w') as file:
#            file.write(f'dwg_file = "{dwg_file}"')
        x_min = int(self.e1.get())
        x_max = int(self.e2.get())
        y_min = int(self.e3.get())
        y_max = int(self.e4.get())
        text_flag = int(text_type1.get())
        decline_st0 = float(self.e12.get())
#        with open("decline_transform.py", 'w') as file:
#            file.write(f'x_min = "{x_min}"')
#            file.write(f'x_max = "{x_max}"')
#            file.write(f'y_min = "{y_min}"')
#            file.write(f'y_max = "{y_max}"')
#        messagebox.showinfo(title = '提示', message = "配置完成")    
#        type_test.receive_data0(dwg_file)
#        type_test.receive_data1(x_min, x_max, y_min, y_max)   
        try:
            run = transform(dwg_file, save_path, x_min, x_max, y_min, y_max, text_flag, decline_st0)
        except Exception as e:
            messagebox.showerror(title = '提示', message = '坡度/坡长转化错误')
#        subprocess.run(['python', 'type_test.py'])
        if(run == 1):
            messagebox.showinfo(title = '提示', message = '坡度/坡长转化完成')
        else:
            messagebox.showwarning(title = '提示', message = '坡度/坡长转化失败')
            
    def ConfigValue2(self):
        dwg_file = self.OpenEntry2.get()
        save_path = self.SaveEntry2.get()
#        with open("decline_transform.py", 'w') as file:
#            file.write(f'dwg_file = "{dwg_file}"')
        x_min = int(self.e5.get())
        x_max = int(self.e6.get())
        y_min = int(self.e7.get())
        y_max = int(self.e8.get())
        Rx_min = int(self.e9.get())
        Rx_max = int(self.e10.get())
        Ry_min = int(self.e11.get())
        Ry_max = int(self.e19.get())
        RU_layer = self.e20.get()
        ST_layer = self.e21.get()
        sted_0 = int(self.e22.get())
        
        text_flag = int(text_type2.get())
        curve_flag = int(curve_type2.get())
        keywords_L = self.e23.get()
        keywords_R = self.e24.get()
#        with open("decline_transform.py", 'w') as file:
#            file.write(f'x_min = "{x_min}"')
#            file.write(f'x_max = "{x_max}"')
#            file.write(f'y_min = "{y_min}"')
#            file.write(f'y_max = "{y_max}"')
#        messagebox.showinfo(title = '提示', message = "配置完成")    
#        type_test.receive_data0(dwg_file)
#        type_test.receive_data1(x_min, x_max, y_min, y_max)   
#        if(all(x_min, x_max, y_min, y_max, Rx_min, Rx_max, Ry_min, Ry_max, RU_layer, 
#               ST_layer, sted_0, text_flag, curve_flag, keywords_L, keywords_R)): 
#            messagebox.showwarning(title = '提示', message = '请输入完整信息')
        try:
            run = ruler(dwg_file, save_path, x_min, x_max, y_min, y_max, text_flag,
                        curve_flag, keywords_L, keywords_R, Rx_min, Rx_max, Ry_min,
                        Ry_max, RU_layer, ST_layer, sted_0)
        except Exception as e:
            messagebox.showerror(title = '提示', message = '曲线转化错误')
#        subprocess.run(['python', 'type_test.py'])
        if(run == 1):
            messagebox.showinfo(title = '提示', message = '曲线转化完成')
        else:
            messagebox.showwarning(title = '提示', message = '曲线转化失败')

    def ConfigValue3(self):
        dwg_file = self.OpenEntry3.get()
        save_path = self.SaveEntry3.get()
#        with open("decline_transform.py", 'w') as file:
#            file.write(f'dwg_file = "{dwg_file}"')
        x_min = int(self.e13.get())
        x_max = int(self.e14.get())
        y_min = int(self.e15.get())
        y_max = int(self.e16.get())
        layer = self.e17.get()
        keywords_DK = self.e18.get()

        try:
            run = station(dwg_file, save_path, x_min, x_max, y_min, y_max, layer,
                      keywords_DK)
        except Exception as e:
            messagebox.showerror(title = '提示', message = '设施转化错误')
#        subprocess.run(['python', 'type_test.py'])
        if(run == 1):
            messagebox.showinfo(title = '提示', message = '设施转化完成')
        else:
            messagebox.showwarning(title = '提示', message = '设施转化失败，出错位置"%f"' % run)
            
    def ConfigValue4(self):
        xls_file = self.OpenEntry4.get()
        save_path = self.SaveEntry4.get()
#        with open("decline_transform.py", 'w') as file:
#            file.write(f'dwg_file = "{dwg_file}"')
        limitation = int(self.e19.get())
        orientation_flag = int(orientation_type4.get())
        try:
            run = speed_limit(xls_file, save_path, limitation, orientation_flag)
        except Exception as e:
            messagebox.showerror(title = '提示', message = '限速转化错误')
#        subprocess.run(['python', 'type_test.py'])
        if(run == 1):
            messagebox.showinfo(title = '提示', message = '限速转化完成')
        else:
            messagebox.showwarning(title = '提示', message = '限速转化失败')

def center_window(win):
    # 获取屏幕的宽度和高度
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    
    # 获取窗口的宽度和高度
    window_width = 450
    window_height = 440
    
    # 计算窗口左上角应该放置的x、y坐标
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    
    # 设置窗口的位置为居中
    win.geometry("{}x{}+{}+{}".format(window_width, window_height, x, y))
#    win.iconbitmap("favicon.ico")
    win.iconbitmap(get_path('favicon.ico'))
    
def get_path(ico_file):  #设置图标
    try:
        base_path = sys._MEIPASS
        ico_file = 'img\\' + ico_file
    except AttributeError:
        base_path = os.path.abspath('.')
    return os.path.normpath(os.path.join(base_path, ico_file))


if __name__ == '__main__':
    root = tk.Tk()
    root.title('线路条件转化工具')
    
    center_window(root)
    filename1 = tk.StringVar()
    filename2 = tk.StringVar()
    filename3 = tk.StringVar()
    filename4 = tk.StringVar()
    path_var1 = tk.StringVar()
    path_var2 = tk.StringVar()
    path_var3 = tk.StringVar()
    path_var4 = tk.StringVar()
    text_type1 = tk.IntVar()
#    text_type1.set(1)
    text_type2 = tk.IntVar()
    curve_type2 = tk.IntVar()
#    k_Length = tk.StringVar()
#    k_Radius = tk.StringVar()
    orientation_type4 = tk.IntVar()
    
    app = Application(master = root)
    app.mainloop()