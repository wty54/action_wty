# -*- coding: utf-8 -*-
"""
Created on Fri Jan 19 15:54:44 2024

@author: thinkpad
"""


import tkinter as tk
import tkinter.ttk as ttk
import os
import sys
from views import DrawPage
from views2 import WordPage
from views3 import ExcelPage

#from tkinter import messagebox

class MainPage:
    def __init__(self, master: tk.Tk):
        self.root = master
        self.root.title("线路仿真绘图工具")
        self.create_page()
#        self.create_widgets()
    def create_page(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True)
        self.frame1 = DrawPage(self.root)
        self.frame2 = WordPage(self.root)
        self.frame3 = ExcelPage(self.root)
        
#        self.frame1.pack(fill='both', expand=True)
#        self.frame2.pack(fill='both', expand=True)
#        self.frame3.pack(fill=tk.BOTH, expand=True)
        
        self.notebook.add(self.frame1, text="仿真结果绘图")
        self.notebook.add(self.frame2, text="生成word文档")
        self.notebook.insert("end", self.frame3, text="全程统计计算")
#        self.root["menu"] = menubar

#
#
#
#    def create_widgets(self):
#        #坡度坡长组件
#        #       打开文件路径按钮
#        group1 = ttk.Labelframe(self, text="坡度/坡长")
#        group1.grid(row = 1, column = 0, padx = 5, pady = 5)
#
#        self.OpenEntry1 = ttk.Entry(group1, textvariable = filename1)
#        self.OpenEntry1.grid(row = 1, column = 0, padx = 5, pady = 5)
#
#        self.btn01 = ttk.Button(group1)
#        self.btn01["text"] = "选择文件"
#        self.btn01.grid(row = 1, column = 1, padx = 5, pady = 5)
#        self.btn01["command"] = self.open_dxf1
##        filename.set(self.open_dxf)
##       保存文件路径按钮        
#        
#        self.SaveEntry1 = ttk.Entry(group1, textvariable = path_var1)
#        self.SaveEntry1.grid(row = 2, column = 0, padx = 5, pady = 5)
#        
#        self.btn02 = ttk.Button(group1)
#        self.btn02["text"] = "保存文件"
#        self.btn02.grid(row = 2, column = 1, padx = 5, pady = 5)
#        self.btn02["command"] = self.save_xls1
#        
##       配置按钮
#        self.btn04 = ttk.Button(group1)
#        self.btn04["text"] = "配置"
#        self.btn04.grid(row = 1, column = 2, padx = 5, pady = 5)
#        self.btn04["command"] = self.open_popup1
#        
#       
##坡度/坡长打开路径    
#    def open_dxf1(self):
#        try:
#            filetypes = [("DXF", ".dxf")]
#            filepath = filedialog.askopenfilename(title = '打开文件', 
#                      filetypes = filetypes, defaultextension = '.dxf', 
#                      initialdir = 'C:/Users/thinkpad/Desktop')
#            filename1.set(filepath)
#        except Exception as e:
#            messagebox.showwarning(title = '提示', message = '未选择任何文件')
#
#
##坡度/坡长保存路径    
#    def save_xls1(self):
#        try:
#            filetypes = [("XLSX", ".xlsx"), ("XLS", ".xls")]
#            filenewpath = filedialog.asksaveasfilename(title = '保存文件', 
#                      filetypes = filetypes, defaultextension = '.xlsx', 
#                      initialdir = 'C:/Users/thinkpad/Desktop' )
##            self.create_widgets.path_var = filenewpath
#            path_var1.set(filenewpath)
#        #保存文件
#
#        except Exception as e:
#            messagebox.showwarning(title = '提示', message = '未选择任何文件')
#
##坡度/坡长配置子窗口
#    def open_popup1(self):
#        popup = tk.Toplevel(self)
#        center_window(popup)
#        popup.geometry("300x251")
#        popup.title("坡长/坡度转换")
##        popup.attributes('-topmost', True)    
#        
##        x0 = tk.StringVar()
##        x1 = tk.StringVar()
##        y0 = tk.StringVar()
##        y1 = tk.StringVar()
#        
##        e1 = ttk.Entry(popup, textvariable = x0)
##        e2 = ttk.Entry(popup, textvariable = x1)
##        e3 = ttk.Entry(popup, textvariable = y0)
##        e4 = ttk.Entry(popup, textvariable = y1)
#        
#        self.e1 = ttk.Entry(popup)
#        self.e2 = ttk.Entry(popup)
#        self.e3 = ttk.Entry(popup)
#        self.e4 = ttk.Entry(popup)
#        self.e12 = ttk.Entry(popup)
#        
#        self.r1 = ttk.Radiobutton(popup, text = "单行文字", value = 0, variable = text_type1)
#        self.r2 = ttk.Radiobutton(popup, text = "多行文字", value = 1, variable = text_type1)
#        
#        ttk.Label(popup, text = "x坐标最小值").grid(row = 1,column = 2,
#          padx = 5, pady = 5)
#        self.e1.grid(row = 1,column = 3, padx = 5, pady = 5)
#        ttk.Label(popup, text = "x坐标最大值").grid(row = 2,column = 2,
#          padx = 5, pady = 5)
#        self.e2.grid(row = 2,column = 3, padx = 5, pady = 5)
#        ttk.Label(popup, text = "y坐标最小值").grid(row = 3,column = 2,
#          padx = 5, pady = 5)
#        self.e3.grid(row = 3,column = 3, padx = 5, pady = 5)
#        ttk.Label(popup, text = "y坐标最大值").grid(row = 4,column = 2,
#          padx = 5, pady = 5)
#        self.e4.grid(row = 4,column = 3, padx = 5, pady = 5)   
#        ttk.Label(popup, text = "起始公里标（km）").grid(row = 6,column = 2,
#          padx = 5, pady = 5)
#        self.e12.grid(row = 6,column = 3, padx = 5, pady = 5) 
#                
#        self.r1.grid(row = 5,column = 2, padx = 5, pady = 5)
#        self.r2.grid(row = 5,column = 3, padx = 5, pady = 5)
#    
#        ttk.Button(popup, text = "运行", command = self.ConfigValue).grid(row = 7,column = 2,
#          padx = 5, pady = 5)
#        ttk.Button(popup, text = "关闭", command = popup.destroy).grid(row = 7,column = 3,
#          padx = 5, pady = 5)
#
#    
#    def ConfigValue(self):
#        dwg_file = self.OpenEntry1.get()
#        save_path = self.SaveEntry1.get()
##        with open("decline_transform.py", 'w') as file:
##            file.write(f'dwg_file = "{dwg_file}"')
#        x_min = int(self.e1.get())
#        x_max = int(self.e2.get())
#        y_min = int(self.e3.get())
#        y_max = int(self.e4.get())
#        text_flag = int(text_type1.get())
#        decline_st0 = float(self.e12.get())
##        with open("decline_transform.py", 'w') as file:
##            file.write(f'x_min = "{x_min}"')
##            file.write(f'x_max = "{x_max}"')
##            file.write(f'y_min = "{y_min}"')
##            file.write(f'y_max = "{y_max}"')
##        messagebox.showinfo(title = '提示', message = "配置完成")    
##        type_test.receive_data0(dwg_file)
##        type_test.receive_data1(x_min, x_max, y_min, y_max)   
#        try:
#            run = transform(dwg_file, save_path, x_min, x_max, y_min, y_max, text_flag, decline_st0)
#        except Exception as e:
#            messagebox.showerror(title = '提示', message = '坡度/坡长转化错误')
##        subprocess.run(['python', 'type_test.py'])
#        if(run == 1):
#            messagebox.showinfo(title = '提示', message = '坡度/坡长转化完成')
#        else:
#            messagebox.showwarning(title = '提示', message = '坡度/坡长转化失败')
            
def center_window(win):
    # 获取屏幕的宽度和高度
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    
    # 获取窗口的宽度和高度
    window_width = 1520
    window_height = 600
    
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
    center_window(root)
    app = MainPage(root)
    root.mainloop()