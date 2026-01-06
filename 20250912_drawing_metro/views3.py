# -*- coding: utf-8 -*-
"""
Created on Sun Apr 27 09:00:06 2025

@author: thinkpad
"""

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, filedialog
#from tkinter import scrolledtext
import os
from main import sorted_files
from total import cal_total



class ExcelPage(ttk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.LoadFolder_path = tk.StringVar()   #打开统计数据地址
#        self.Open_PIC = tk.StringVar()  #输入设施地址
        self.SaveXLS_path = tk.StringVar()  #生成全部统计数据表地址
        self.kilo_info = tk.StringVar()
        self.stop_info = tk.IntVar()
        self.click_first = 0
        self.click_second = 0
        self.click_count = 0
        self.add_info = tk.BooleanVar()  #是否添加文字说明
        self.gap_info = tk.IntVar()  #间隔图片数量
        self.gap_info.set(1)
        self.create_widget()
                
    def create_widget(self):
        self.group1 = ttk.Labelframe(self, text="设置统计数据打开路径和全统计表保存路径")
        self.group1.grid(row = 0, column = 0, padx = 10, pady = 10)
        ttk.Label(self.group1, text="统计数据命名格式示例：统计_1500V_4M2T_AW3_上行").grid(row=0, column=0, padx=10, pady=10, columnspan=3)
        ttk.Label(self.group1, text="统计数据打开路径").grid(row=1, column=0, padx=10, pady=10)
        ttk.Entry(self.group1, textvariable = self.LoadFolder_path).grid(row=1, column=1, padx=10, pady=10)
        ttk.Button(self.group1, text="设置", command=self.open_folder).grid(row=1, column=2, padx=10, pady=10)
        
        ttk.Label(self.group1, text="全统计表保存路径").grid(row=2, column=0, padx=10, pady=10)
        ttk.Entry(self.group1, textvariable = self.SaveXLS_path).grid(row=2, column=1, padx=10, pady=10)
        ttk.Button(self.group1, text="设置", command=self.save_folder).grid(row=2, column=2, padx=10, pady=10)
        ttk.Label(self.group1, text="总里程公里数(km)").grid(row=3, column=0, padx=10, pady=10)
        self.kilo_Entry = ttk.Entry(self.group1, width=8, textvariable=self.kilo_info)
        self.kilo_Entry.grid(row=3, column=1, padx=10, pady=10)
        
        ttk.Label(self.group1, text="折返时间(s)").grid(row=4, column=0, padx=10, pady=10)
        self.stop_Entry = ttk.Entry(self.group1, width=8, textvariable=self.stop_info)
        self.stop_Entry.grid(row=4, column=1, padx=10, pady=10)
        
        self.listbox2 = tk.Listbox(self.group1, width=50 , selectmode=tk.MULTIPLE)
        self.listbox2.grid(row=5, column=0, padx=10, pady=10, columnspan=3)
        sc = ttk.Scrollbar(self.group1, command=self.listbox2.yview)
        sc.grid(row=5, column=4, sticky='ns')
        sc2 = ttk.Scrollbar(self.group1, orient='horizontal', command=self.listbox2.xview)
        sc2.grid(row=6, column=0, sticky='we', columnspan=3)
        # 定义右键多选功能
        self.listbox2.config(yscrollcommand=sc.set, xscrollcommand=sc2.set)
        self.listbox2.bind("<Button-3>", self.on_right_click)
#        sc.grid(row=0, column=0)
        # 创建删除按钮并绑定delete_selected函数
        add_button2 = ttk.Button(self.group1, text="添加", command = self.add_dir)
        add_button2.grid(row=7, column=0, padx=10, pady=10)
        delete_button2 = ttk.Button(self.group1, text="删除", command=self.delete_dir)
        delete_button2.grid(row=7, column=1, padx=10, pady=10)
        self.run_button2 = ttk.Button(self.group1, text="运行", command=self.run_program)
        self.run_button2.grid(row=7, column=2, padx=10, pady=10)
        # 自定义信息输入
#        self.group2 = ttk.Labelframe(self, text="设置自定义信息")
#        self.group2.grid(row = 0, column = 1, padx = 10, pady = 10)
#        ttk.Label(self.group2, text="是否增加自定义信息").grid(row=0, column=0, padx=10, pady=10, sticky="w")
#        self.check_add = ttk.Checkbutton(self.group2, offvalue=False, onvalue=True, variable=self.add_info, command=self.toggle_state)
#        self.check_add.grid(row=0, column=1, padx=10, pady=10)
#        ttk.Label(self.group2, text = "间隔多少张图片写一段描述").grid(row=1, column=0, padx=10, pady=10, sticky="w")
#        self.gap_entry = ttk.Entry(self.group2, width=3, textvariable=self.gap_info, state="disabled")
#        self.gap_entry.grid(row=1, column=1, padx=10, pady=10)
#        ttk.Label(self.group2, text = "文字描述内容（注：先检查图片命名格式是否为“序号 起点——终点 网压_编组_载荷_行车方向_曲线名称”）").grid(row=2, column=0, padx=10, pady=10)
#        self.information_entry = scrolledtext.ScrolledText(self.group2, wrap=tk.WORD, width=80, height=10)
#        self.information_entry.grid(row=3, column=0, padx=10, pady=10, columnspan=2, sticky="w")
#        self.information_entry.config(state="disabled")
        
    def open_folder(self):
        try:
            # 打开文件夹选择对话框，并获取选择的文件夹路径
            folder_path = filedialog.askdirectory() #(initialdir = self.desktop_path)
            self.LoadFolder_path.set(folder_path)
            
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件夹')
    def save_folder(self):
        try:
            # 打开文件夹选择对话框，并获取选择的文件夹路径
            folder_path = filedialog.askdirectory() #(initialdir = self.desktop_path)
            self.SaveXLS_path.set(folder_path)
            
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件夹')
        
    def add_dir(self):
        # 选择文件夹中的.jpg或.jpeg文件
        file_path = filedialog.askopenfilenames(initialdir=self.LoadFolder_path, title="选择数据", filetypes=(("Excel files", "*.xls *.xlsx"),))
        if file_path:  # 如果用户选择了文件
            for fp in file_path:
                fp = os.path.basename(fp)
                self.listbox2.insert(tk.END, fp)
#            print(type(self.listbox2.get(0, tk.END)))
    def delete_dir(self):
        # 获取选中的项的索引列表
        selected_indices2 = self.listbox2.curselection()
        for index in sorted(selected_indices2, reverse=True):
            self.listbox2.delete(index)
            
    def on_right_click(self, event):
        # 获取当前鼠标点击的条目索引
        if self.click_count == 0:
            self.listbox2.selection_clear(0, tk.END)
            self.click_first = self.listbox2.nearest(event.y)
            self.listbox2.selection_set(self.click_first)
            self.click_count += 1
        else:
            self.click_second = self.listbox2.nearest(event.y)
            for i in range(min(self.click_first, self.click_second), max(self.click_first, self.click_second) + 1):
                self.listbox2.selection_set(i)
            self.click_count = 0

    def run_program(self):
        # 获取统计数据打开地址和全程统计存储地址
        Load = self.LoadFolder_path.get()
        Save = self.SaveXLS_path.get()
        Kilo = float(self.kilo_info.get())
        Stop = self.stop_info.get()
        # 获取列表中的所有数据名称
        selected_files = list(self.listbox2.get(0, tk.END))
        selected_files = sorted_files(selected_files, "1500V", '统计')
#        cal_total(Load, Save, Kilo, Stop, selected_files)
#        messagebox.showinfo(title = '提示', message = '全程数据计算完成')
        mess,error_file = cal_total(Load, Save, Kilo, Stop, selected_files)
        if mess == 1:
            messagebox.showerror(title = '提示', message = f"统计数据计算错误，错误文件：{error_file}")
        else:
            messagebox.showinfo(title = '提示', message = f"统计数据计算完成")



