# -*- coding: utf-8 -*-
"""
Created on Thu Dec 19 13:42:18 2024

@author: thinkpad
"""

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, filedialog
import os
from main import sorted_files
from exportFigure1 import ex_Fig
import re
#import traceback
import threading
#print(threading.current_thread() == threading.main_thread())
import queue
import gc
#import time
#from gui_test import MainPage

class DrawPage(ttk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.LoadFolder_path = tk.StringVar()   #输入明细地址
        self.Open_station = tk.StringVar()  #输入设施地址
        self.SavePIC_path = tk.StringVar()  #输入图片存放地址
        self.x_axis = tk.IntVar()   #输入x轴选项标志，0——里程，1——时间
        self.x_station = tk.IntVar()   #输入x轴每站或全程标志，0——每站，1——全程
        self.Ue = tk.StringVar()  #输入额定网压
        self.Ue.set('25000V')
        self.PIC_num = tk.IntVar()   #输入初始图片序号
        self.PIC_num.set(1)
        self.desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')    #设置初始地址为桌面
#        group1.grid(row = 1, column = 0, padx = 5, pady = 5)
        self.y_label = ["y轴", "速度(km/h)", "距离(km)", "时间(s)", "网侧电流(A)", "电机电流(A)", "阻力(kN)", "累计能耗(kWh)", "电机输出功率(kW)", "网侧输入功率(kW)", "加速度(m/s^2)"]
        self.colors = ['red', 'green', 'blue', 'yellow', 'cyan', 'magenta', 'black', 'chocolate', 'orange', 'gray']
        #  创建默认y轴范围
        self.yRange_default = ['[0, 200]', '[0, 10]', '[0, 600]', '[-500, 500]', '[-500, 500]', '[-500, 500]', '[0, 500]', '[-800, 800]', '[-8000, 8000]', '[-5, 5]']
        #  创建默认刻度格数
        self.scale_default = [10, 10, 10, 10, 10, 10, 10, 10, 10, 10]
#        self.colors_choose = tk.StringVar()
#        self.colors_choose.set("white")
        self.yaxis_Range = {}
        self.scale = {}
        self.check_vars = {}
        self.selected_colors = {}
#        self.PIC = []
        #  创建速度选项控件
        self.yaxis_labels = {}  #y轴标题
        self.checkbuttons = {}  #勾选框字典
        self.entries = {}  #输入框字典
        self.Combos = {}  #下拉菜单字典
        self.color_labels = {}  #颜色显示框字典
        #  创建输出字典
        self.x_axis_station_out = {}
        self.x_axis_type_out = {}
        self.yaxis_labels_out = {}
        self.yaxis_Range_out = {}
        self.scale_out = {}
        self.selected_colors_out = {}
        self.plotName_out = {}
        self.pic_count = 0  #图组计数器
        #  创建输入框输入类型
        for i in range(1, 11):
            self.yaxis_Range[f'yaxis_Range{i}'] = tk.StringVar()
            self.yaxis_Range[f'yaxis_Range{i}'].set(self.yRange_default[i-1])
            self.scale[f'scale{i}'] = tk.IntVar()
            self.scale[f'scale{i}'].set(self.scale_default[i-1])
            self.check_vars[f'check_vars{i}'] = tk.BooleanVar()
            self.selected_colors[f'selected_colors{i}'] = tk.StringVar()
            self.selected_colors[f'selected_colors{i}'].set(self.colors[i-1])
        self.is_running = False
        #  创建队列用于线程间通信
        self.queue = queue.Queue()
#        self.stop_event = threading.Event()
        self.after(100, self.process_queue)
        self.create_page()
        
    
    def create_page(self):
        #  创建路径获取控件
        self.group1 = ttk.Labelframe(self, text="设置明细打开路径和图片保存路径")
        self.group1.grid(row = 0, column = 0, padx = 10, pady = 10, rowspan=2)
        ttk.Label(self.group1, text="明细数据命名格式示例：明细_25000V_2M2T_AW3_上行").grid(row=0, column=0, padx=5, pady=5, columnspan=3)
        ttk.Label(self.group1, text="明细导入路径").grid(row=1, column=0, padx=5, pady=5)
        ttk.Entry(self.group1, textvariable = self.LoadFolder_path).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.group1, text="设置", command=self.open_folder).grid(row=1, column=2, padx=5, pady=5)
        
        ttk.Label(self.group1, text="设施导入路径").grid(row=2, column=0, padx=5, pady=5)
        ttk.Entry(self.group1, textvariable = self.Open_station).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(self.group1, text="设置", command=self.open_STA).grid(row=2, column=2, padx=5, pady=5)
        
        ttk.Label(self.group1, text="图片保存路径").grid(row=3, column=0, padx=5, pady=5)
        ttk.Entry(self.group1, textvariable = self.SavePIC_path).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(self.group1, text="设置", command=self.save_folder).grid(row=3, column=2, padx=5, pady=5)
         # 创建明细列表框
        self.group7 = ttk.Labelframe(self, text="选择明细文件")
        self.group7.grid(row = 2, column = 0, padx = 10, pady = 10)
        self.listbox2 = tk.Listbox(self.group7, width=50 , selectmode=tk.MULTIPLE)
        self.listbox2.grid(row=0, column=0, padx=10, pady=10, columnspan=2)
        sc = ttk.Scrollbar(self.group7, command=self.listbox2.yview)
        sc.grid(row=0, column=2, sticky='ns')
        self.listbox2.config(yscrollcommand=sc.set)

#        sc.grid(row=0, column=0)
        # 创建删除按钮并绑定delete_selected函数
        add_button2 = ttk.Button(self.group7, text="添加", command = self.add_dir)
        add_button2.grid(row=1, column=0, padx=10, pady=10)
        delete_button2 = ttk.Button(self.group7, text="删除", command=self.delete_dir)
        delete_button2.grid(row=1, column=1, padx=10, pady=10)
        
        #  创建x轴选项控件
        self.group2 = ttk.Labelframe(self, text="配置x轴")
        self.group2.grid(row=0, column=1, padx=10, pady=10)
        ttk.Radiobutton(self.group2, text="每站", value=0, variable=self.x_station).grid(row=0, column=0, padx=10, pady=10)
        ttk.Radiobutton(self.group2, text="全程", value=1, variable=self.x_station).grid(row=0, column=1, padx=10, pady=10)
        ttk.Radiobutton(self.group2, text="运行里程(km)", value=0, variable=self.x_axis).grid(row=1, column=0, padx=10, pady=10)
        ttk.Radiobutton(self.group2, text="运行时间(s)", value=1, variable=self.x_axis).grid(row=1, column=1, padx=10, pady=10)
        
        #  创建y轴选项控件
        self.group3 = ttk.Labelframe(self, text="配置y轴")
        self.group3.grid(row=1, column=1, padx=5, pady=5, rowspan=2)
        ttk.Label(self.group3, text = "选择").grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(self.group3, text = "显示范围([下限,上限])").grid(row=0, column=2, padx=10, pady=10)
        ttk.Label(self.group3, text = "刻度格数").grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(self.group3, text = "颜色").grid(row=0, column=4, padx=5, pady=5)

        for i in range(1, 11):
            self.yaxis_labels[f'yaxis_label{i}'] = ttk.Label(self.group3, text=self.y_label[i]).grid(row=i, column=0, padx=5, pady=5)
            self.Combos[f'Combo{i}'] = ttk.Combobox(self.group3, width=8, values=self.colors, textvariable=self.selected_colors[f'selected_colors{i}'], state="disabled")
            self.Combos[f'Combo{i}'].grid(row=i, column=4, padx=5, pady=5)
            self.color_labels[f'color_label{i}'] = ttk.Label(self.group3, width=3, state="disabled")
            self.color_labels[f'color_label{i}'].grid(row=i, column=5, padx=5, pady=5)

            self.Combos[f'Combo{i}'].bind("<<ComboboxSelected>>", self.change_color)
            self.checkbuttons[f'checkbutton{i}'] = ttk.Checkbutton(self.group3, offvalue=False, onvalue=True, variable=self.check_vars[f'check_vars{i}'], command=self.toggle_state)
            self.checkbuttons[f'checkbutton{i}'].grid(row=i, column=1, padx=5, pady=5)

            self.entries[f'entry{i}{2}'] = ttk.Entry(self.group3, width=14, textvariable=self.yaxis_Range[f'yaxis_Range{i}'], state="disabled")
            self.entries[f'entry{i}{2}'].grid(row=i, column=2, padx=5, pady=5)
            self.entries[f'entry{i}{3}'] = ttk.Entry(self.group3, width=3, textvariable=self.scale[f'scale{i}'], state="disabled")
            self.entries[f'entry{i}{3}'].grid(row=i, column=3, padx=5, pady=5)

        add_button1 = ttk.Button(self.group3, text="添加", command = self.add_info)
        add_button1.grid(row=11, column=2, padx=10, pady=10)
#        ttk.Button(self.group3, text="删除").grid(row=9, column=2, padx=5, pady=5)
#        ttk.Button(self.group3, text="运行").grid(row=9, column=3, padx=5, pady=5)
#        ttk.Button(self.group3, text = "关闭", command=self.on_exit).grid(row=9, column=4, padx=10, pady=10)
#    def on_exit(self):
#        if messagebox.askyesno("退出", "确定要退出吗？"):
#            self.destory()    
        #  创建绘图加载列表框和标准网压输入框
        self.group4 = ttk.Labelframe(self, text="绘图加载")
        self.group4.grid(row=2, column=2, padx=10, pady=10)
        # 创建列表框
        self.listbox = tk.Listbox(self.group4, width=80 , selectmode=tk.MULTIPLE)
        self.listbox.grid(row=0, column=0, padx=10, pady=10, columnspan=2)
        sc2 = ttk.Scrollbar(self.group4, command=self.listbox.yview)
        sc2.grid(row=0, column=2, sticky='ns')
        self.listbox.config(yscrollcommand=sc2.set)
        # 创建删除按钮并绑定delete_selected函数
        self.delete_button = ttk.Button(self.group4, text="删除", command=self.delete_selected)
#        delete_button.pack()
        self.delete_button.grid(row=1, column=0, padx=10, pady=10)
        self.run_button = ttk.Button(self.group4, text="运行", command=self.create_progressbar)
#        run_button.pack()
        self.run_button.grid(row=1, column=1, padx=10, pady=10)
        # 创建标准网压输入框
        self.group5 = ttk.Labelframe(self, text="输入额定网压")
        self.group5.grid(row=0, column=2, padx=10, pady=10)
        ttk.Label(self.group5, text="额定网压(例如:25000V)").grid(row=0, column=0, padx=10, pady=10)
        ttk.Entry(self.group5, textvariable = self.Ue).grid(row=0, column=1, padx=10, pady=10)
        #创建初始图片序号设置输入框
        self.group6 = ttk.Labelframe(self, text="设置初始图片序号")
        self.group6.grid(row=1, column=2, padx=10, pady=10)
        ttk.Label(self.group6, text="初始图片序号").grid(row=0, column=0, padx=10, pady=10)
        ttk.Entry(self.group6, textvariable = self.PIC_num).grid(row=0, column=1, padx=10, pady=10)
        
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
            self.SavePIC_path.set(folder_path)
            
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件夹')
    def open_STA(self):
        try:
            filetypes = [("xls", ".xls"), ("xlsx", ".xlsx")]
            filepath = filedialog.askopenfilename(title = '打开文件', 
                      filetypes = filetypes, defaultextension = '.xls')
            self.Open_station.set(filepath)
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')
    #  显示颜色
    def change_color(self, event):
        for i in range(1, 11):
            self.color_labels[f'color_label{i}'].config(background=self.selected_colors[f'selected_colors{i}'].get())

    def toggle_state(self):
    # 获取勾选框的状态
        for i in range(1, 11):
            if self.check_vars[f'check_vars{i}'].get():
                # 如果勾选框被选中，启用输入框和按钮
                self.entries[f'entry{i}{2}'].config(state='normal')
                self.entries[f'entry{i}{3}'].config(state='normal')
                self.Combos[f'Combo{i}'].config(state='normal')
            else:
                # 如果勾选框未被选中，禁用输入框和按钮
                self.entries[f'entry{i}{2}'].config(state='disabled')
                self.entries[f'entry{i}{3}'].config(state='disabled')
                self.Combos[f'Combo{i}'].config(state='disabled')
#                self.entries[f'entry{i}{2}'].delete(0, tk.END)
#                self.entries[f'entry{i}{3}'].delete(0, tk.END)
                # 如果勾选框未被选中，清空颜色框
                self.Combos[f'Combo{i}'].set('')
                self.color_labels[f'color_label{i}'].config(background='')

    def add_info(self):
        #  添加绘图配置
        yaxis_names_temp = []
        plot_lims_temp = []
        scale_lists_temp = []
        color_lists_temp = []
        xaxis_station_temp = self.x_station.get()
        xaxis_type_temp = self.x_axis.get()
        for i in range(1, 11):
            if self.check_vars[f'check_vars{i}'].get():
                yaxis_names_temp.append(self.y_label[i])
#                plot_lims_temp.append(self.entries[f'entry{i}{2}'].get())
                plot_lims_temp.append(self.yaxis_Range[f'yaxis_Range{i}'].get())
#                scale_lists_temp.append(int(self.entries[f'entry{i}{3}'].get()))
                scale_lists_temp.append(self.scale[f'scale{i}'].get())
#                color_lists_temp.append(self.Combos[f'Combo{i}'].get())
                color_lists_temp.append(self.selected_colors[f'selected_colors{i}'].get())
            else:
                continue
        if yaxis_names_temp:
            results = [(re.sub(r'[^\u4e00-\u9fa5]+', '', name)) for name in yaxis_names_temp]
            result = "、".join(results)
            self.x_axis_station_out[f'group{self.pic_count}'] = xaxis_station_temp
            self.x_axis_type_out[f'group{self.pic_count}'] = xaxis_type_temp
            self.yaxis_labels_out[f'group{self.pic_count}'] = yaxis_names_temp
            self.yaxis_Range_out[f'group{self.pic_count}'] = plot_lims_temp
            self.scale_out[f'group{self.pic_count}'] = scale_lists_temp
            self.selected_colors_out[f'group{self.pic_count}'] = color_lists_temp
            if xaxis_station_temp == 0:
                if xaxis_type_temp == 0:
                    self.plotName_out[f'group{self.pic_count}'] = result + "对运行里程曲线"
                    self.listbox.insert(tk.END, "第" + f'{self.pic_count + 1}' + "组图为：" + "每站_" + self.plotName_out[f'group{self.pic_count}'])
    #                print("第" + f'{self.pic_count}' + "组图为：\n" + result + "对运行里程曲线")
                else:
                    self.plotName_out[f'group{self.pic_count}'] = result + "对运行时间曲线"
                    self.listbox.insert(tk.END, "第" + f'{self.pic_count + 1}' + "组图为：" + "每站_" + self.plotName_out[f'group{self.pic_count}'])
#                print("第" + f'{self.pic_count}' + "组图为：\n" + result + "对运行时间曲线")
            else:
                if xaxis_type_temp == 0:
                    self.plotName_out[f'group{self.pic_count}'] = result + "对运行里程曲线"
                    self.listbox.insert(tk.END, "第" + f'{self.pic_count + 1}' + "组图为：" + "全程_" + self.plotName_out[f'group{self.pic_count}'])
    #                print("第" + f'{self.pic_count}' + "组图为：\n" + result + "对运行里程曲线")
                else:
                    self.plotName_out[f'group{self.pic_count}'] = result + "对运行时间曲线"
                    self.listbox.insert(tk.END, "第" + f'{self.pic_count + 1}' + "组图为：" + "全程_" + self.plotName_out[f'group{self.pic_count}'])
            self.pic_count += 1
        else:
            messagebox.showwarning(title = '提示', message = '未选择任何绘图参数')
    def delete_selected(self):
        # 获取选中的项的索引列表
        selected_indices = self.listbox.curselection()
        # 从后往前删除选中的项，避免因索引改变导致的错误
        for index in sorted(selected_indices, reverse=True):
            self.x_axis_station_out =  self.delete_key(self.x_axis_station_out, index)
            self.x_axis_type_out =  self.delete_key(self.x_axis_type_out, index)
            self.yaxis_labels_out = self.delete_key(self.yaxis_labels_out, index)
            self.yaxis_Range_out = self.delete_key(self.yaxis_Range_out, index)
            self.scale_out = self.delete_key(self.scale_out, index)
            self.selected_colors_out = self.delete_key(self.selected_colors_out, index)
            self.plotName_out = self.delete_key(self.plotName_out, index)
#            print(self.x_axis_type_out)
#            print(self.yaxis_labels_out)                
#            print(self.yaxis_Range_out)
#            print(self.scale_out)
#            print(self.selected_colors_out)
#            print(self.plotName_out)
            self.listbox.delete(index)
            self.pic_count -= 1
    # 删除并重建字典序列
    def delete_key(self, input_dict, index):
        output_dict = {}
        del input_dict[f'group{index}']
        for i in range(len(input_dict)):
            output_dict[f'group{i}'] = list(input_dict.items())[i][1]
        return output_dict
    def add_dir(self):
        # 让用户选择文件夹中的.xls或.xlsx文件
        file_path = filedialog.askopenfilenames(initialdir=self.LoadFolder_path, title="选择明细文件", filetypes=(("Excel files", "*.xls *.xlsx"),))
        if file_path:  # 如果用户选择了文件
            for fp in file_path:
                fp = os.path.basename(fp)
                self.listbox2.insert(tk.END, fp)
    
    def delete_dir(self):
        # 获取选中的项的索引列表
        selected_indices2 = self.listbox2.curselection()
        for index in sorted(selected_indices2, reverse=True):
            self.listbox2.delete(index)
       
#            print(Load)
#            print(Open)
#            print(Save)
#            print(lineVoltage)
#            print(self.x_axis_type_out)
#            print(self.yaxis_labels_out)                
#            print(self.yaxis_Range_out)
#            print(self.scale_out)
#            print(self.selected_colors_out)
#            print(self.plotName_out)
    # 运行绘图程序
    def create_progressbar(self):
        # 更新按钮状态
        self.run_button.config(state=tk.DISABLED)        
        #创建进度条组件
        self.top = tk.Toplevel(self)
        self.top.title("任务进度")
        self.top.geometry("300x150")
        
        
#        self.task_length = 100
        self.progresslabel = tk.Label(self.top, text="进度:")
        self.progresslabel.pack(pady=10)
        # 进度条
#        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.top, 
            maximum=100,
            mode="determinate"
        )
        self.progress_bar.pack(pady=20, padx=20, fill=tk.X)
     
       
        # 停止按钮
        self.stop_button = tk.Button(
            self.top, 
            text="停止任务", 
            command=self.stop_task,
            state=tk.NORMAL
        )
        self.stop_button.pack(pady=10)

        # 重置停止事件
#        self.stop_event.clear()
        
        # 重置进度条
        self.progress_bar['value'] = 0

        # 启动任务线程
        self.thread = threading.Thread(target=self.run_program, daemon=True)
        self.thread.start()
        # 窗口关闭时的事件处理
        self.top.protocol("WM_DELETE_WINDOW", self.stop_task)
    def run_program(self):
        Load = self.LoadFolder_path.get()
        Open = self.Open_station.get()
        Save = self.SavePIC_path.get()
        lineVoltage = self.Ue.get()
        PIC_NO = self.PIC_num.get()
        file_names = []
        # 获取明细表中的所有项
        selected_files = self.listbox2.get(0, tk.END)
     
   # 将获取到的内容添加到数组中
        for item in selected_files:
            file_names.append(item)
        
        file_names = sorted_files(file_names, lineVoltage, '明细')
        self.task_length = len(file_names)  #设定总进度条长度
   #线程控制器使能
        self.is_running = True
#        try:
#            for i in range(1, 6):
#                # 检查是否收到停止信号
##                if self.stop_event.is_set():
#                if not self.is_running:
#                    break
#                p = i * 100 // 5
#                # 模拟工作
#                time.sleep(1)
##                print(i)
#                # 更新进度
#                self.queue.put(('progress', p))
#            
#            # 任务完成
##            if not self.stop_event.is_set():
#            if self.is_running:
#                self.queue.put(('status', "任务完成!"))
#                self.reset_buttons()
#                self.after(100, self.top.destroy)
#        except Exception as e:
#            self.queue.put(('error', str(e)))
#            self.after(100, self.top.destroy)
        
        try:
            for fff in range(len(file_names)):
                if not self.is_running:
                    break
                percent = (fff + 1) * 100 // len(file_names)
#                self.queue.put(('progress', percent))
                PIC_series = ex_Fig(file_names[fff], Load, Open, Save, PIC_NO, self.x_axis_station_out, self.x_axis_type_out, self.yaxis_labels_out, self.yaxis_Range_out, self.scale_out, self.selected_colors_out, self.plotName_out) 
                PIC_NO = PIC_series
                # 更新进度
                self.queue.put(('progress', percent))
                if fff % 10 ==0:
                    gc.collect()
            
            if not self.plotName_out:
                messagebox.showwarning(title = '提示', message = '未选择绘图项目')
            # 任务完成
#            if not self.stop_event.is_set():
            if self.is_running:
                self.queue.put(('status', "任务完成!"))
                self.reset_buttons()
                self.after(100, self.top.destroy)
        except Exception as e:
            self.queue.put(('error', str(e)))
            self.after(100, self.top.destroy)
#            except Exception as e:            
#                self.queue.put(('error', str(e)))
##                traceback.print_exc()
##                print(f"发生了一个错误：{e}")
##                return 0
            
#        
#        if not self.plotName_out:
#            messagebox.showwarning(title = '提示', message = '未选择绘图项目')
#        else:
##            messagebox.showinfo(title = '提示', message = '绘图完成')
#            # 任务完成后关闭窗口
#            self.queue.put(('status', "绘图完成!"))
    def stop_task(self):
        """停止任务"""
#        self.stop_event.set()
        self.is_running = False
        self.queue.put(('status', "任务已停止"))
        self.reset_buttons()
        self.after(100, self.top.destroy)
#        self.stop_event.clear()        
    def process_queue(self):
        """处理来自工作线程的消息"""
        try:
            while self.is_running:
                msg = self.queue.get_nowait()
#                print(msg[1])
                if msg[0] == 'progress':
                    self.progress_bar['value'] = msg[1]
                    self.progresslabel.config(text=f"进度: {msg[1]}%")
                elif msg[0] == 'status':
                    messagebox.showinfo(title = '提示', message = msg[1])
#                    self.top.destroy()
                    self.reset_buttons()
                
                elif msg[0] == 'error':
                    messagebox.showerror(title = '提示', message = f'发生了一个错误：{msg[1]}，绘图失败')
#                    self.top.destroy()
                    self.reset_buttons()
            msg = self.queue.get_nowait()
            if msg[0] == 'status':
                messagebox.showinfo(title = '提示', message = msg[1])
                
#            print(msg[1])
        except queue.Empty:
            pass
        
        # 继续定期检查队列
        self.after(100, self.process_queue)
    
    def reset_buttons(self):
        """重置按钮状态"""
        self.run_button.config(state=tk.NORMAL)
#        self.stop_button.config(state=tk.DISABLED)
    
#    def on_close(self):
#        """窗口关闭事件处理"""
#        self.after(100, self.top.destroy)
#        self.stop_task()
        
#        self.top.destroy()
#        print(self.queue.get_nowait())
    
            
            