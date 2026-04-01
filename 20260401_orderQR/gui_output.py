# -*- coding: utf-8 -*-
"""
Created on Thu Mar 26 10:51:37 2026

@author: thinkpad
"""
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
from tkinter import scrolledtext
import os
import sys
from label_word import start, get_resource_path
import traceback

class MainPage(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("采购订单信息生成工具")
        self.LoadExcel_path = tk.StringVar()  # 导入表格路径输入框
        self.save_path1 = tk.StringVar()  # 唛头文件保存路径输入框
        self.save_path2 = tk.StringVar()  # 标签文件保存路径输入框
        self.QR_size1 = tk.StringVar()  # 唛头二维码尺寸
        self.QR_size1.set(3)
        self.QR_size2 = tk.StringVar()  # 标签二维码尺寸
        self.QR_size2.set(3.8)
        self.radio1 = tk.IntVar()  # 0_缺省唛头模板  1_自定义唛头模板
        self.radio2 = tk.IntVar()  # 0_缺省标签模板  1_自定义标签模板
        self.model1 = tk.StringVar()  # 自定义唛头模板路径
        self.model2 = tk.StringVar()  # 自定义标签模板路径
#        self.model_excel_path = "C:\\Users\\thinkpad\\Desktop\\test\\发货唛头信息模板.xlsx"  # 需要替换
        self.model_excel_path = get_resource_path(os.path.join("template", "发货信息模板.xlsx"))
        self.model1_path = get_resource_path(os.path.join("template", "唛头模板.docx"))
        self.model2_path = get_resource_path(os.path.join("template", "标签模板.docx"))
        
        self.QR_list = ["项目名称", "采购订单", "物资编号", "规格型号", "物料名称", "计量单位", "数量", "序列号", "生产日期", "保质期", "列数", "供应商代码", "供应商公司", "备用1", "备用2", "备用3", "备用4", "备用5"]
        self.create_menu()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.create_page()
    
    def create_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="退出", command=self.ask_quit)
    
    def ask_quit(self):
        if messagebox.askyesno("确认退出", "你真的要退出吗？"):
            self.destroy()
    def on_closing(self):
        if messagebox.askyesno("确认退出", "你真的要退出吗？"):
            self.destroy()
    def create_page(self):
        self.group1 = ttk.Labelframe(self, text="导入发货信息统计表&模板确认")
        self.group1.grid(row=0, column=0, padx=10, pady=10)
        group1_label1 = ttk.Label(self.group1, text="导入唛头信息表格路径")
        group1_label1.grid(row=0, column=0, padx=10, pady=10)
        group1_Entry1 = ttk.Entry(self.group1, textvariable = self.LoadExcel_path)
        group1_Entry1.grid(row=0, column=1, padx=10, pady=10)
        group1_button1 = ttk.Button(self.group1, text="打开", command=self.open_excel)
        group1_button1.grid(row=0, column=2, padx=10, pady=10)
        group1_button2 = ttk.Button(self.group1, text="下载表格模板", command=self.download_excel)
        group1_button2.grid(row=0, column=3, padx=10, pady=10)
        group1_label2 = ttk.Label(self.group1, text="导入唛头模板")
        group1_label2.grid(row=1, column=0, padx=10, pady=10)
        self.group1_radio1 = ttk.Radiobutton(self.group1, text = "使用缺省模板", value=0, variable=self.radio1, command=self.toggle_state1)
        self.group1_radio1.grid(row=1, column=1, padx=10, pady=10)
        self.group1_radio2 = ttk.Radiobutton(self.group1, text = "使用自定义模板", value=1, variable=self.radio1, command=self.toggle_state1)
        self.group1_radio2.grid(row=1, column=2, padx=10, pady=10)
        group1_label2 = ttk.Label(self.group1, text="导入唛头模板路径")
        group1_label2.grid(row=2, column=0, padx=10, pady=10)
        self.group1_Entry2 = ttk.Entry(self.group1, textvariable = self.model1, state="disabled")
        self.group1_Entry2.grid(row=2, column=1, padx=10, pady=10)
        self.group1_button3 = ttk.Button(self.group1, text="打开", command=self.open_model1, state="disabled")
        self.group1_button3.grid(row=2, column=2, padx=10, pady=10)
        self.group1_button4 = ttk.Button(self.group1, text="下载唛头模板", command=self.download_model1, state="disabled")
        self.group1_button4.grid(row=2, column=3, padx=10, pady=10)
        group1_label3 = ttk.Label(self.group1, text="导入标签模板")
        group1_label3.grid(row=3, column=0, padx=10, pady=10)
        self.group1_radio3 = ttk.Radiobutton(self.group1, text = "使用缺省模板", value=0, variable=self.radio2, command=self.toggle_state2)
        self.group1_radio3.grid(row=3, column=1, padx=10, pady=10)
        self.group1_radio4 = ttk.Radiobutton(self.group1, text = "使用自定义模板", value=1, variable=self.radio2, command=self.toggle_state2)
        self.group1_radio4.grid(row=3, column=2, padx=10, pady=10)
        group1_label4 = ttk.Label(self.group1, text="导入标签模板路径")
        group1_label4.grid(row=4, column=0, padx=10, pady=10)
        self.group1_Entry3 = ttk.Entry(self.group1, textvariable = self.model2, state="disabled")
        self.group1_Entry3.grid(row=4, column=1, padx=10, pady=10)
        self.group1_button5 = ttk.Button(self.group1, text="打开", command=self.open_model2, state="disabled")
        self.group1_button5.grid(row=4, column=2, padx=10, pady=10)
        self.group1_button6 = ttk.Button(self.group1, text="下载标签模板", command=self.download_model2, state="disabled")
        self.group1_button6.grid(row=4, column=3, padx=10, pady=10)
        
        
        self.group2 = ttk.Labelframe(self, text="二维码信息录入")
        self.group2.grid(row=0, column=1, padx=10, pady=10)
        self.group2_listbox1 = tk.Listbox(self.group2, width=30, height=13, selectmode=tk.MULTIPLE)
        self.group2_listbox1.grid(row=0, column=0, padx=10, pady=10, rowspan=4, sticky='nsew')
        listbox1_scollbar = ttk.Scrollbar(self.group2, orient="vertical", command=self.group2_listbox1.yview)
        self.group2_listbox1.config(yscrollcommand=listbox1_scollbar.set)
        listbox1_scollbar.grid(row=0, column=1, padx=0, pady=0, rowspan=4, sticky='ns')
        for item in self.QR_list:
            self.group2_listbox1.insert(tk.END, item)
        group2_button1 = ttk.Button(self.group2, text=">>", command=self.move_list2, width=3)
        group2_button1.grid(row=0, column=2, padx=5, pady=5, rowspan=4)
        self.group2_listbox2 = tk.Listbox(self.group2, width=30, height=5 ,selectmode=tk.MULTIPLE)
        self.group2_listbox2.grid(row=0, column=3, padx=10, pady=10, sticky='nsew')
        listbox2_scollbar = ttk.Scrollbar(self.group2, orient="vertical", command=self.group2_listbox2.yview)
        self.group2_listbox2.config(yscrollcommand=listbox2_scollbar.set)
        listbox2_scollbar.grid(row=0, column=4, padx=0, pady=0, sticky='ns')
        group2_button3 = ttk.Button(self.group2, text="上移", command=self.up_list2)
        group2_button3.grid(row=1, column=3, padx=5, pady=5, columnspan=2)
        group2_button4 = ttk.Button(self.group2, text="下移", command=self.down_list2)
        group2_button4.grid(row=2, column=3, padx=5, pady=5, columnspan=2)
        group2_button5 = ttk.Button(self.group2, text="删除", command=self.remove_list2)
        group2_button5.grid(row=3, column=3, padx=5, pady=5, columnspan=2)
        
        
        self.group3 = ttk.Labelframe(self, text="制作唛头")
        self.group3.grid(row=1, column=0, padx=10, pady=10)
#        group3_label1 = ttk.Label(self.group3, text="编辑二维码信息")
#        group3_label1.grid(row=4, column=0, padx=5, pady=5)
        self.group3_scrolledtext = scrolledtext.ScrolledText(self.group3, wrap=tk.WORD, width=30, height=5)
        self.group3_scrolledtext.grid(row=0, column=0, padx=5, pady=5)
        group3_button6 = ttk.Button(self.group3, text="二维码信息生成", command=self.generate_QRinfo1)
        group3_button6.grid(row=0, column=1, padx=5, pady=5)
        group3_label2 = ttk.Label(self.group3, text="二维码尺寸")
        group3_label2.grid(row=1, column=0, padx=5, pady=5)
        group3_Entry2 = ttk.Entry(self.group3, textvariable = self.QR_size1, width=5)
        group3_Entry2.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        group3_label1 = ttk.Label(self.group3, text="唛头文件保存路径")
        group3_label1.grid(row=2, column=0, padx=5, pady=5)
        group3_Entry1 = ttk.Entry(self.group3, textvariable = self.save_path1)
        group3_Entry1.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        group3_button1 = ttk.Button(self.group3, text="选择", command=self.save1)
        group3_button1.grid(row=2, column=2, padx=5, pady=5, sticky='e')
        
        run_button1 = ttk.Button(self, text="制作唛头", command=self.run1)
        run_button1.grid(row=2, column=0, padx=10, pady=10)
        
        self.group4 = ttk.Labelframe(self, text="制作标签")
        self.group4.grid(row=1, column=1, padx=10, pady=10)
#        group3_label1 = ttk.Label(self.group3, text="编辑二维码信息")
#        group3_label1.grid(row=4, column=0, padx=5, pady=5)
        self.group4_scrolledtext = scrolledtext.ScrolledText(self.group4, wrap=tk.WORD, width=30, height=5)
        self.group4_scrolledtext.grid(row=0, column=0, padx=5, pady=5)
        group4_button6 = ttk.Button(self.group4, text="二维码信息生成", command=self.generate_QRinfo2)
        group4_button6.grid(row=0, column=1, padx=5, pady=5)
        group4_label2 = ttk.Label(self.group4, text="二维码尺寸")
        group4_label2.grid(row=1, column=0, padx=5, pady=5)
        group4_Entry2 = ttk.Entry(self.group4, textvariable = self.QR_size2, width=5)
        group4_Entry2.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        group4_label1 = ttk.Label(self.group4, text="标签文件保存路径")
        group4_label1.grid(row=2, column=0, padx=5, pady=5)
        group4_Entry1 = ttk.Entry(self.group4, textvariable = self.save_path2)
        group4_Entry1.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        group4_button1 = ttk.Button(self.group4, text="选择", command=self.save2)
        group4_button1.grid(row=2, column=2, padx=5, pady=5, sticky='e')
        
        run_button2 = ttk.Button(self, text="制作标签", command=self.run2)
        run_button2.grid(row=2, column=1, padx=10, pady=10)
        
    def open_excel(self):
        try:
            filetypes = [("Excel工作簿(*.xlsx)", ".xlsx"), ("Excel97-2003工作簿(*.xls)", ".xls")]
            filepath = filedialog.askopenfilename(title = '打开文件', 
                      filetypes = filetypes, defaultextension = '.xlsx', 
                      initialdir = 'C:/Users/thinkpad/Desktop')
            self.LoadExcel_path.set(filepath)
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未打开任何文件')
            
    def toggle_state1(self):
        if self.radio1.get():
            self.group1_Entry2.config(state='normal')
            self.group1_button3.config(state='normal')
            self.group1_button4.config(state='normal')
        else:
            self.group1_Entry2.delete(0, tk.END)
            self.group1_Entry2.config(state='disabled')
            self.group1_button3.config(state='disabled')
            self.group1_button4.config(state='disabled')
            self.model1.set("")
    def open_model1(self):
        try:
            filetypes = [("Word文档(*.docx)", ".docx"), ("Word97-2003文档(*.doc)", ".doc")]
            filepath = filedialog.askopenfilename(title = '打开文件', 
                      filetypes = filetypes, defaultextension = '.docx')
            self.model1.set(filepath)
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未打开任何文件')
    def download_model1(self):
        try:
            filetypes = [("Word文档(*.docx)", ".docx"), ("Word97-2003文档(*.doc)", ".doc")]
            filenewpath = filedialog.asksaveasfilename(title = '保存模板文件', 
                      filetypes = filetypes, defaultextension = '.docx', 
                      initialdir = 'C:/Users/thinkpad/Desktop',
                      initialfile = '唛头模板')
#            self.create_widgets.path_var = filenewpath
            if filenewpath:
                import shutil
                shutil.copy2(self.model1_path, filenewpath)
                self.model1.set(filenewpath)
                messagebox.showinfo("提示", f"模板文件已成功保存到:\n{filenewpath}")
            else:
                messagebox.showinfo("提示", f"下载已取消")
        #保存文件
        except Exception as e:
            messagebox.showwarning(title = '提示', message = f'未保存任何文件，错误信息{e}')
    
    def toggle_state2(self):
        if self.radio2.get():
            self.group1_Entry3.config(state='normal')
            self.group1_button5.config(state='normal')
            self.group1_button6.config(state='normal')
        else:
            self.group1_Entry3.delete(0, tk.END)
            self.group1_Entry3.config(state='disabled')
            self.group1_button5.config(state='disabled')
            self.group1_button6.config(state='disabled')
            self.model2.set("")
    def open_model2(self):
        try:
            filetypes = [("Word文档(*.docx)", ".docx"), ("Word97-2003文档(*.doc)", ".doc")]
            filepath = filedialog.askopenfilename(title = '打开文件', 
                      filetypes = filetypes, defaultextension = '.docx')
            self.model2.set(filepath)
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未打开任何文件')
    def download_model2(self):
        try:
            filetypes = [("Word文档(*.docx)", ".docx"), ("Word97-2003文档(*.doc)", ".doc")]
            filenewpath = filedialog.asksaveasfilename(title = '保存模板文件', 
                      filetypes = filetypes, defaultextension = '.docx', 
                      initialdir = 'C:/Users/thinkpad/Desktop',
                      initialfile = '标签模板')
#            self.create_widgets.path_var = filenewpath
            if filenewpath:
                import shutil
                shutil.copy2(self.model2_path, filenewpath)
                self.model2.set(filenewpath)
                messagebox.showinfo("提示", f"模板文件已成功保存到:\n{filenewpath}")
            else:
                messagebox.showinfo("提示", f"下载已取消")
        #保存文件
        except Exception as e:
            messagebox.showwarning(title = '提示', message = f'未保存任何文件，错误信息{e}')
    def download_excel(self):
        try:
            filetypes = [("Excel工作簿(*.xlsx)", ".xlsx"), ("Excel97-2003工作簿(*.xls)", ".xls")]
            filenewpath = filedialog.asksaveasfilename(title = '保存模板文件', 
                      filetypes = filetypes, defaultextension = '.xlsx', 
                      initialdir = 'C:/Users/thinkpad/Desktop',
                      initialfile = '发货信息模板')
#            self.create_widgets.path_var = filenewpath
            if filenewpath:
                import shutil
                shutil.copy2(self.model_excel_path, filenewpath)
                self.LoadExcel_path.set(filenewpath)
                messagebox.showinfo("提示", f"模板文件已成功保存到:\n{filenewpath}")
            else:
                messagebox.showinfo("提示", f"下载已取消")
        #保存文件
        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未保存任何文件')
            
    def move_list2(self):
        all_items = self.group2_listbox1.curselection()
        selected_items = [self.group2_listbox1.get(idx) for idx in all_items]
        
        # 复制到目标 Listbox
        for item in selected_items:
            # 检查是否已存在（可选）
            if item not in self.group2_listbox2.get(0, tk.END):
                self.group2_listbox2.insert(tk.END, item)
    def up_list2(self):
        selected = self.group2_listbox2.curselection()
        if selected and selected[0] > 0:
            index = selected[0]
            item = self.group2_listbox2.get(index)
            
            self.group2_listbox2.delete(index)
            self.group2_listbox2.insert(index - 1, item)
            self.group2_listbox2.selection_set(index - 1)
            self.group2_listbox2.see(index - 1)
    def down_list2(self):
        selected = self.group2_listbox2.curselection()
        if selected and selected[0] < self.group2_listbox2.size() - 1:
            index = selected[0]
            item = self.group2_listbox2.get(index)
            
            self.group2_listbox2.delete(index)
            self.group2_listbox2.insert(index + 1, item)
            self.group2_listbox2.selection_set(index + 1)
            self.group2_listbox2.see(index + 1)
    def remove_list2(self):
        all_items = self.group2_listbox2.curselection()
        for idx in reversed(all_items):
            self.group2_listbox2.delete(idx)
            
    def generate_QRinfo1(self):
        all_items = ""
        for i in range(self.group2_listbox2.size()):
            all_items += (self.group2_listbox2.get(i) + ";")
        self.group3_scrolledtext.insert(tk.END, all_items)
        self.group2_listbox2.delete(0, tk.END)
    
    def generate_QRinfo2(self):
        all_items = ""
        for i in range(self.group2_listbox2.size()):
            all_items += (self.group2_listbox2.get(i) + ";")
        self.group4_scrolledtext.insert(tk.END, all_items)
        self.group2_listbox2.delete(0, tk.END)
        
    def save1(self):
        try:
            filetypes = [("Word文档(*.docx)", ".docx"), ("Word97-2003工作簿", ".doc")]
            filenewpath = filedialog.asksaveasfilename(title = '选择保存路径', 
                      filetypes = filetypes, defaultextension = '.docx')
#            self.create_widgets.path_var = filenewpath
            self.save_path1.set(filenewpath)
        #保存文件

        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')
    def run1(self):
        try:
            load_excel = self.LoadExcel_path.get()
            save_word = self.save_path1.get()
            QR_size = float(self.QR_size1.get())
            QR_code = self.group3_scrolledtext.get(1.0, tk.END).rstrip('\n')
            QR_info = { "information" : QR_code,
                        "size" : QR_size,
                    }
#            print(QR_info)
            if self.model1.get():
                model1_path = self.model1.get()
            else:
                model1_path = self.model1_path
            start(model1_path, save_word, load_excel, QR_info)
            messagebox.showinfo("提示", f"唛头已生成，文件已成功保存到:\n{save_word}")
            
        except Exception as e:
            show = traceback.print_exc()
            messagebox.showerror(title = '提示', message = f'生成唛头错误，请检查输入信息是否完整，错误信息{show}')
            
    def save2(self):
        try:
            filetypes = [("Word文档(*.docx)", ".docx"), ("Word97-2003工作簿", ".doc")]
            filenewpath = filedialog.asksaveasfilename(title = '选择保存路径', 
                      filetypes = filetypes, defaultextension = '.docx')
#            self.create_widgets.path_var = filenewpath
            self.save_path2.set(filenewpath)
        #保存文件

        except Exception as e:
            messagebox.showwarning(title = '提示', message = '未选择任何文件')
    def run2(self):
        try:
            load_excel = self.LoadExcel_path.get()
            save_word = self.save_path2.get()
            QR_size = float(self.QR_size2.get())
            QR_code = self.group4_scrolledtext.get(1.0, tk.END).rstrip('\n')
            QR_info = { "information" : QR_code,
                        "size" : QR_size,
                    }
#            print(QR_info)
            if self.model2.get():
                model2_path = self.model2.get()
            else:
                model2_path = self.model2_path
            start(model2_path, save_word, load_excel, QR_info)
            messagebox.showinfo("提示", f"标签已生成，文件已成功保存到:\n{save_word}")
            
        except Exception as e:
            show = traceback.print_exc()
            messagebox.showerror(title = '提示', message = f'生成标签错误，请检查输入信息是否完整，错误信息{show}')
        
def center_window(win, window_width, window_height):
    # 获取屏幕的宽度和高度
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    
#    # 获取窗口的宽度和高度
#    window_width = 1200
#    window_height = 600
    
    # 计算窗口左上角应该放置的x、y坐标
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    
    # 设置窗口的位置为居中
    win.geometry("{}x{}+{}+{}".format(window_width, window_height, x, y))
#    win.iconbitmap("favicon.ico")
    win.iconbitmap(get_path('favicon.icns'))
    
def get_path(ico_file):  #设置图标
    try:
        base_path = sys._MEIPASS
        ico_file = 'img\\' + ico_file
    except AttributeError:
        base_path = os.path.abspath('.')
    return os.path.normpath(os.path.join(base_path, ico_file))

if __name__ == '__main__':
    app = MainPage()
    center_window(app, window_width=1200, window_height=600)
    app.mainloop()
