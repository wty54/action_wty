# -*- coding: utf-8 -*-
"""
Created on Wed Apr  1 13:31:12 2026

@author: thinkpad
"""

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, filedialog
from gui_output import center_window
from gui_output import MainPage

class Login_page(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("采购订单信息生成工具")
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.username.set("admin")
        self.protocol("WM_DELETE_WINDOW", self.quit_page)
        self.create_page()
    def create_page(self):
        self.mainframe = ttk.Frame(self)
        self.mainframe.place(anchor='c', relx=.50, rely=.30)
        user_label = ttk.Label(self.mainframe, text="用户名")
        user_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        user_Entry = ttk.Entry(self.mainframe, textvariable = self.username)
        user_Entry.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')
        password_label = ttk.Label(self.mainframe, text="密码")
        password_label.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')
        password_Entry = ttk.Entry(self.mainframe, textvariable = self.password, show="*")
        password_Entry.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')
        
        self.mainframe2 = ttk.Frame(self)
        self.mainframe2.place(anchor='c', relx=.50, rely=.70)
        login_button = ttk.Button(self.mainframe2, text="登录", command=self.login)
        login_button.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        quit_button = ttk.Button(self.mainframe2, text="退出", command=self.quit_page)
        quit_button.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')
        
    def login(self):
        username1 = self.username.get()
        password1 = self.password.get()
        if username1 == "admin" and password1 == "HT123456":
#            print("登陆成功")
            self.destroy()
            app = MainPage()
            center_window(app, window_width=1200, window_height=600)
            app.mainloop()
        else:
            messagebox.showwarning(title="提示", message="登录失败，请检查用户名、密码是否正确。")
        
    def quit_page(self):
        self.destroy()
        self.quit()
        
if __name__ == "__main__":
    app = Login_page()
    center_window(app, window_width=300, window_height=200)
    app.mainloop()
    