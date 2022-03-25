# -*- coding: utf-8 -*-
# @Time : 2021-3-6 9:10
# @Author : Administrator
# @Email : lcmmljs@foxmail.com
# @File : main.py
# @Project : Project

import tkinter.filedialog as tkFileDialog
import tkinter.messagebox
from tkinter import *
from tkinter.font import Font
from tkinter.ttk import *

import SendEmail


class Application_ui(Frame):
    # 这个类仅实现界面生成功能，具体事件处理代码在子类Application中。
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('SendToKindle')
        self.master.geometry('372x166')
        self.createWidgets()

    def createWidgets(self):
        self.top = self.winfo_toplevel()

        self.style = Style()

        self.List1Var = StringVar(value='')
        self.List1Font = Font(font=('宋体', 9))
        self.List1 = Listbox(self.top, listvariable=self.List1Var, font=self.List1Font)
        self.List1.place(relx=0.215, rely=0.096, relwidth=0.626, relheight=0.458)

        self.style.configure('Label2.TLabel', anchor='w', font=('宋体', 9))
        self.Label2 = Label(self.top, text='附件列表：', style='Label2.TLabel')
        self.Label2.place(relx=0.022, rely=0.096, relwidth=0.175, relheight=0.102)

        self.style.configure('Command1.TButton', font=('宋体', 9))
        self.Command1 = Button(self.top, text='添加附件', command=self.Command1_Cmd, style='Command1.TButton')
        self.Command1.place(relx=0.258, rely=0.578, relwidth=0.239, relheight=0.151)

        self.style.configure('Command2.TButton', font=('宋体', 9))
        self.Command2 = Button(self.top, text='发送到Kindle', command=self.Command2_Cmd, style='Command2.TButton')
        self.Command2.place(relx=0.516, rely=0.578, relwidth=0.239, relheight=0.151)


class Application(Application_ui):
    # 这个类实现具体的事件处理回调函数。界面生成代码在Application_ui中。
    def __init__(self, master=None):
        Application_ui.__init__(self, master)
        self.m = SendEmail.Mail("***", "smtp.qq.com", "***", "***")  # 此处需要自行设置，第一个kindle接收邮箱，第二个smtp的邮件
                                                                     # 服务器，第三个邮箱，第四个smtp服务器密码·

    def Command1_Cmd(self, event=None):
        filenames = tkFileDialog.askopenfilenames(filetypes=[('mobi', '.mobi'), ('txt', '.txt')])
        for filename in filenames:
            self.List1.insert(END, filename.split("/")[-1])
            self.m.add_attachment(filename)

    def Command2_Cmd(self, event=None):
        if m.send():
            tkinter.messagebox.showinfo("提示", "邮件发送成功！")
        else:
            tkinter.messagebox.showinfo("提示", "邮件发送失败！")


if __name__ == "__main__":
    top = Tk()
    Application(top).mainloop()
    try:
        top.destroy()
    except:
        pass
