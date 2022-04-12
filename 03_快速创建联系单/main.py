# -*- coding: utf-8 -*-
# @Time : 2021-12-6 9:10
# @Author : Administrator
# @Email : lcmmljs@foxmail.com
# @File : main.py
# @Project : Project

import sys
import datetime
import tkinter.filedialog
import tkinter.messagebox as msgbox

from docxtpl import DocxTemplate
from tkinter import *
from tkinter.ttk import *


class Application_ui(Frame):
    # 这个类仅实现界面生成功能，具体事件处理代码在子类Application中。
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('创建联系单')
        self.master.geometry('250x450+800+200')
        self.createWidgets()

    def createWidgets(self):
        self.top = self.winfo_toplevel()

        self.style = Style()

        self.Frame3 = LabelFrame(self.top, text='发单信息')
        self.Frame3.place(relx=0.066, rely=0.565, relwidth=0.867, relheight=0.312)

        self.Frame2 = LabelFrame(self.top, text='图纸信息')
        self.Frame2.place(relx=0.066, rely=0.383, relwidth=0.867, relheight=0.166)

        self.Frame1 = LabelFrame(self.top, text='联系单信息')
        self.Frame1.place(relx=0.066, rely=0.018, relwidth=0.867, relheight=0.349)

        self.Command1 = Button(self.top, text='生成联系单', command=self.Command1_Cmd)
        self.Command1.place(relx=0.066, rely=0.893, relwidth=0.336, relheight=0.075)

        self.Command2 = Button(self.top, text='退出', command=self.Command2_Cmd)
        self.Command2.place(relx=0.465, rely=0.893, relwidth=0.336, relheight=0.075)

        self.Text5Var = StringVar(value='')
        self.Text5 = Entry(self.Frame3, text='Text5', textvariable=self.Text5Var)
        self.Text5.place(relx=0.344, rely=0.759, relwidth=0.579, relheight=0.182)

        self.Text4Var = StringVar(value='')
        self.Text4 = Entry(self.Frame3, text='Text4', textvariable=self.Text4Var)
        self.Text4.place(relx=0.306, rely=0.467, relwidth=0.426, relheight=0.182)

        self.Text3Var = StringVar(value='')
        self.Text3 = Entry(self.Frame3, text='Text3', textvariable=self.Text3Var)
        self.Text3.place(relx=0.344, rely=0.175, relwidth=0.502, relheight=0.182)

        self.style.configure('Label7.TLabel', anchor='w')
        self.Label7 = Label(self.Frame3, text='联系方式：', style='Label7.TLabel')
        self.Label7.place(relx=0.038, rely=0.759, relwidth=0.287, relheight=0.204)

        self.style.configure('Label6.TLabel', anchor='w')
        self.Label6 = Label(self.Frame3, text='发单人：', style='Label6.TLabel')
        self.Label6.place(relx=0.038, rely=0.467, relwidth=0.23, relheight=0.204)

        self.style.configure('Label5.TLabel', anchor='w')
        self.Label5 = Label(self.Frame3, text='发单日期：', style='Label5.TLabel')
        self.Label5.place(relx=0.038, rely=0.175, relwidth=0.287, relheight=0.204)

        self.Text2Var = StringVar(value='')
        self.Text2 = Entry(self.Frame2, text='Text2', textvariable=self.Text2Var)
        self.Text2.place(relx=0.23, rely=0.438, relwidth=0.656, relheight=0.342)

        self.style.configure('Label4.TLabel', anchor='w')
        self.Label4 = Label(self.Frame2, text='图名：', style='Label4.TLabel')
        self.Label4.place(relx=0.038, rely=0.438, relwidth=0.172, relheight=0.384)

        self.Text1Var = StringVar(value='')
        self.Text1 = Entry(self.Frame1, text='Text1', textvariable=self.Text1Var)
        self.Text1.place(relx=0.421, rely=0.732, relwidth=0.541, relheight=0.163)

        self.Combo2List = ['总体总包', '高架段', '地下段']
        self.Combo2 = Combobox(self.Frame1, values=self.Combo2List)
        self.Combo2.place(relx=0.421, rely=0.471, relwidth=0.541, relheight=0.131)
        self.Combo2.set(self.Combo2List[0])

        self.Combo1List = ['送审联系单', '回复咨询联系单', '出图联系单']
        self.Combo1 = Combobox(self.Frame1, values=self.Combo1List)
        self.Combo1.place(relx=0.421, rely=0.209, relwidth=0.541, relheight=0.131)
        self.Combo1.set(self.Combo1List[0])

        self.style.configure('Label3.TLabel', anchor='w')
        self.Label3 = Label(self.Frame1, text='联系单单号：', style='Label3.TLabel')
        self.Label3.place(relx=0.038, rely=0.732, relwidth=0.349, relheight=0.216)

        self.style.configure('Label2.TLabel', anchor='w')
        self.Label2 = Label(self.Frame1, text='联系单种类：', style='Label2.TLabel')
        self.Label2.place(relx=0.038, rely=0.471, relwidth=0.388, relheight=0.216)

        self.style.configure('Label1.TLabel', anchor='w')
        self.Label1 = Label(self.Frame1, text='联系单类型：', style='Label1.TLabel')
        self.Label1.place(relx=0.038, rely=0.209, relwidth=0.349, relheight=0.216)


class Application(Application_ui):
    # 这个类实现具体的事件处理回调函数。界面生成代码在Application_ui中。
    def __init__(self, master=None):
        Application_ui.__init__(self, master)
        # 获取今天日期，给发单日期一个默认值
        contactTime = datetime.date.today()
        self.Text3Var.set(contactTime)

    def Command1_Cmd(self, event=None):
        # 获取用户选择联系单类型，指定相应模板
        contactType = self.Combo1.current()
        if contactType == 0:
            fileRoute = r'.\铁四院[01标...]联字（2022）第...号--关于“提交...送咨询强审审查”的函.docx'
        elif contactType == 1:
            fileRoute = r'.\铁四院[01标...]联字（2022）第...号--关于“提交...结构施工图咨询意见回复”的函.docx'
        elif contactType == 2:
            fileRoute = r'.\铁四院[01标...]联字（2022）第...号--关于“提交...结构施工图”的函.docx'
        # 建立模板输入相关数据
        doc = DocxTemplate(fileRoute)
        docDict = {'type': self.Combo2.get(),
                   'num': self.Text1.get(),
                   'name': self.Text2.get(),
                   'date': self.Text3.get(),
                   'man': self.Text4.get(),
                   'phone': self.Text5.get(),
                   }
        doc.render(docDict)
        # 询问用户存储路径，生成模板文件
        try:
            docRoute = tkinter.filedialog.askdirectory()
            docNameList = fileRoute.lstrip(r'.\\').split('...')
            docName = docNameList[0] + docDict['type'] + docNameList[1] + str(docDict['num']) + docNameList[2] \
                      + docDict['name'] + docNameList[3]
            doc.save(docRoute + '/' + docName)
            msgbox.showinfo('提示', '文件生成完成！！！')
        except:
            pass

    def Command2_Cmd(self, event=None):
        sys.exit()


if __name__ == "__main__":
    top = Tk()
    Application(top).mainloop()
    try:
        top.destroy()
    except:
        pass
