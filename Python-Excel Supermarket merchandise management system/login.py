import pymysql
import tkinter
import tkinter.font as tkFont
from tkinter import *
from tkinter import messagebox

from index import Application as index_ui

# 登陆界面ui设计与功能实现
class Application(Frame):
    # 01 初始化参数
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.username = tkinter.StringVar
        self.password = tkinter.StringVar
        self.createWidget()

    def createWidget(self):
        # 设置背景图片
        global photo_login
        photo_login = PhotoImage(file='./images/login.gif')
        self.label_bg = Label(self, image=photo_login)
        self.label_bg.pack()

        # 欢迎语区域设计
        global photo_login_center
        photo_login_center = PhotoImage('./images/login_center.gif')
        label_login_center = Label(self, image=photo_login_center, width='600', height='340')
        label_login_center.place(x='100', y='100')
        welcome_font_1 = tkFont.Font(family='宋体', size=30, weight=tkFont.BOLD)
        label_welcome_1 = Label(self, text='Welcome!!!', font=welcome_font_1, bg='#56cdff')
        label_welcome_1.place(x=175, y=170)
        welcome_font_2 = tkFont.Font(family='宋体', size=20, weight=tkFont.BOLD)
        label_welcome_2 = Label(self, text='欢迎登录Python!', font=welcome_font_2, bg='#56cdff')
        label_welcome_2.place(x=150, y=260)
        label_welcome_3 = Label(self, text='超市管理系统', font=welcome_font_2, bg='#56cdff')
        label_welcome_3.place(x=170, y=320)

        # 数据输入区域
        login_font = tkFont.Font(family='宋体', size=15)
        login_user = Label(self, text='账户', font=login_font, bg='#ffffff')
        login_user.place(x=430, y=160)
        Entry(self, highlightthickness=1, font=('宋体', 15), bg='#ffffff', textvariable= self.username).\
            place(x=430, y=190, width=240, height=40)
        login_pass = Label(self, text='密码', font=login_font, bg='#F3F3F4', show='*',
                           textvariable=self.password).place(x=430, y=270, width=240, height=40)
        Button(self, text='立即登录',font=('宋体', 15, 'bold'), fg='#000000', bg='#56cdff',
               command=self.login).place(x=430, y=340, width=240, height=40)
    # 登录方法
    def login(self):
        errMessage=''
        username = self.username.get()
        password = self.password.get()
        new_errMessage = self.login_verify(username, password)
        errMessage = errMessage + new_errMessage
        if errMessage != '':
            messagebox.showinfo('提示', errMessage)
            self.password.set()
        if new_errMessage =='登录成功':
            self.login_destory()
            # 加载index界面
            index_ui(self.master)

    # 登录验证
    def login_verity(self, username, password):
        errMessage=''
        result = False
        if len(username) ==0:
            errMessage = errMessage + '用户名不能为空'
        elif len(password) ==0:
            errMessage = errMessage + '用户名不能为空'
        else:
            result = self.login_db(username, password)
        if result:
            errMessage = errMessage + '登录成功'
        else:
            errMessage = errMessage + '用户名或密码错误'
        return errMessage

    # 登录处理
    def login_db(self, username, password):
        print(username, password)
        conn = pymysql.connect(host='localhost', user='root', password='123456', database='db_shopping')
        cursor = conn.cursor()
        sql = 'select * from user where username="'+username+'"and password="'+ password
        cursor.execute(sql)
        result = cursor.fetchall()
        conn.commit()
        conn.close()
        if len(result) ==0:
            return False
        else:
            return True

    # 销毁界面
    def login_destory(self):
        self.destroy()

