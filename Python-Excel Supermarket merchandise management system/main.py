from tkinter import *
from login import Application as login

# 程序入口
if __name__ == 'main':
    win = Tk()
    w = 800
    h = 540
    # 得到屏幕的宽度和高度
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    # 窗口居中
    x = (sw - w) / 2
    y = (sh - h) / 2
    win.geometry('%dx%d+%d+%d' % (w, h, x, y))
    win.title('Python-Excel超市管理系统-bhml')
    win.resizable(width=False, height=False)  # 界面大小不允许修改
    # 加载login界面
    login(master = win)
    win.mainloop()
