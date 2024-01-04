import tkinter as tk
from SMS import SMS
from UI import ui

if __name__ == '__main__':
    root = tk.Tk()
    w = 800
    h = 600
    root['width'] = w
    root['height'] = h
    # 得到屏幕的宽度和高度
    sw = root.winfo_screenwidth()
    # 得到屏幕的高度
    sh = root.winfo_screenmmheight()
    # 窗口居中
    x = (sw - w) / 2
    y = (sh - h) / 2
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    root.title('python学生信息管理系统（文件版）-html')
    sms = SMS()
    ui(root, sms)
    root.mainloop()

