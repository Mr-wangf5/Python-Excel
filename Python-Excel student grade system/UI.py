import tkinter as tk
import tkinter.messagebox

class ui:
    # 界面实现
    # 01 初始化
    def __init__(self, master, student):
        self.master = master
        self.student = student
        self.lbl = tk.Label(self.master, text='python学生信息管理系统（文件版）',
                            font=('microsoft yahei', 25), fg='DarkViolet')
        self.lbl.place(x=80, y=0)
        self.f1 = None
        self.createWidgest()

    def createWidgest(self):
        if self.f1:
            self.f1.destory()
        self.f1 = tk.Frame(self.master)
        self.f1['width'] = 800
        self.lab01 = tk.Label(self.f1, text='学生信息管理', font=('黑体', 18), width=23, fg='balck')
        self.lab01.place(x=65, y=0)
        self.lab02 = tk.Label(self.f1, text='学生信息管理', width=23, font=('microsoft yahei', 25), fg='DarkViolet')
        self.lab02.place(x=460, y=0)

        # 增加学生信息
        self.btn1 = tk.Button(self.f1, text='增加学生信息', font=('microsoft yahei', 12),
                              width=23, fg='DarkViolet', command=self.add_student_info)
        self.btn1.config(bg='skyblue')
        self.btn1.place(x=80, y=50)

        # 增加学生成绩信息
        self.btn12 = tk.Button(self.f1, text='增加学生成绩信息', font=('microsoft yahei', 12),
                              width=23, fg='DarkViolet', command=self.add_grade_info)
        self.btn12.config(bg='skyblue')
        self.btn12.place(x=480, y=50)

        # 修改学生信息
        self.btn2 = tk.Button(self.f1, text='修改学生信息', font=('microsoft yahei', 12),
                              width=23, fg='DarkViolet', command=self.modify_student_info)
        self.btn2.config(bg='skyblue')
        self.btn2.place(x=80, y=100)

        # 修改学生成绩
        self.btn22 = tk.Button(self.f1, text='修改学生成绩', font=('microsoft yahei', 12),
                              width=23, fg='DarkViolet', command=self.modify_grade_info)
        self.btn22.config(bg='skyblue')
        self.btn22.place(x=480, y=100)

        # 删除学生信息
        self.btn3 = tk.Button(self.f1, text='删除学生信息', font=('microsoft yahei', 12),
                              width=23, fg='DarkViolet', command=self.delete_student_info)

        self.btn3.config(bg='skyblue')
        self.btn3.place(x=80, y=150)

        # 删除学生成绩
        self.btn32 = tk.Button(self.f1, text='删除学生成绩', font=('microsoft yahei', 12),
                              width=23, fg='DarkViolet', command=self.delete_grade_info)

        self.btn32.config(bg='skyblue')
        self.btn32.place(x=480, y=150)

        # 查询学生信息
        self.btn4 = tk.Button(self.f1, text='查询学生信息', font=('microsoft yahei', 12),
                               width=23, fg='DarkViolet', command=self.search_student_info)

        self.btn4.config(bg='skyblue')
        self.btn4.place(x=80, y=200)

        # 查询学生成绩
        self.btn42 = tk.Button(self.f1, text='查询学生成绩', font=('microsoft yahei', 12),
                              width=23, fg='DarkViolet', command=self.search_grade_info)

        self.btn42.config(bg='skyblue')
        self.btn42.place(x=480, y=200)

        # 显示学生所有信息
        self.btn5 = tk.Button(self.f1, text='显示学生所有信息', font=('microsoft yahei', 12),
                               width=23, fg='DarkViolet', command=self.display_student)

        self.btn5.config(bg='skyblue')
        self.btn5.place(x=80, y=250)

        # 显示学生所有成绩
        self.btn52 = tk.Button(self.f1, text='显示学生所有成绩', font=('microsoft yahei', 12),
                              width=23, fg='DarkViolet', command=self.display_grade)

        self.btn52.config(bg='skyblue')
        self.btn52.place(x=480, y=250)

        # 生产学生个人信息表按钮
        self.btn6 = tk.Button(self.f1, text='显示学生所有成绩', font=('microsoft yahei', 12),
                               width=23, fg='DarkViolet', command=self.create_new_xlsx)

        self.btn6.config(bg='skyblue')
        self.btn6.place(x=0, y=80)

    # -------------------功能实现————————————————————
    # 01 添加学生信息
    def add_student_info(self):
        if self.f1:
            self.f1.destory()
        self.f1['width'] = 800
        self.f1['height'] = 800
        self.lf = tk.LabelFrame(self.f1, text='添加学生信息')
        self.sName = tk.Label(self.lf, text='姓名')
        self.sName.pack()
        self.sName_entry = tk.Entry(self.lf)
        self.sName_entry.pack()
        self.sNo = tk.Label(self.lf, text='学号')
        self.sNo.pack()
        self.sNo_entry = tk.Entry(self.lf)
        self.sNo_entry.pack()
        self.sSex = tk.Label(self.lf, text='性别')
        self.sSex.pack()
        self.sSex_entry = tk.Entry(self.lf)
        self.sSex_entry.pack()
        self.tele_num = tk.Label(self.lf, text='手机号')
        self.tele_num.pack()
        self.tele_num_entry = tk.Entry(self.lf)
        self.tele_num_entry.pack()

        # 添加数据到数据表中
        def cmd1():
            self.student.add_student(self.sName_entry.get(),
                                     self.sNo_entry.get(),
                                     self.sSex_entry.get(),
                                     self.tele_num_entry.get(),
            )
            self.sName_entry.delete(0, tk.END)
            self.sNo_entry.delete(0, tk.END)
            self.sSex_entry.delete(0, tk.END)
            self.tele_num_entry.delete(0, tk.END)

        self.btn1 = tk.Button(self.lf, text='返回', command = self.createWidgest())
        self.btn1.pack(self='left')
        self.btn1 = tk.Button(self.lf, text='确定', command=cmd1)
        self.btn1.pack(self='right')
        self.lf.place(x=130, y=40)
        self.f1.place(x=0, y=80)

    # 02 添加成绩信息
    def add_grade_info(self):
        """添加学生信息"""
        if self.f1:
            self.f1.destory()
        self.f1['width'] = 400
        self.f1['height'] = 500
        self.lf = tk.LabelFrame(self.f1, text='添加学生成绩信息')
        # 输入框：学号，姓名，语文，数学，英语，物理，化学，生物
        self.sNo = tk.Label(self.lf, text='学号')
        self.sNo.pack()
        self.sNo_entry = tk.Entry(self.lf)
        self.sNo_entry.pack()
        self.sName = tk.Label(self.lf, text='姓名')
        self.sName.pack()
        self.sName_entry = tk.Entry(self.lf)
        self.sName_entry.pack()
        self.sLange = tk.Label(self.lf, text='语文')
        self.sLange.pack()
        self.sLange_entry = tk.Entry(self.lf)
        self.sLange_entry.pack()
        self.sMath = tk.Label(self.lf, text='数学')
        self.sMath.pack()
        self.sMath_entry = tk.Entry(self.lf)
        self.sMath_entry.pack()
        self.sOLange = tk.Label(self.lf, text='英语')
        self.sOLange.pack()
        self.sOLange_entry = tk.Entry(self.lf)
        self.sOLange_entry.pack()
        self.sPysics = tk.Label(self.lf, text='物理')
        self.sPysics.pack()
        self.sPysics_entry = tk.Entry(self.lf)
        self.sPysics_entry.pack()
        self.sChemistry = tk.Label(self.lf, text='化学')
        self.sChemistry.pack()
        self.sChemistry_entry = tk.Entry(self.lf)
        self.sChemistry_entry.pack()
        self.sHenwu = tk.Label(self.lf, text='生物')
        self.sHenwu.pack()
        self.sHenwu_entry = tk.Entry(self.lf)
        self.sHenwu_entry.pack()
        # 添加数据到数据表中
        def cmd1():
            self.student.add_student(self.sNo_entry.get(),
                                     self.sName_entry.get(),
                                     self.sLange_entry.get(),
                                     self.sOLange_entry.get(),
                                     self.sMath_entry.get(),
                                     self.sPysics_entry.get(),
                                     self.sChemistry_entry.get(),
                                     self.sHenwu_entry.get()
                                     )
            self.sName_entry.delete(0, tk.END)
            self.sNo_entry.delete(0, tk.END)
            self.sLange_entry.delete(0, tk.END)
            self.sOLange_entry.delete(0, tk.END)
            self.sMath_entry.delete(0, tk.END)
            self.sPysics_entry.delete(0, tk.END)
            self.sChemistry_entry.delete(0, tk.END)
            self.sHenwu_entry.delete(0, tk.END)

        self.btn1 = tk.Button(self.lf, text='返回', command=self.createWidgest())
        self.btn1.pack(self='left')
        self.btn1 = tk.Button(self.lf, text='确定', command=cmd1)
        self.btn1.pack(self='right')
        self.lf.place(x=130, y=40)
        self.f1.place(x=0, y=80)

    # 03 修改学生成绩
    def modify_grade_info(self):
        if self.f1:
            self.f1.destory()
            self.f1 = tk.Frame(self.master)
            self.f1['width'] = 500
            self.f1['height'] = 700
            self.lf = tk.LabelFrame(self.f1, text='修改学生成绩')
            self.index1 = -1
        # 检索学号是否存在
        def cmd2(sno):
            if sno == '':
                tk.messagebox.showinfo('提示', '请输入学号！')
            else:
                self.index1 = self.student.judge_student(sno)

        self.old_sNo = tk.Label(self.lf, text='学号')
        self.old_sNo.pack()
        self.old_sNo_entry = tk.Entry(self.lf)
        self.old_sNo_entry.pack()
        self.btn = tk.Button(self.lf, text='检索', command=lambda:cmd2(self.old_sNo_entry.get()))
        self.btn.pack(side='left')

        # 修改信息
        self.sName = tk.Label(self.lf, text='姓名')
        self.sName.pack()
        self.sName_entry = tk.Entry(self.lf)
        self.sName_entry.pack()
        self.sLange = tk.Label(self.lf, text='语文')
        self.sLange.pack()
        self.sLange_entry = tk.Entry(self.lf)
        self.sLange_entry.pack()
        self.sMath= tk.Label(self.lf, text='数学')
        self.sMath.pack()
        self.sMath_entry = tk.Entry(self.lf)
        self.sMath_entry.pack()
        self.sPysics= tk.Label(self.lf, text='物理')
        self.sPysics.pack()
        self.sPysics_entry = tk.Entry(self.lf)
        self.sPysics_entry.pack()
        self.sHenwu = tk.Label(self.lf, text='生物')
        self.sHenwu.pack()
        self.sHenwu_entry = tk.Entry(self.lf)
        self.sHenwu_entry.pack()
        self.sChemistry= tk.Label(self.lf, text='化学')
        self.sChemistry.pack()
        self.sChemistry_entry = tk.Entry(self.lf)
        self.sChemistry_entry.pack()
        def cmd():
            self.student.modify_grade(self.index1,
                                      self.sNo_entry.get(),
                                      self.sName_entry.get(),
                                      self.sLange_entry.get(),
                                      self.sOLange_entry.get(),
                                      self.sMath_entry.get(),
                                      self.sPysics_entry.get(),
                                      self.sChemistry_entry.get(),
                                      self.sHenwu_entry.get()
            )
            self.sName_entry.delete(0, tk.END)
            self.sNo_entry.delete(0, tk.END)
            self.sLange_entry.delete(0, tk.END)
            self.sOLange_entry.delete(0, tk.END)
            self.sMath_entry.delete(0, tk.END)
            self.sPysics_entry.delete(0, tk.END)
            self.sChemistry_entry.delete(0, tk.END)
            self.sHenwu_entry.delete(0, tk.END)

        self.btn1 = tk.Button(self.lf, text='返回', command=self.createWidgest())
        self.btn1.pack(self='left')
        self.btn1 = tk.Button(self.lf, text='确定', command=lambda :cmd())
        self.btn1.pack(self='right')
        self.lf.place(x=130, y=40)
        self.f1.place(x=0, y=80)



    # 04 修改学生信息
    def modify_student_info(self):
        if self.f1:
            self.f1.destory()
            self.f1 = tk.Frame(self.master)
            self.f1['width'] = 400
            self.f1['height'] = 300
            self.lf = tk.LabelFrame(self.f1, text='修改学生信息')
            self.index1 = -1
        # 检索学号是否存在
        def cmd2(sno):
            if sno == '':
                tk.messagebox.showinfo('提示', '请输入学号！')
            else:
                self.index1 = self.student.judge_student(sno)

        self.old_sNo = tk.Label(self.lf, text='学号')
        self.old_sNo.pack()
        self.old_sNo_entry = tk.Entry(self.lf)
        self.old_sNo_entry.pack()

        # 修改信息
        self.sName = tk.Label(self.lf, text='姓名')
        self.sName.pack()
        self.sName_entry = tk.Entry(self.lf)
        self.sName_entry.pack()
        self.sNo = tk.Label(self.lf, text='学号')
        self.sNo.pack()
        self.sSex = tk.Label(self.lf, text='性别')
        self.sSex.pack()
        self.sSex_entry = tk.Entry(self.lf)
        self.sSex_entry.pack()
        def cmd():
            self.student.modify_student(self.index1,
                                     self.sName_entry.get(),
                                     self.sNo_entry.get(),
                                     self.sSex_entry.get(),
                                     self.tele_num_entry.get(),
                                     )
            self.old_sNo_entry.delete(0, tk.END)
            self.sName_entry.delete(0, tk.END)
            self.sNo_entry.delete(0, tk.END)
            self.sSex_entry.delete(0, tk.END)
            self.tele_num_entry.delete(0, tk.END)
        self.btn1 = tk.Button(self.lf, text='返回', command=self.createWidgest())
        self.btn1.pack(self='left')
        self.btn1 = tk.Button(self.lf, text='确定', command=lambda:cmd())
        self.btn1.pack(self='right')
        self.lf.place(x=130, y=0)
        self.f1.place(x=0, y=80)


        self.btn = tk.Button(self.lf, text='检索', command=lambda:cmd2(self.old_sNo_entry.get()))
        self.btn.pack(side='left')
        self.btn1 = tk.Button(self.lf, text='删除', command=lambda: self.student.delete_student(self.index1))
        self.btn1.pack(side='right')
        self.btn2 = tk.Button(self.lf, text='返回', command= self.createWidgest)
        self.btn2.pack()
        self.lf.place(x=130, y=0)
        self.f1.place(x=0, y=80)

    # 05 删除学生信息
    def delete_grade_info(self):
        """删除学生"""
        if self.f1:
            self.f1.destory()
            self.f1 = tk.Frame(self.master)
            self.f1['width'] = 400
            self.f1['height'] = 300
            self.lf = tk.LabelFrame(self.f1, text='删除学生信息')
            self.index2 = -1
        # 检索学号是否存在
        def cmd3(sno):
            if sno == '':
                tk.messagebox.showinfo('提示', '请输入学号！')
            else:
                self.index2 = self.student.judge_student(sno)

        self.old_sNo = tk.Label(self.lf, text='学号')
        self.old_sNo.pack()
        self.old_sNo_entry = tk.Entry(self.lf)
        self.old_sNo_entry.pack()
        self.btn = tk.Button(self.lf, text='检索', command=lambda:cmd3(self.old_sNo_entry.get()))
        self.btn.pack(side='left')
        self.btn1 = tk.Button(self.lf, text='删除', command=lambda: self.student.delete_student(self.index2))
        self.btn1.pack(side='right')
        self.btn2 = tk.Button(self.lf, text='返回', command= self.createWidgest)
        self.btn2.pack()
        self.lf.place(x=130, y=0)
        self.f1.place(x=0, y=80)

    # 06 删除学生成绩信息
    def delete_grade_info(self):
        """删除学生"""
        if self.f1:
            self.f1.destory()
            self.f1 = tk.Frame(self.master)
            self.f1['width'] = 400
            self.f1['height'] = 300
            self.lf = tk.LabelFrame(self.f1, text='删除学生信息')
            self.index2 = -1
        # 检索学号是否存在
        def cmd3(sno):
            if sno == '':
                tk.messagebox.showinfo('提示', '请输入学号！')
            else:
                self.index2 = self.student.judge_grade(sno)

        self.old_sNo = tk.Label(self.lf, text='学号')
        self.old_sNo.pack()
        self.old_sNo_entry = tk.Entry(self.lf)
        self.old_sNo_entry.pack()
        self.btn = tk.Button(self.lf, text='检索', command=lambda:cmd3(self.old_sNo_entry.get()))
        self.btn.pack(side='left')
        self.btn1 = tk.Button(self.lf, text='删除', command=lambda: self.student.delete_student(self.index2))
        self.btn1.pack(side='right')
        self.btn2 = tk.Button(self.lf, text='返回', command= self.createWidgest)
        self.btn2.pack()
        self.lf.place(x=130, y=0)
        self.f1.place(x=0, y=80)

    # 07 查询学生信息
    def search_student_info(self):
        """查找学生"""
        if self.f1:
            self.f1.destory()
            self.f1 = tk.Frame(self.master)
            self.f1['width'] = 400
            self.f1['height'] = 300
            self.lf = tk.LabelFrame(self.f1, text='搜索学生信息')
            self.index1 = -1
        # 检索学号是否存在
        def cmd3(sno):
            if sno == '':
                tk.messagebox.showinfo('提示', '请输入学号！')
            else:
                self.index2 = self.student.serch_student(sno)

        self.old_sNo = tk.Label(self.lf, text='学号')
        self.old_sNo.pack()
        self.btn = tk.Button(self.lf, text='返回', command=self.createWidgest)
        self.btn.pack(side='left')
        self.btn1 = tk.Button(self.lf, text='查找', command=lambda: cmd3(self.sNo_entry.get()))
        self.lf.place(x=130, y=0)
        self.f1.place(x=0, y=80)

    # 查询学生成绩
    def search_grade_info(self):
        """查找学生"""
        if self.f1:
            self.f1.destory()
            self.f1 = tk.Frame(self.master)
            self.f1['width'] = 400
            self.f1['height'] = 300
            self.lf = tk.LabelFrame(self.f1, text='搜索学生成绩')
            self.index1 = -1
        # 检索学号是否存在
        def cmd3(sno):
            if sno == '':
                tk.messagebox.showinfo('提示', '请输入学号！')
            else:
                self.index2 = self.student.serch_grade(sno)

        self.old_sNo = tk.Label(self.lf, text='学号')
        self.old_sNo.pack()
        self.btn = tk.Button(self.lf, text='返回', command=self.createWidgest)
        self.btn.pack(side='left')
        self.btn1 = tk.Button(self.lf, text='查找', command=lambda: cmd3(self.sNo_entry.get()))
        self.lf.place(x=130, y=0)
        self.f1.place(x=0, y=80)

    # 09 显示学生信息表
    def display_student(self):
        if self.f1:
            self.f1.destory()
        self.f1['width'] = 400
        self.f1['height'] = 300
        self.lf = tk.LabelFrame(self.f1, text='显示学生信息表')
        str1 = self.student.display_students()
        self.text = tk.Text(self.lf, width=56, height=15)
        self.text.insert(0.0, str1)
        self.text.pack()
        self.btn1 = tk.Button(self.lf, text='返回', command=self.createWidgest)
        self.btn1.pack()
        self.lf.place(x=0, y=40)
        self.f1.place(x=0, y=80)

    # 显示学生成绩表
    def display_grade(self):
        if self.f1:
            self.f1.destory()
        self.f1['width'] = 400
        self.f1['height'] = 300
        self.lf = tk.LabelFrame(self.f1, text='导出学生信息表')

        self.index2, self.index1 =-1, -1
        # 检索学号是否存在
        def cmd3(sno):
            if sno == '':
                tk.messagebox.showinfo('提示', '请输入学号！')
            else:
                self.index2 = self.student.judge_grade(sno)
                self.index1 = self.student.judge_student(sno)

        def cmd4(s1, s2):
            self.student.crate_personal_xlsx(s1, s2)

            self.old_sNo = tk.Label(self.lf, text='学号')
            self.old_sNo.pack()
            self.old_sNo_entry = tk.Entry(self.lf)
            self.old_sNo_entry.pack()
            self.btn = tk.Button(self.lf, text='检索', command=lambda :cmd3(self.sNo_entry.get()))
            self.btn.pack(side='left')
            self.btn1 = tk.Button(self.lf, text='保存', command=lambda: cmd4(self.index1, self.index2))
            self.btn1.pack()
            self.btn2 = tk.Button(self.lf, text='返回', command=self.createWidgest)
            self.btn2.pack()
            self.lf.place(x=130, y=0)
            self.f1.place(x=0, y=80)















