# -*- coding: utf-8 -*-
import xlrd
# 简单的图形界面GUI（Graphical User Interface）
from tkinter import *
import tkinter.filedialog
import tkinter.messagebox as messagebox
import win32gui
import win32con
import win32api
import win32clipboard as w
import time

root = Tk()


class Application(Frame):  # 从Frame派生出Application类，它是所有widget的父容器
    def __init__(self, master=None):  # master即是窗口管理器，用于管理窗口部件，如按钮标签等，顶级窗口master是None，即自己管理自己
        Frame.__init__(self, master)
        self.pack()  # 将widget加入到父容器中并实现布局
        self.createWidgets()
        self.table = 0
        self.row = 0
        self.directory = ""
        self.win = 0

    def createWidgets(self):
        inputTableVar = StringVar()
        self.inputTableTagLabel = Label(self, text='请输入需要取出的table')  # 创建一个标签显示内容到窗口
        self.inputTableTagLabel.pack()
        self.inputTable = Entry(self, textvariable=inputTableVar)  # 创建一个输入框，以输入内容
        self.inputTable.pack()
        inputTableVar.set("0")

        inputColumnVar = StringVar()
        self.inputColumnTagLabel = Label(self, text='请输入需要取出的列')  # 创建一个标签显示内容到窗口
        self.inputColumnTagLabel.pack()
        self.inputColumn = Entry(self, textvariable=inputColumnVar)  # 创建一个输入框，以输入内容
        self.inputColumn.pack()
        inputColumnVar.set("0")

        self.nameButton = Button(self, text='选择excel', command=self.selectExcel)  # 创建一个hello按钮，点击调用hello方法，实现输出
        self.nameButton.pack()

        # self.selectFilesButton = Button(self, text='选择文件夹', command=self.selectFiles)  # 创建一个hello按钮，点击调用hello方法，实现输出
        # self.selectFilesButton.pack()

        operationWinVar = StringVar()
        self.operationWinTagLabel = Label(self, text='要操作的窗口')  # 创建一个标签显示内容到窗口
        self.operationWinTagLabel.pack()
        self.operationWinTable = Entry(self, textvariable=operationWinVar)  # 创建一个输入框，以输入内容
        self.operationWinTable.pack()
        # operationWinVar.set("MiaoMore3.0")

        self.nextButton = Button(self, text='下一个', command=self.nextRow)  # 创建一个hello按钮，点击调用hello方法，实现输出
        self.nextButton.pack()

    def selectExcel(self):
        # name = self.input.get()  # 获取输入的内容
        # messagebox.showinfo('Message', 'hello,%s' % name)  # 显示输出
        filename = tkinter.filedialog.askopenfilename()
        if filename != '':
            print("您选择的文件是：" + filename)
            excelFile = xlrd.open_workbook(filename)
            print("table "+self.inputTable.get())
            self.table = excelFile.sheets()[int(self.inputTable.get())]  # 通过索引顺序获取
        else:
            print("您没有选择任何文件")

    # def selectFiles(self):
    #     # name = self.input.get()  # 获取输入的内容
    #     # messagebox.showinfo('Message', 'hello,%s' % name)  # 显示输出
    #     self.directory = tkinter.filedialog.askdirectory()
    #     if self.directory != '':
    #         print("您选择的文件夹是：" + self.directory)
    #     else:
    #         print("您没有选择任何文件")

    def nextRow(self):
        value = self.table.cell_value(self.row, int(self.inputColumn.get()))  # 返回单元格对象
        print(value)
        # while self.win == 0:
        #     hwnd = win32gui.FindWindow("SunAwtFrame", None)
        #     title = win32gui.GetWindowText(hwnd)
        #     if title.startswith(self.operationWinTable.get()):
        #         self.win = hwnd
        print(self.win)
        if self.win == 0:
            hWndList = []
            win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)
            for h in hWndList:
                if not h:
                    return
                title = win32gui.GetWindowText(h)
                clsname = win32gui.GetClassName(h)
                if clsname.__eq__("SunAwtFrame") and title.startswith(self.operationWinTable.get()):
                    print('窗口句柄:%s ' % (h))
                    print('窗口标题:%s' % (title))
                    print('窗口类名:%s' % (clsname))
                    self.win = h
                    win32gui.ShowWindow(self.win, win32con.SW_SHOWNORMAL)
                    win32gui.SetForegroundWindow(self.win)
                    break
        print(self.win)
        win32gui.ShowWindow(self.win, win32con.SW_SHOWNORMAL)
        win32gui.SetForegroundWindow(self.win)
        time.sleep(0.1)
        if self.row == 0:
            # ctrl H
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0);
            win32api.keybd_event(72, 0, 0, 0);
            win32api.keybd_event(72, 0, win32con.KEYEVENTF_KEYUP, 0);
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0);
            time.sleep(0.1)
        self.setClipboardtext(value)
        time.sleep(0.1)
        # ctrl A
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0);
        win32api.keybd_event(65, 0, 0, 0);
        win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0);
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)
        # ctrl V
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0);
        win32api.keybd_event(86, 0, 0, 0);
        win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0);
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)
        self.row += 1

    # 写入剪切板内容
    def setClipboardtext(self,aString):
        print("setClipboardtext" + aString)
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
        w.CloseClipboard()


app = Application()
app.master.title("查找器")  # 窗口标题
app.mainloop()  # 主消息循环
