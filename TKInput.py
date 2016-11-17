#!/usr/bin/env python
# -*- coding: utf-8 -*-

from Tkinter import *
import tkMessageBox
import ExcelMani
import Statistic

class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.nameInput = Entry(self)
        self.nameInput.pack()
        self.alertButton = Button(self, text='Get Res XLS', command=self.generateXls)
        self.alertButton.pack()
        self.quitButton = Button(self, text='Quit', command=self.quit)
        self.quitButton.pack()

    def generateXls(self):
        name = self.nameInput.get()
        ExcelMani.main(name)
        Statistic.main(name)

def main():
    app = Application()
    # 设置窗口标题:
    app.master.title('Name Get')
    # 主消息循环:
    app.mainloop()

if __name__ == '__main__':
    main()
