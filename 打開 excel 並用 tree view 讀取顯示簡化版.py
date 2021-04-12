# -*- coding: utf-8 -*-
from tkinter import filedialog
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedStyle
import pandas as pd
def op():
 filename = filedialog.askopenfilename(title='選擇', filetypes=[('Excel', '*.xlsx'), ('All Files', '*')])
 return filename
def read(filename):
 df = pd.read_excel(filename,header=0)
 cols = list(df.columns)
 tree = ttk.Treeview(root)
 tree.pack()
 tree["columns"] = cols
 for i in cols:
    tree.column(i,anchor="center")
    tree.heading(i,text=i,anchor='center')
 for index, row in df.iterrows():
    tree.insert("",'end',text = index,values=list(row))
# x 設定零，讓 tree view 從 ui 的初始位置開始，
# y 的設定不要擋到上方的按鈕，
# heigh 的設定決定 tree view 上下有多寬，
# width 設定一，填滿 ui 的左右
 tree.place(relx=0,rely=0.1,relheight=0.7,relwidth=1)
def oas():
 filename = op()
 read(filename)
def main():
 global root
 root = tk.Tk()
 style = ThemedStyle(root)
 style.set_theme("breeze")
 root.geometry("1600x750")
 B1 = tk.Button(root, text="打開",command = oas).pack()
 root.mainloop()
if __name__=='__main__':
 main()
