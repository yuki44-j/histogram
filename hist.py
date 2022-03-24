#!/usr/bin/env python
# coding: utf-8

import tkinter
from tkinter import filedialog
from tkinter import simpledialog
import pandas as pd
import xlwings as xw

# histogram を作る本体関数
def make_hist():
    singleDf = pd.read_excel(filename, index_col=0)
    wb = xw.Book(filename)
    sheet = wb.sheets.add('hist')

    histFig = singleDf.plot.hist(bins=binNum).get_figure()
    sheet.pictures.add(histFig)

    binnedP = pd.cut(singleDf[:, 0], binNum)
    bp = binnedP.value_counts(sort=False)

    bpPair = zip(map(str, bp.index), bp)
    sheet.range("I2").value = list(bpPair)
    sheet.range("I1").value = ["ﾃﾞｰﾀ区間", "頻度"]

# ﾌｧｲﾙと整数入力のﾀﾞｲｱﾛｸﾞ
tk = tkinter.Tk()
tk.withdraw()

fTyp = [("Files","*.*")]
iDir = r'C:\Users'
titleText = "Select file"
file = filedialog.askopenfilename(filetypes = fTyp, title = titleText, initialdir = iDir)
filename = '\\'.join(file.split('/'))

entry = simpledialog.askinteger('bin の数', '10以外にしたいなら入力')

# ここから実行 code
binNum = 10
if entry:
    binNum = entry
    
make_hist()
