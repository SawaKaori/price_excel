!/bin/env python
# -*- encoding: cp932 -*-
"""
    ExcelのCOMオブジェクトの使い方テスト
"""
import sys
import win32com
import pdb
from win32com.client import *
import pprint
import datetime
import msvcrt

def main():
    xapp = win32com.client.Dispatch("Excel.Application")
    xapp.Visible = True
    
    book = xapp.Workbooks.Open(r"H:\user\test3\sp2Changes.xls")
    sheet = book.Worksheets.Item(r"SP-2 Changes")
    
    wk = sheet.UsedRange
    x_cnt = wk.Columns.Count
    y_cnt = wk.Rows.Count
    print "x_cnt=%d y_cnt=%d" % (x_cnt, y_cnt)
    
    st_tm = datetime.datetime.today()
    cnt = 0
    wk2 = wk.Value
    for y in range(0, y_cnt - 1):
        for x in range(0, x_cnt - 1):
            print "check x=%d y=%d  \r" % (x + 1, y + 1),
            if wk2[y][x] != None:
                cnt += 1
    print "総有効セル数=%d" % (cnt)
    en_tm = datetime.datetime.today()
    print "検査開始:%s\n検査終了:%s\n実行時間:%s" % (st_tm, en_tm, en_tm - st_tm)

if __name__ == "__main__":
    main()