import pyautogui

##プロセスを制御するためにOS周りのモジュール
import re
import os
import subprocess
import sys
import time
import array

##Win32のUI情報と制御用モジュール
import win32api
import win32gui
import win32con

def ok_btn_action(action_name):

    time_end = 60
    window_cnt = 1
    if action_name == "upload_click":
        time_end = 120
        window_cnt = 2      
        

    # result = 0
    for i in range(0, time_end, 10):
        
        #実行前の待機(秒)
        time.sleep(2)
        #画面サイズの取得
        screen_x,screen_y = pyautogui.size()

        #win32guiを使ってウインドウタイトルを探す
        #Windowのハンドル取得('クラス名','タイトルの一部')で検索クラスがわからなかったらNoneにする
        parent_handle = win32gui.FindWindow("#32770","Microsoft Excel")

        if parent_handle > 0 :

            # titlebar = win32gui.GetWindowText(parent_handle)
            # classname = win32gui.GetClassName(parent_handle)
            # print('titlebar='+titlebar)
            # print('classname='+classname)

            if action_name == "nameSearch_click":
                x=112
                y=105
            elif action_name =="upload_click":
                x=160
                y=104
            #ウィンドウを最前面へ
            win32gui.SetForegroundWindow(parent_handle)

            # serch_xy(parent_handle,x,y)
            pyautogui.press('enter')
            window_cnt -= 1

            if window_cnt == 0:
                break
        
            #serch_xy(parent_handle,x,y)

        else:
            parent_handle = win32gui.FindWindow(None,"確認")
            if parent_handle > 0 :
                titlebar = win32gui.GetWindowText(parent_handle)
                classname = win32gui.GetClassName(parent_handle)
                # print('titlebar='+titlebar)
                # print('classname='+classname)                
                x=157
                y=123
                serch_xy(parent_handle,x,y)

                window_cnt -= 1

                if window_cnt == 0:
                    break
            # else:
            #     print(u"Microsoft Excelウィンドウなし")



def serch_xy(parent_handle,x,y):
    win_x1,win_y1,win_x2,win_y2 = win32gui.GetWindowRect(parent_handle)
    
    apw_x = win_x2 - win_x1
    apw_y = win_y2 - win_y1
    
    #ウィンドウを最前面へ
    win32gui.SetForegroundWindow(parent_handle)
    #ウインドウの完全な情報を取ってくる、FindWindowで部分一致だったりした場合の補完用

    #ボタン位置
    apw_btn_x = win_x1 + x
    apw_btn_y = win_y1 + y

    #[OK]ボタンクリック
    pyautogui.click(apw_btn_x,apw_btn_y)
    #print(u"ボタンの座標:"+str(apw_btn_x)+"/"+str(apw_btn_y))


if __name__ == "__main__":

    ok_btn_action("Microsoft Excel")
  

   