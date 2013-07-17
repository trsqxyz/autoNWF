#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
NewWorkFriend の自動化
Excel にデータを出力して csv の保存
'''

import datetime
import time
import subprocess
import jholiday
from pywinauto import application

def MakePath():
    # 出力用のファイル生成
    d = DataDate()
    path = u'C:\\Users\\Public\\出荷データ' + d + u'.csv'
    path = path.encode('mbcs')    
    return path

def DataDate():
    '''
    出力データの日付設定 AM PM で日付を変更
    営業日のみを出力
    '''
    d = datetime.date.today()
    t1 = datetime.datetime.now().strftime('%H%M')
    t2 = datetime.time(15,00).strftime('%H%M')
    add_d = datetime.timedelta(days = 1)

    # 午後のとき、翌日分に
    if t1 < t2:
        pass
    elif t1 > t2:
        d = d + add_d

    while True:
        # 平日に修正
        while d.weekday() >= 5:
            d = d + add_d
        
        #祝日のとき、翌日へ
        if jholiday.holiday_name(date = d):
            d = d + add_d
        else:
            break
        
    d = d.strftime('%Y%m%d')
    
    return d



def login():
    '''
    NWF ログイン処理
    '''
    
    app = application.Application()
    # password 入力
    while True:
        dialog = app[u'ログイン']
        while True:
            try:
                if dialog.Edit2.WindowText() == '':
                    pass
            except:
                break

        try:
            dialog = app[u'検索条件値の入力']
            dialog.Edit.SetText(DataDate())
            dialog.TypeKeys('{ENTER}')
            break
        except:
            break

def copyToClipboard( text ):
    '''
    text copy to clipboard
    '''
    
    try:
        p = subprocess.Popen( ['clip'], stdin=subprocess.PIPE, shell=True )
        p.stdin.write( text )
        p.stdin.close()
 
        retcode = p.wait()
        return True
    except Exception, inst:
        return False

def NWF():
    '''
    NWF を実行し、適切なファイル形式で保存する。
    '''
    
    app = application.Application()
    try:
        app.start_('C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE')
    except:
        app.connect_('C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE')
    
    # 新規ブック
    try:
        app.Microsoft_Excel.TypeKeys('^n')
    except:
        dialog = app[u'使用中のファイル']
        dialog.TypeKeys('{ESC}')
        app.Microsoft_Excel.TypeKeys('^n')
    
    # NewWorkFriend 実行
    app.Microsoft_Excel.TypeKeys('%a' 'q')

    # ウインドウ遷移
    dialog = app[u"実行ファイル選択"]

    # 実行ファイル選択
    dialog.TypeKeys('{TAB}')
    dialog.RightClickInput()

    # ソート
    dialog.TypeKeys('{DOWN 2}' '{ENTER}')

    # 選択
    dialog.TypeKeys('{DOWN 21}' '{ENTER}') #DOWN の数で調整
    
    login()   
    try:
        dialog = app.top_window_()
        #file close
        app.Microsoft_Excel.TypeKeys('%f' 'c')
        time.sleep(1)
        app.Microsoft_Excel.TypeKeys('{ENTER}')
        dialog = app[u'名前を付けて保存']
    
        if copyToClipboard(MakePath()):
            #paste path and select .csv
            dialog.TypeKeys('^v' '{TAB}' 'c'
            '%s')
            dialog = app.top_window_()
            #message dialog 
            dialog.TypeKeys('%y')
            dialog = app.top_window_()
            dialog.TypeKeys('{ENTER}')
            dialog = app.top_window_()
            dialog.TypeKeys('%y')
    except:
        print 'output missed...'
        input ()
    
def Main():
    try:
        NWF()
    except Exception:
        print 'somethig is BAD.'
        input()
   
if __name__ == '__main__':
    Main()