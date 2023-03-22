import pyperclip
import pyautogui as pag
import time
import openpyxl
import os
import glob
from pathlib import Path
"""
aa= pag.position()
print(aa)
"""
x = 1540
y = 290
s = 0

#  座標

#  英語  x 406 y 375
#  ベトナム語 418 394
#　インドネシア語 421 520
#　とりあえずこのコードはスライド一項目につきるものである。


def pushget():
    x = 1540
    y = 290
    pag.click(x, y, duration=0.5) 
    time.sleep(0.8) #編集画面入るまでの時間
    pag.click(1843,325,duration = 0.5) # Translation 選択
    pag.click(454,271,duration=1) # 言語リスト選択
    pag.click(370,372,duration=1) # 翻訳する言語位置選択
    pag.doubleClick(360,372,duration=1)
    pag.hotkey('ctrl','a',duration=1)
    pag.press('del')
    pag.hotkey('ctrl', 'v')
    pag.click(975,967,duration=0.5)
    pag.click(975,967,duration = 0.5)
    pag.click(781,953,duration = 0.5)
    time.sleep(0.5)
    pag.click(942,961,duration = 0.5)
    time.sleep(2)
    pag.click(1149,952,duration = 0.5)
    time.sleep(2.8)
    pag.click(82,442,duration = 0.5)
    pag.click(70,118,duration = 0.5)
    time.sleep(1)
    pag.click(57,101,duration = 0.5)
    y = y+33
   


def MouseMove():
    pag.moveTo(1540,290,duration=0.5)#default
    pag.moveTo(1540,323,duration=0.5) #33
    pag.moveTo(1540,357,duration=0.5)
    pag.moveTo(1540,392,duration=0.5)
    pag.moveTo(1540,427,duration=0.5)
    pag.moveTo(1540,459,duration=0.5)
    pag.moveTo(1540,491,duration=0.5)
    pag.moveTo(1540,523,duration=0.5)
    pag.moveTo(1540,555,duration=0.5)
    pag.moveTo(1540,duration=0.5)
    pag.moveTo(1540,430,duration=0.5)
    pag.moveTo(1540,430,duration=0.5)



file_path1 = "C:\\Users\\ウヨン\\Downloads\\介護フレーズ_0315_翻訳入り_v4.xlsx"

wb = openpyxl.load_workbook(file_path1,data_only=True)
for ws in wb.worksheets:
    start_row = 4
    start_column = 6
    dif_col = 3
    vt_col = 10
    quote = []
    B = "B:"
    A = "A:"
    while not ws.cell(start_row , start_column).value is None:
        #print(ws.cell(start_row,vt_col).value)
        #if (ws.cell(row = start_row , column = start_column).value ==  ws.cell(row = start_row + 1 , column = start_column).value):
        quote.extend([B+ws.cell(start_row,start_column).value , ws.cell(start_row,vt_col).value , A+ws.cell(start_row,start_column).value , ws.cell(start_row,vt_col).value])
        
        if (ws.cell(row = start_row , column = dif_col).value !=  ws.cell(row = start_row + 1 , column = dif_col).value):
            #print("-----------------------")
            quoteline = "\n".join(quote)
            pyperclip.copy(quoteline)
            Paste = pyperclip.paste()
            #print(Paste)

            pushget()


            #i = i+1
          
            quote = []
        start_row += 1
    
    wb.save(file_path1) 

