import pyperclip
import pyautogui as pag
import time
import openpyxl
import os
import glob
from pathlib import Path
"""
#インドネシア語 421 520
def pushget(xx,yy,zz):
    x = 1540
    y = 290
    pag.click(x, y, duration=0.5) 
    time.sleep(0.8) #編集画面入るまでの時間
    pag.click(1843,325,duration = 0.5) # Translation 選択
    pag.click(454,271,duration=1) # 言語リスト選択
    pag.click(xx,yy,duration=1) # 翻訳する言語位置選択
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
pushget(421,520,33)

"""
def testplus(x,y):
    i = 0
    global ii
    x = 20
    xx= x + ii
    print(xx)
        
ii = 1
testplus(5,6)
