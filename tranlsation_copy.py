import pyautogui as pag
import time

def paste_translate(x):
    i= 1
    x
    yy = 33
    while i <= x:
        pag.click(x=1536,y=257+yy, duration=0.5) #鉛筆ボタン
        pag.click(781,953,duration = 0.5) #save
        time.sleep(0.4)
        pag.click(942,961,duration = 0.5) #compile
        time.sleep(1.8)
        pag.click(1149,952,duration = 0.5) #run
        time.sleep(3)
        pag.click(70,118,duration = 0.5)#戻る
        time.sleep(0.5)
        pag.click(57,101,duration = 0.5)#戻る
        i = i+1
        yy = yy+33

#paste_translate(9)
print(pag.position())