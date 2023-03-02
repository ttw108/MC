import time
import autoit as at
import pyautogui as ag
from tkinter import *
import tkinter as tk
from tkinter import simpledialog
root = Tk()
print(root.tk.exprstring('$tcl_library'))
print(root.tk.exprstring('$tk_library'))

root = tk.Tk()
root.withdraw()
# Create an input dialog
wk_flow = simpledialog.askstring("Input", "有多少個工作底稿",
                             parent=root)
wk_cht = simpledialog.askstring("Input", "每個底稿有多少個圖表視窗？",
                             parent=root)

wk_flow=int(wk_flow)+1
wk_cht=int(wk_cht)+1

at.opt("wintitlematchmode",2)
title=at.win_get_title("[CLASS:ATL_MCMDIMainFrame]")
at.win_activate(title)
at.win_wait_active(title)

agwin=ag.getWindowsWithTitle(title)
ag.moveTo( agwin[0].left+400,agwin[0].top+20)
ag.click( agwin[0].left+400,agwin[0].top+20)
time.sleep(1)
at.send("^{PGUP}"*8)
time.sleep(5)
ag.moveTo( agwin[0].left+400,agwin[0].top+20)
ag.click( agwin[0].left+400,agwin[0].top+20)
ag.moveTo(agwin[0].left + 400, agwin[0].top + 120)
ag.click(agwin[0].left + 400, agwin[0].top + 120)

for ix in range(1,wk_flow):
    title = at.win_get_title("[CLASS:ATL_MCMDIMainFrame]")
    at.win_activate(title)
    at.win_wait_active(title)
    agwin = ag.getWindowsWithTitle(title)
    time.sleep(1)
    for i in range(1,wk_cht):
        at.win_activate(title)
        at.win_wait_active(title)
        at.send("!v")
        time.sleep(1)
        at.send("{up 6}")
        time.sleep(1)
        at.send("{enter}")
        time.sleep(1)
        at.opt("wintitlematchmode", 2)
        reporttitle = at.win_get_title("[CLASS:REPORT_WINDOW]")
        time.sleep(3)
        at.win_activate(reporttitle)
        at.win_wait_active(reporttitle)
        #at.win_wait_active(reporttitle)
        #wincod = at.control_get_pos(reporttitle,"","[CLASS:XPFlatToolbar]")
        wincod=ag.getWindowsWithTitle(reporttitle)
        x=wincod[0].left
        y = wincod[0].top
        r= wincod[0].right
        ag.click(x+105,y+47)
        time.sleep(1)
        at.send("{enter}")
        time.sleep(360)
        active_window = ag.getActiveWindow()
        act_r=active_window.right
        act_y=active_window.top
        print(act_r,act_y)
        ag.click(act_r -15, act_y + 5)
        time.sleep(2)
        at.win_activate(title)
        time.sleep(2)
        #ag.moveTo(x + 935, y + 10)
        #ag.click(x + 935, y + 10)
        #time.sleep(1)
        print(x, y)
        at.win_activate(reporttitle)
        at.win_wait_active(reporttitle)
        time.sleep(1)
        active_window1 = ag.getActiveWindow()
        act_r=active_window1.right
        act_y=active_window1.top
        time.sleep(1)
        ag.moveTo(act_r -30 , act_y + 15)
        time.sleep(1)
        ag.click(act_r -30 , act_y + 15)
        time.sleep(1)
        ag.moveTo(agwin[0].left + 400, agwin[0].top + 20)
        ag.click(agwin[0].left + 400, agwin[0].top + 20)
        time.sleep(1)
        at.send("^{tab}")
        time.sleep(5)
        ag.moveTo(agwin[0].left + 400, agwin[0].top + 20)
        ag.click(agwin[0].left + 400, agwin[0].top + 20)

    at.win_activate(title)
    at.send("^{PGDN}")
    time.sleep(10)





