import openpyxl
import pyperclip
import pyautogui
 
def autozada(x,y):
    pyperclip.copy(x)
    pyautogui.doubleClick(y,duration=0.1)
    pyautogui.hotkey("ctrl","v")