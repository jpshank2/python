#! python3

import pyautogui
from threading import Timer
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

def navigate():
    width, height = pyautogui.size()
    pyautogui.click(((width / 4) + 40), (height - 15))
    pyautogui.hotkey("alt", "d")
    pyautogui.typewrite("C:\\Users\\jeremyshank\\Desktop\\test")
    pyautogui.press('enter')

def selectTop():
    width, height = pyautogui.size()
    pyautogui.click((width / 4), 400)
    pyautogui.press("down")
    pyautogui.press("up")

def unzip():
    pyautogui.hotkey("shift", "f10")
    pyautogui.press("7")
    pyautogui.press("down")
    pyautogui.press("down")
    pyautogui.press("enter")
    pyautogui.hotkey("alt", "d")
    pyautogui.typewrite("C:\\Users\\jeremyshank\\Desktop\\result")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("space")
    pyautogui.press("enter")

def changeScreen():
    pyautogui.keyDown("alt")
    pyautogui.keyDown("shift")
    pyautogui.press("tab")
    pyautogui.keyUp("shift")
    pyautogui.press("tab")
    pyautogui.keyUp("alt")

def delete():
    width, height = pyautogui.size()
    pyautogui.hotkey("shift", "f10")
    pyautogui.press("d")
    if pyautogui.pixelMatchesColor(358, 269, (205, 42, 80)):
        pyautogui.click(((width / 4) + 40), (height - 15))
        runMe() 
    else:
        print("finished!")

def checker():
    coordinate = x, y = 1104, 629   
    im = pyautogui.screenshot()
    if pyautogui.pixelMatchesColor(1104, 629, (240, 240, 240)):
        t = Timer(45.0, checker)
        t.start()
    else: 
        selectTop()
        delete()

def runMe():
    try:
        navigate()
        selectTop()
        unzip()
        changeScreen()
        checker()
    except KeyboardInterrupt:
        print("Stopped")

runMe()