import pyautogui, os, time

def installer():
    os.startfile(r"\\fs01\shares\it\it fixes\Fortinet_CA_SSL.cer")
    time.sleep(1.0)

def changeScreen():
    pyautogui.keyDown("alt")
    pyautogui.keyDown("shift")
    pyautogui.press("tab")
    pyautogui.keyUp("shift")
    pyautogui.press("tab")
    pyautogui.keyUp("alt")
    time.sleep(1.0)

def certificate():
    pyautogui.press("tab")
    pyautogui.press("enter")
    time.sleep(1.0)

def certImportWizard():
    pyautogui.press("down")
    pyautogui.press("enter")
    time.sleep(1.0)

def certStore():
    pyautogui.press("down")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")
    time.sleep(1.0)
    pyautogui.press("down")
    pyautogui.press("enter")
    pyautogui.press ("tab")
    pyautogui.press("enter")
    pyautogui.press("enter")

def runIt():
    installer()
    changeScreen()  
    certificate()
    certImportWizard()
    certStore()

runIt()