import os, sys, pyautogui, time

def start():
    os.startfile("C:\\windows\\system32\\cmd.exe")
    time.sleep(5.0)
    pyautogui.write("wmic")
    pyautogui.press("enter")

def spbinder():
    start()
    pyautogui.write("""product where name="SPbinder" call uninstall /nointeractive""")
    pyautogui.press("enter")
    time.sleep(5.0)

def novapdf8driver():
    start()
    pyautogui.write("""product where name="novaPDF 8 Printer Driver" call uninstall /nointeractive""")
    pyautogui.press("enter")
    time.sleep(5.0)

def novapdf10driver():
    start()
    pyautogui.write("""product where name="novaPDF SDK 10 Printer Driver" call uninstall /nointeractive""")
    pyautogui.press("enter")
    time.sleep(5.0)

def novapdf1064():
    start()
    pyautogui.write("""product where name="novaPDF SDK 10 SDK COM (x64)" call uninstall /nointeractive""")
    pyautogui.press("enter")
    time.sleep(5.0)

def novapdf1086():
    start()
    pyautogui.write("""product where name="novaPDF SDK 10 SDK COM (x86)" call uninstall /nointeractive""")
    pyautogui.press("enter")
    time.sleep(5.0)

def novapdf864():
    start()
    pyautogui.write("""product where name="novaPDF 8 SDK COM (x64)" call uninstall /nointeractive""")
    pyautogui.press("enter")
    time.sleep(5.0)

def novapdf886():
    start()
    pyautogui.write("""product where name="novaPDF 8 SDK COM (x86)" call uninstall /nointeractive""")
    pyautogui.press("enter")
    time.sleep(5.0)

def end():
    pyautogui.write("exit")
    pyautogui.press("enter")
    pyautogui.write("exit")
    pyautogui.press("enter")

def run():
    spbinder()
    novapdf8driver()
    novapdf10driver()
    novapdf1064()
    novapdf1086()
    novapdf864()
    novapdf886()

run()