import os, sys, pyautogui, time

def sophos():
    os.startfile("C:\\windows\\system32\\cmd.exe")
    time.sleep(1.0)
    pyautogui.write(""""C:\\Program Files\\Sophos\\Sophos Endpoint Agent\\uninstallcli.exe" """)
    pyautogui.press("enter")
    time.sleep(3.0)

def sentinel():
    os.startfile(r"\\fs01\Shares\Programs\Sentinel One\SentinelAgent_windows_v3_7_2_45.exe")

def start():
    sophos()
    sentinel()

start()