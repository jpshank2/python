import subprocess, os

subprocess.run("taskkill /f /im outlook.exe")

subprocess.run("""wmic product where name="GoFileRoom Client Add-In" call uninstall /nointeractive""")
os.startfile(r"c:\it fixes\gofileroom.bat")