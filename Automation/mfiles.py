import subprocess, os

subprocess.run("taskkill /f /im outlook.exe")

subprocess.run("""wmic product where name="M-Files 2018" call uninstall /nointeractive""")
os.startfile(r"\\fs01\shares\programs\m-files\M-Files_client.msi")