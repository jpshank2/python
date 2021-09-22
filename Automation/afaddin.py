import subprocess, os

subprocess.run("taskkill /f /im outlook.exe")

subprocess.run("""wmic product where name="AdvanceFlow Client Add-Ins" call uninstall /nointeractive""")
os.startfile(r"\\fs01\Shares\Data Drive\AdvanceFlow\AdvanceFlowClient32.msi")