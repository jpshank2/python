import subprocess as sp

print("What is the new employee's first name?")
firstName = input()
print("What is the new employee's middle initial?")
middle = input()
print("What is the new employee's last name?")
lastName = input()
# print("What is the new employee's office?")
# office = input()
# groups = []
# print("How many groups are we assigning them?")
# amountOfGroups = int(input())
# for i in range(amountOfGroups):
#     print("What is group " + str(i+1) + "?")
#     groups.append(input())

print("give me a number")
number1 = input()
print("another one")
number2 = input()

userName = (firstName[:1] + lastName).lower()
validUser = False
# fullName = firstName + " " + lastName
# email = userName + "@bmss.com"
#ou = "OU=Users,OU=" + office + ",OU=BMSSMAIN,DC=bmss,DC=com"

# print(userName + " " + fullName)
#p = sp.Popen(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", "Set-ExecutionPolicy -ExecutionPolicy Unrestricted | . \"C:\\Users\\jeremyshank\\Documents\\BMSS Assets\\Code\\PowerShell\\adding.ps1\";", "Add-Numbers " + number1 + " " + number2], stdout=sp.PIPE, stdin=sp.PIPE, stderr=sp.PIPE)

def checker(userName):
    p = sp.Popen(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", "$User = $(try {Get-ADUser " + userName + "} catch {$null}); if($User -eq $null) {return 0} else {return 1}"], stdout=sp.PIPE, stdin=sp.PIPE, stderr=sp.PIPE)
    output = p.communicate()
    if output == r"b'1\r\n'":
        return {valid: False, userName: (firstName[:1] + middle + lastName).lower()}

while not validUser:
    if checker(userName) == r"b'1\r\n'":
        userName = (firstName[:1] + middle + lastName).lower()
        checker(userName)
    

