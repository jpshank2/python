import pyautogui
from PIL import Image
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

print("Press Ctrl-C to quit.")



try:
    while True:
        x, y = pyautogui.position()
        posString = "X: " + str(x).rjust(4) + " Y: " + str(y).rjust(4)
        print(posString, end="")
        print('\b' * len(posString), end='', flush=True)
except KeyboardInterrupt:
    print("\nDone")
