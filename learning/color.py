import pyautogui
from PIL import Image
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

im = pyautogui.screenshot()

coordinate = x, y = 358, 269

print(str(im.getpixel(coordinate)))