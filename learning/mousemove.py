import pyautogui
from pynput.mouse import Listener
pyautogui.FAILSAFE = True


def on_move(x, y):
    xPos, yPos = pyautogui.position()
    pyautogui.moveTo(xPos + 10, yPos + 10, duration=0.5)
def on_click(x, y, button, pressed):
    if pressed:
        if "{2}".format(x, y, button) == "Button.right":
            listener.stop()
with Listener(on_move=on_move, on_click=on_click) as listener:
    listener.join()
