from pynput.keyboard import Key, Listener
import pyautogui

def on_press(key):
    global typed
    global listening

    key_str = str(key).replace('\'', '')
    if key_str in starter:
        print('listening')
        typed = []
        listening = True
    
    if listening:
        if key_str.isalpha():
            typed.append(key_str)

        if key == ender:
            print('not listening')
            keyword = ""
            keyword = keyword.join(typed)
            print(keyword)
            listening = False
            if keyword in ['remote', 'Remote', 'remoted', 'Remoted']:
                print('keyword!')
                pyautogui.press('backspace', presses=len(keyword) + 2)
                pyautogui.typewrite('Happy Birthday to Me!')
                listening = False

starter = ['r', 'R']
ender = Key.space
typed = []
listening = True

with Listener(on_press=on_press) as listener:
    listener.join()