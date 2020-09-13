from pynput import keyboard

# 全局已经按下的按键
pressed = set([])


def on_press(key):
    try:
        # print('alphanumeric key {0} pressed'.format(key.char))
        pressed.add(chr(key.vk))
    except AttributeError:
        # print('special key {0} pressed'.format(key))
        pressed.add(key.name)
    # print(key)
    print(pressed)


def on_release(key):
    # print('{0} released'.format(key))
    try:
        # print('alphanumeric key {0} released'.format(key.char))
        if chr(key.vk) in pressed:
            pressed.remove(chr(key.vk))
    except AttributeError:
        # print('special key {0} released'.format(key))
        if key.name in pressed:
            pressed.remove(key.name)
    # print(key)
    # print(pressed)


while True:
    with keyboard.Listener(
            on_press=on_press,
            on_release=on_release)as listener:
        listener.join()
