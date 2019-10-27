import os

"""
Provides a simple API for performing mouse and keyboard operations in Python for Windows.
This implementation using the ctypes library does not require the application itself to be in focus when performing its
operations and can therefore be ran in the background.
Supports:
- Clicking/Holding mouse buttons
- Moving the mouse relative to its current position or absolute coordinates in the primary monitors resolution
- Performing mouse scrolling operations in any directions
- Performing keyboard button presses (including media buttons, see _win_key dict initialization)
- Perform multi-key keyboard combinations
- Retrieving the current mouse position
Does not (really) support:
- Writing capitalized letters (can be done by toggling 'caps') 
- Writing special symbols (can be made by holding down the correct modifier buttons)
"""

timeout: int = 5  # ms

if os.name == "nt":
    import ctypes
    import enum
    import string
    import time
    import typing

    '''
    Setup all types required to communicate with the user32 API
    '''

    c_PUL = ctypes.POINTER(ctypes.c_ulong)


    class CPoint(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        pass


    class CMouseInput(ctypes.Structure):
        _fields_ = [("dx", ctypes.c_long),
                    ("dy", ctypes.c_long),
                    ("mouseData", ctypes.c_ulong),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", c_PUL)]


    class CKeyBdInput(ctypes.Structure):
        _fields_ = [("wVk", ctypes.c_ushort),
                    ("wScan", ctypes.c_ushort),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", c_PUL)]


    class CHardwareInput(ctypes.Structure):
        _fields_ = [('uMsg', ctypes.c_ulong),
                    ('wParamL', ctypes.c_long),
                    ('wParamH', ctypes.c_long)]


    class _CInputUnion(ctypes.Union):
        _fields_ = [('mi', CMouseInput),
                    ('ki', CKeyBdInput),
                    ('hi', CHardwareInput)]


    class CInput(ctypes.Structure):
        _fields_ = [('type', ctypes.c_ulong),
                    ('ctypes.Union', _CInputUnion)]


    class MouseKey(enum.IntEnum):
        RIGHT_BUTTON = 0,
        LEFT_BUTTON = 1,
        MIDDLE_BUTTON = 2


    class MouseScrollDirection(enum.IntEnum):
        DOWN = 0,
        UP = 1,
        LEFT = 2,
        RIGHT = 3


    '''
    Setup 'defines'/masks
    '''
    # Input types
    INPUT_MOUSE = 0x00
    INPUT_KEYBOARD = 0x01
    INPUT_HARDWARE = 0x02
    # Key events
    KEY_EVENT_EXTENDED_KEY = 0x0001
    KEY_EVENT_RELEASE = 0x0002
    KEY_EVENT_UNICODE = 0x0004
    KEY_EVENT_SCAN_CODE = 0x0008
    # Mouse events (https://docs.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-mouseinput)
    MOUSE_WHEEL_DELTA = 120  # Value per scroll tick
    MOUSE_EVENT_MOVE = 0x0001  # Mouse is moving
    MOUSE_EVENT_ABSOLUTE = 0x8000  # The movement of the mouse is in absolute coordinates
    MOUSE_EVENT_HWHEEL = 0x0100  # The mouse scroll wheel movement was made horizontally
    MOUSE_EVENT_WHEEL = 0x0800  # The mouse scroll wheel was moved
    MOUSE_EVENT_CLICK_PRESS = 0x0080
    MOUSE_EVENT_CLICK_RELEASE = 0x0100
    MOUSE_EVENT_LEFT_PRESS = 0x0002
    MOUSE_EVENT_LEFT_RELEASE = 0x0004
    MOUSE_EVENT_RIGHT_PRESS = 0x0008
    MOUSE_EVENT_RIGHT_RELEASE = 0x0010
    MOUSE_EVENT_MIDDLE_PRESS = 0x0020
    MOUSE_EVENT_MIDDLE_RELEASE = 0x0040
    MOUSE_EVENT_VIRTUAL_DESK = 0x4000  # Send this to map mouse events to the entire screen area (multi-monitor)
    MOUSE_EVENT_COORDINATES = 65535  # The amount of coordinates in x and y used by mouse input

    _SCREEN_WIDTH_PX: int = 0
    _SCREEN_HEIGHT_PX: int = 0
    _SCREEN_X_MULTIPLIER: int = 0
    _SCREEN_Y_MULTIPLIER: int = 0


    def refresh_monitor_size() -> None:
        """
        If the primary monitor resolution has changed, this function
        needs to be called to update the mapping for mouse events
        """
        global _SCREEN_HEIGHT_PX
        global _SCREEN_WIDTH_PX
        global _SCREEN_X_MULTIPLIER
        global _SCREEN_Y_MULTIPLIER
        _SCREEN_WIDTH_PX = ctypes.windll.user32.GetSystemMetrics(0)
        _SCREEN_HEIGHT_PX = ctypes.windll.user32.GetSystemMetrics(1)
        _SCREEN_X_MULTIPLIER = MOUSE_EVENT_COORDINATES // _SCREEN_WIDTH_PX
        _SCREEN_Y_MULTIPLIER = MOUSE_EVENT_COORDINATES // _SCREEN_HEIGHT_PX


    # Set initial resolution of primary monitor for mouse event mappings
    refresh_monitor_size()

    def get_primary_resolution() -> typing.Tuple[int, int]:
        return _SCREEN_WIDTH_PX, _SCREEN_HEIGHT_PX

    '''
    Map characters and special keys to the correct VK keys for KeyBdInput
    '''
    _win_key = {}  # Key code fetch dict (Keyboard)


    def _add_sequence_keys(key_dict: dict, keys: typing.Union[str, typing.Tuple[str, ...]], start: int) -> None:
        for key, value in zip(keys, range(start, start + len(keys))):
            key_dict[key] = value


    # Load numbers and letters and functions keys
    _add_sequence_keys(_win_key, string.digits, 0x30)
    # Use unicode for all letters
    _add_sequence_keys(_win_key, string.ascii_lowercase, 0x41)
    # Can't really differentiate between lower and upper for VK (if we don't use shift modifiers...)
    _add_sequence_keys(_win_key, string.ascii_uppercase, 0x41)
    _add_sequence_keys(_win_key, ('F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10',
                                  'F11', 'F12', 'F13', 'F14', 'F15', 'F16', 'F17', 'F18',
                                  'F19', 'F20', 'F21', 'F22', 'F23', 'F24'), 0x70)
    # Manually input special keys
    _win_key['backspace'] = 0x08
    _win_key['tab'] = 0x09
    _win_key['enter'] = 0x0D
    _win_key['shift'] = 0x10
    _win_key['ctrl'] = 0x11
    _win_key['alt'] = 0x12
    _win_key['pause'] = 0x13
    _win_key['caps'] = 0x14
    _win_key['esc'] = 0x1B
    _win_key['space'] = 0x20
    _win_key['page-up'] = 0x21
    _win_key['page-down'] = 0x22
    _win_key['end'] = 0x23
    _win_key['home'] = 0x24
    _win_key['left'] = 0x25
    _win_key['up'] = 0x26
    _win_key['right'] = 0x27
    _win_key['down'] = 0x28
    _win_key['select'] = 0x29
    _win_key['print'] = 0x2A
    _win_key['execute'] = 0x2B
    _win_key['prtscn'] = 0x2C
    _win_key['insert'] = 0x2D
    _win_key['delete'] = 0x2E
    _win_key['win'] = 0x5B  # Left windows key
    _win_key['vol-mute'] = 0xAD
    _win_key['vol-down'] = 0xAE
    _win_key['vol-up'] = 0xAF
    _win_key['media-next'] = 0xB0
    _win_key['media-prev'] = 0xB1
    _win_key['media-stop'] = 0xB2
    _win_key['media_pp'] = 0xB3  # VK_MEDIA_PLAY_PAUSE


    def _get_hex_code(key: str) -> int:
        """
        Retrieves the VK for the given key
        :param key:
        :return: (key_code) None if the character does not exist.
        """
        return _win_key.get(key, None)


    def _sleep(ms: int) -> None:
        time.sleep(ms / 1000.0)


    def _send_input(c_input: CInput) -> None:
        ctypes.windll.user32.SendInput(1, ctypes.byref(c_input), ctypes.sizeof(c_input))


    def mouse_move(x: int, y: int, relative: bool) -> bool:
        """
        Moves the mouse the given x, y position (in pixels) on the main monitor.
        A relative values means that x, y is treated as dx, dy to the current mouse pointer position.
        """
        if x < 0 or x > _SCREEN_WIDTH_PX or y < 0 or y > _SCREEN_HEIGHT_PX:
            return False
        flags = MOUSE_EVENT_MOVE
        if not relative:
            flags |= MOUSE_EVENT_ABSOLUTE
        inp_struct = CInput(ctypes.c_ulong(INPUT_MOUSE),
                            _CInputUnion(mi=CMouseInput(dx=x * _SCREEN_X_MULTIPLIER, dy=y * _SCREEN_Y_MULTIPLIER,
                                                        dwFlags=flags)))
        _send_input(inp_struct)
        return True


    def mouse_hold(mouse_btn: MouseKey) -> bool:
        flags = 0
        if mouse_btn is MouseKey.LEFT_BUTTON:
            flags |= MOUSE_EVENT_LEFT_PRESS
        elif mouse_btn is MouseKey.RIGHT_BUTTON:
            flags |= MOUSE_EVENT_RIGHT_PRESS
        elif mouse_btn is MouseKey.MIDDLE_BUTTON:
            flags |= MOUSE_EVENT_MIDDLE_PRESS
        else:
            return False
        mouse_inp: CInput = CInput(ctypes.c_ulong(INPUT_MOUSE), _CInputUnion(mi=CMouseInput(dwFlags=flags)))
        _send_input(mouse_inp)
        return True


    def mouse_release(mouse_btn: MouseKey) -> bool:
        flags = 0
        if mouse_btn is MouseKey.LEFT_BUTTON:
            flags |= MOUSE_EVENT_LEFT_RELEASE
        elif mouse_btn is MouseKey.RIGHT_BUTTON:
            flags |= MOUSE_EVENT_RIGHT_RELEASE
        elif mouse_btn is MouseKey.MIDDLE_BUTTON:
            flags |= MOUSE_EVENT_MIDDLE_RELEASE
        else:
            return False
        mouse_inp: CInput = CInput(ctypes.c_ulong(INPUT_MOUSE), _CInputUnion(mi=CMouseInput(dwFlags=flags)))
        _send_input(mouse_inp)
        return True


    def mouse_press(mouse_btn: MouseKey) -> bool:
        if not mouse_hold(mouse_btn):
            return False
        _sleep(timeout)
        if not mouse_release(mouse_btn):
            return False
        return True


    def mouse_scroll(ticks: int, direction: MouseScrollDirection) -> bool:
        d_scroll: int = ticks * 120
        if direction is MouseScrollDirection.DOWN:
            mouse_inp = CInput(ctypes.c_ulong(INPUT_MOUSE),
                               _CInputUnion(mi=CMouseInput(mouseData=-d_scroll, dwFlags=MOUSE_EVENT_WHEEL)))
        elif direction is MouseScrollDirection.UP:
            mouse_inp = CInput(ctypes.c_ulong(INPUT_MOUSE),
                               _CInputUnion(mi=CMouseInput(mouseData=d_scroll, dwFlags=MOUSE_EVENT_WHEEL)))
        elif direction is MouseScrollDirection.LEFT:
            mouse_inp = CInput(ctypes.c_ulong(INPUT_MOUSE),
                               _CInputUnion(
                                   mi=CMouseInput(mouseData=-d_scroll, dwFlags=MOUSE_EVENT_WHEEL | MOUSE_EVENT_HWHEEL)))
        elif direction is MouseScrollDirection.RIGHT:
            mouse_inp = CInput(ctypes.c_ulong(INPUT_MOUSE),
                               _CInputUnion(
                                   mi=CMouseInput(mouseData=d_scroll, dwFlags=MOUSE_EVENT_WHEEL | MOUSE_EVENT_HWHEEL)))
        else:
            return False
        _send_input(mouse_inp)
        return True


    def mouse_get_xy() -> typing.Tuple[int, int]:
        pt = CPoint()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        return pt.x, pt.y


    def _keyboard_hold(key_code: int) -> None:
        key_inp = CInput(ctypes.c_ulong(INPUT_KEYBOARD), _CInputUnion(ki=CKeyBdInput(wVk=key_code)))
        _send_input(key_inp)


    def _keyboard_release(key_code: int) -> None:
        key_inp = CInput(ctypes.c_ulong(INPUT_KEYBOARD),
                         _CInputUnion(ki=CKeyBdInput(wVk=key_code, dwFlags=KEY_EVENT_RELEASE)))
        _send_input(key_inp)


    def keyboard_press(key: str) -> bool:
        """
        Press the given keyboard key, if successful return true otherwise return false
        """
        key_code = _get_hex_code(key)
        if key_code is None:
            return False
        _keyboard_hold(key_code)
        _sleep(timeout)
        _keyboard_release(key_code)
        return True


    def keyboard_press_combo(key_set: str) -> bool:
        """
        Executes a combination of keys delimited by +, for example 'ctrl + alt + delete'. Spaces are ignored.
        See the initialization of _win_key to see the (case-sensitive) names required.
        """
        keys = key_set.replace(" ", "").split("+")
        handled_keys = set()  # Set to make sure we don't try to do something stupid like "Ctrl + Ctrl + A"
        key_codes: typing.List[int, ...] = []
        for key in keys:
            key_code = _get_hex_code(key)
            if key_code is None or key_code in handled_keys:
                return False
            handled_keys.add(key_code)
            key_codes.append(key_code)
        # Now we got a list of keys that we can handle
        # Press in sequence
        for key_code in key_codes:
            _keyboard_hold(key_code)
        _sleep(timeout)
        # Release in reversed sequence
        for key_code in reversed(key_codes):
            _keyboard_release(key_code)
        return True
else:
    raise NotImplementedError("The winkeys API is only support for Windows!")
