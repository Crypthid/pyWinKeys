"""
 Created by Daniel Månsson at 2020-03-01
 File: execution_api.py
 -----------------------------------------------------------------------------
 MIT License

 Copyright (c) 2020 Daniel Månsson

 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:

 The above copyright notice and this permission notice shall be included in all
 copies or substantial portions of the Software.

 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 SOFTWARE.
 -----------------------------------------------------------------------------
 API towards the pyWinKeys module.
 Allows execution of keyboard and mouse operations by the use of string input.

 TODO: Allow 'write' to output upper-case characters
 TODO: Allow relative coordinates to be set as a mode in the script
 TODO: Allow scrolling actions to be performed in a script
 TODO: Maybe (?) add keyboard key holding/releasing support
"""
import winkeys.winkeys as pyw
import sys
from typing import Union


class ExecutionAPI:
    @staticmethod
    def press(sequence: str):
        """
        Presses a keyboard sequence and releases it afterwards.
        :param sequence: The keyboard sequence to press and release.
        :return: True if successful, False otherwise.
        """
        return pyw.keyboard_press_combo(sequence)

    @staticmethod
    def write(sequence: str):
        """
        Write a sequence on the keyboard.
        :param sequence: The keyboard sequence to write.
        :return: True if all letters were written, False if a character couldn't be resolved.
        """
        for c in sequence:
            if not pyw.keyboard_press(c):
                print("ExecutionAPI-write: Could not write character \'{0}\' in sequence \'{1}\'!"
                      .format(c, sequence), file=sys.stderr)
                return False
        return True

    @staticmethod
    def move(x: str, y: str):
        """
        Move the mouse to the given x, y coordinates of the main monitor.
        :param x: X-coordinate in pixels.
        :param y: Y-coordinate in pixels.
        :return: True if a valid successful move, False otherwise.
        """
        if not x.isdigit() or not y.isdigit():
            print("ExecutionAPI-move: X or Y is not an integer x:\'{0}\' y:\'{1}\' !".format(x, y), file=sys.stderr)
            return
        return pyw.mouse_move(int(x), int(y), False)

    @staticmethod
    def hold_mouse(key: str):
        """
        Hold the given mouse button until a release event is performed.
        :param key: The mouse key to hold.
        :return: True if a valid key was held, False otherwise.
        """
        key_code: Union[pyw.MouseKey, None] = None
        if key.lower() == "right":
            key_code = pyw.MouseKey.RIGHT_BUTTON
        elif key.lower() == "middle":
            key_code = pyw.MouseKey.MIDDLE_BUTTON
        elif key.lower() == "left":
            key_code = pyw.MouseKey.LEFT_BUTTON
        if key_code is None:
            print("ExecutionAPI-hold_mouse: Invalid button \'{0}\'!".format(key), file=sys.stderr)
            return False
        return pyw.mouse_hold(key_code)

    @staticmethod
    def release_mouse(key: str):
        """
        Release the given mouse button.
        :param key: The mouse key to release.
        :return: True if a valid key was released, False otherwise.
        """
        key_code: Union[pyw.MouseKey, None] = None
        if key.lower() == "right":
            key_code = pyw.MouseKey.RIGHT_BUTTON
        elif key.lower() == "middle":
            key_code = pyw.MouseKey.MIDDLE_BUTTON
        elif key.lower() == "left":
            key_code = pyw.MouseKey.LEFT_BUTTON
        if key_code is None:
            print("ExecutionAPI-release_mouse: Invalid button \'{0}\'!".format(key), file=sys.stderr)
            return False
        return pyw.mouse_release(key_code)

    @staticmethod
    def _hold_keyboard(key: str):
        """
        Holding is dangerous, make sure to release when using in a script.
        :param key: The key to hold.
        :return: False if no valid key code exists, True otherwise.
        """
        key_code = pyw._get_hex_code(key)
        if key_code is None:
            print("ExecutionAPI-hold: key \'{0}\' has no valid key code!".format(key), file=sys.stderr)
            return False
        pyw._keyboard_hold(key_code)
        return True

    @staticmethod
    def _release_keyboard(key: str):
        """
        Releases a keyboard key, make sure to release when using hold in a script.
        :param key: The key to release.
        :return: False if no valid key code exists, True otherwise.
        """
        key_code = pyw._get_hex_code(key)
        if key_code is None:
            print("ExecutionAPI-release: key \'{0}\' has no valid key code!".format(key), file=sys.stderr)
            return False
        pyw._keyboard_release(key_code)
        return True
