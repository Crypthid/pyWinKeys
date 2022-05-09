"""
 Created by Daniel Månsson at 2020-03-01
 File: script_executor.py
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
 Provides a simple script executor class, which utilizes a ScriptReader and an
 ExecutionAPI to perform scripted actions in a sequence.
"""
from typing import List, Tuple, Dict, Callable
from pyWinKeys.scripts.script_reader import ScriptReader
from pyWinKeys.scripts.execution_api import ExecutionAPI
from time import sleep
import sys


class ScriptExecutor:
    # Allowed command strings and the amount of parameters, as well as handler to be defined.
    # All parameters use ',' as delimiter
    _commands: Dict[str, Tuple[int, Callable[..., bool]]] = {'press': (1, ExecutionAPI.press),
                                                             'write': (1, ExecutionAPI.write),
                                                             'move': (2, ExecutionAPI.move),
                                                             'hold_mouse': (1, ExecutionAPI.hold_mouse),
                                                             'release_mouse': (1, ExecutionAPI.release_mouse)}

    def __init__(self, filename: str):
        self.scripts: Dict = ScriptReader.load_scripts(filename)
        return

    def _internal_execute(self, script: List[Tuple[int, str, str]]) -> bool:
        """
        Internal execution loop for a script.
        Thread will be reserved until the entire script has executed (or an error occurs)
        :param script: The interpreted script which has been loaded.
        :return: True if execution was successful, False if it aborted.
        """
        for command in script:
            if len(command) != 3:
                print("ScriptExecutor-internal_execute: Command is of wrong size {0}!"
                      .format(len(command)), file=sys.stderr)
                return False
            # Run delay ms -> s
            sleep(command[0] / 1000)
            # Check if command exists
            if command[1] not in self._commands.keys():
                print("ScriptExecutor-internal_execute: Command \'{0}\' does not exist!"
                      .format(command[1]), file=sys.stderr)
                return False
            # TODO: Add escaping, eg. \\\, -> \, where , not used for split
            parameters: List[str] = command[2].split(",")
            if len(parameters) != self._commands[command[1]][0]:
                print("ScriptExecutor-internal_execute: Command \'{0}\' expects \'{1}\' parameters, got \'{2}\'"
                      .format(command[1], self._commands[command[1]][0], len(parameters)), file=sys.stderr)
                return False
            # Execute command with the parameters
            # print("ScriptExecutor-internal_execute: EXECUTING command \'{0}\'".format(command[1]))
            self._commands[command[1]][1](*parameters)
        return True

    def get_script_names(self) -> tuple[str]:
        return tuple(self.scripts.keys()) if self.scripts is not None else tuple()

    def execute(self, script_name: str) -> bool:
        """
        Executes script 'script_name' which has to exist in the loaded file.
        :param script_name: The name of the script (eg. the --- name tag in the loaded script file).
        :return: True if execution was successful, False otherwise.
        """
        if not self.scripts or script_name not in self.scripts.keys():
            print("ScriptExecutor-execute: No script \'{0}\' loaded!".format(script_name), file=sys.stderr)
            return False
        return self._internal_execute(self.scripts[script_name])
