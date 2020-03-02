"""
 Created by Daniel Månsson at 2020-03-01
 File: script_reader.py
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
 Simple script reader for reading a primitive automation script into easy-to-execute
 Dict-structured scripts.

 The primary purpose of this module is to create test scripts for pyWinKeys, but it also
 works perfectly fine for creating simple automation tasks that can be directly and repeatedly executed inside
 a Python environment.
 -----------------------------------------------------------------------------
 TODO: Add support for initial script values (specified by the reserved % tag)
"""
from typing import Optional, Tuple
import os
import string
import sys


class ScriptReader:
    @staticmethod
    def _contains_prefix(prefix: str, line: str) -> bool:
        # Sanity check
        if len(prefix) > len(line) or len(line) == 0:
            return False
        for i in range(0, len(prefix)):
            if prefix[i] != line[i]:
                return False
        return True

    @staticmethod
    def _remove_obsolete_characters(mark_char: str, rm_chars: Tuple[str, ...], line: str) -> str:
        """
        Purges a line of text from all rm_char characters that are not inside a group marked by mark_char characters.
        :param mark_char: MUST be of length 1 and only one (Character to match start/end of a group).
        :param rm_chars: MUST be a tuple consisting of strings of length 1 (Characters to remove).
        :param line: input string.
        :return: Cleaned string.
        """
        output: str = ""
        found: bool = False
        for i in range(0, len(line)):
            if line[i] == mark_char:
                # Make sure that it is not escaped!
                # TODO: Add smarter escaping (eg. read \ before interpreting the following character e.g \\\"->\")
                if i != 0 and line[i - 1] != '\\':
                    found = not found
            if found or (not found and line[i] not in rm_chars):
                output += line[i]
        return output

    @staticmethod
    def load_scripts(filename: str) -> Optional[dict]:
        """
        Loads all scripts from the given file located in the same folder or from a sub-folder.

        The loaded script data has the following format, with the list being ordered:
        {'script name': [(delay, cmd, "parameters"), (delay, cmd, "parameters"), ...]
        ,
        'script name2':...
        ,
        ...}
        :param filename: Name of the file containing the scripts.
        :return: Any scripts from the given file, or None if the file does not exist or if it was empty.
        """
        scripts: dict = {}
        if not os.path.isfile(filename):
            return None

        s_name: str = ""  # Current script name
        with open(filename, 'r') as f:
            for line_n, line in enumerate(f, 1):
                # Remove all whitespaces appearing outside parameters, eg outside " " groups
                temp: str = ScriptReader._remove_obsolete_characters('\"', string.whitespace, line)
                # Traverse til we find a script
                if not s_name:
                    if ScriptReader._contains_prefix("---", temp):
                        if len(temp) < 4:
                            print("load_scripts: ERROR: script at line {0} has no name!"
                                  .format(line_n), file=sys.stderr)
                            return None
                        s_name = temp[3:]
                        if s_name in scripts.keys():
                            print("load_scripts: ERROR: dict already contains a script with name {0}!"
                                  .format(s_name), file=sys.stderr)
                            return None
                        # print("load_scripts: READING script \'{0}\'".format(s_name))
                        scripts[s_name] = list()
                # Otherwise, try to read commands in script until we reach the end
                else:
                    # Ignore empty rows and comment rows (and initial values % as well for now)
                    if len(temp) == 0 or temp[0] in ('#', '%'):
                        continue
                    # Close script if an exit tag is found
                    elif ScriptReader._contains_prefix("---", temp):
                        # print("load_scripts: FINISHING script \'{0}\'".format(s_name))
                        s_name = ""
                        continue
                    first_split = temp.find(',')
                    second_split = temp.find(',', first_split + 1)
                    first_param = temp.find('\"', second_split + 1)
                    second_param = temp.find("\"", first_param + 1)
                    if first_split <= 0 or second_split <= 0 or first_param <= 0 or second_param <= 0 or \
                            second_split <= first_split or first_param <= second_split or second_param <= first_param:
                        print("load_scripts: ERROR, invalid split indexes {0}, {1}, {2}, {3} string [{4}]"
                              .format(first_split, second_split, first_param, second_param, temp)
                              , file=sys.stderr)
                        return None
                    delay_str = temp[:first_split]
                    cmd_str = temp[first_split + 1:second_split]
                    param_str = temp[first_param + 1:second_param]
                    # Get start and end of parameters (in case of comments)
                    if not delay_str.isdigit():
                        print("load_scripts: ERROR, delay not an integer \'{0}\'".format(delay_str)
                              , file=sys.stderr)
                        return None
                    # Add command to list, it's up to the command executing mechanism to test command eligibility
                    scripts[s_name].append((int(delay_str), cmd_str, param_str))
                    # print("load_scripts: SUCCESS, read delay={0}, cmd_str=\'{1}\' params=\'{2}\'"
                    #      .format(int(delay_str), cmd_str, param_str))
        return scripts
