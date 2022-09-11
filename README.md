# pyWinKeys

Provides a simple API for performing mouse and keyboard operations in Python for Windows. This implementation using the
ctypes library does not require the application itself to be in focus when performing its operations and can therefore
be ran in the background. The API requires at least Python 3.5 due to usage of the typing hint library.

Also provides a small easy-to-use scripting capability, which allows a user to create sequential delay-based automation
scripts which can be executed directly from a python environment. These simple automation scripts are limited to the
capabilities of the pyWinKeys API, meaning that the following actions can be scripted:

- Mouse movement actions
- Mouse click and hold/release actions
- Keyboard hotkey execution
- Keyboard write actions

Future actions which will potentially be added at a later point are:

- Mouse scrolling actions (already supported in the pyWinKeys API)
- Keyboard key hold/release actions (already supported in the pyWinKeys API)
- Uppercase support in keyboard write actions (should be added to the pyWinKeys API)

A small test program utilizing the script capabilities can be found in test/script-reader-example which consists of the
following files:

- main.py
  -- Executes the 'test' program, which loads and executes a script defined in example-script.txt
- example-script.txt
  -- Defines two easy scripts which can be executed by invoking main.py