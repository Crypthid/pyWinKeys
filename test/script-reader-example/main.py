"""
 Created by Daniel Månsson at 2020-03-02
 File: main.py

 Simple execution test of script capabilities.
"""
from pyWinKeys.scripts.script_executor import ScriptExecutor


def main():
    print("Loading file example-script.txt...")
    se = ScriptExecutor("example-script.txt")
    # Run the "open_firefox" script
    print("Executing script 'open_firefox'...")
    se.execute("open_firefox")
    # Run the "test_selection" script
    # print("Exectuin script 'test_selection'...")
    # se.execute("test_selection")


if __name__ == '__main__':
    main()
    pass
