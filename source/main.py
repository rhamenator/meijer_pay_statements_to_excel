import sys
import traceback
from interface_library import create_gui, run
from logging_utils import ErrorLogger
from main_utils import command_line_mode, is_gui_mode

def main(*args):
    global root
    if len(args) > 0:
        input_file = args[1] if len(args) > 1 else None
        output_file = args[2] if len(args) > 2 else None
    if command_line_mode():
        # try:
            root = None
            run(input_file, output_file)
        # except Exception as e:
        #     print(f"An error occurred: {e}", file=sys.stderr)
        #     traceback.print_exc(file=sys.stderr)  # Print the traceback
        #     if root:
        #         ErrorLogger.showerror(f"An error occurred:\n{e}")
        #         if root is not None: root.destroy()
        #     raise
    elif is_gui_mode():
        create_gui(input_file, output_file)
    else:
        # try:
            root = None
            run(input_file, output_file)

if __name__ == "__main__":
#    try:
    global root
    root = None
    main(sys.argv)
    # Exit with code 0 for success
    # except Exception as e:
    #     print(f"An error occurred: {e}", file=sys.stderr)
    #     traceback.print_exc(file=sys.stderr)
    if root is not None: root.destroy()  # Close the GUI if it's open
    # sys.exit(1)  # Exit with code 1 for failure
    sys.exit(0)