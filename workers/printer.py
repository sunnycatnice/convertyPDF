import os
import sys

def print_file(file_path):
    if sys.platform == "darwin":
        os.system(f"lpr {file_path}")
    elif sys.platform == "win32":
        os.startfile(file_path, "print")
    else:
        print("Printing not supported on this OS")
