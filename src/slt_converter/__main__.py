import sys
from slt_converter.converter import main
from slt_converter.path_entry import add_to_path

if __name__ == "__main__":
    if sys.platform == "win32":
        add_to_path()  # optional PATH modification on Windows
    main()
