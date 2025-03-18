#!/usr/bin/env python3

import os
import sys
import tkinter as tk
from duplex_printer import DuplexPrinterApp

def main():
    if sys.platform == 'darwin':  # macOS specific configuration
        os.environ['NSRequiresAquaSystemAppearance'] = 'NO'
    root = tk.Tk()
    app = DuplexPrinterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()