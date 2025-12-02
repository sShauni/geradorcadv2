# main.py
import tkinter as tk
from interface import AppGUI

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()