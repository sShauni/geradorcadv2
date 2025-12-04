import tkinter as tk
import sys
import os
from interface import AppGUI

def resource_path(relative_path):
    """Permite acessar arquivos incluídos no executável PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return relative_path

if __name__ == "__main__":
    root = tk.Tk()

    # Define o ícone da janela
    root.iconbitmap(resource_path("icon.ico"))

    app = AppGUI(root)
    root.mainloop()
