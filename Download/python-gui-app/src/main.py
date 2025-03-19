import tkinter as tk
from tkinter import filedialog, messagebox
from gui.app_gui import AppGUI

def main():
    root = tk.Tk()
    root.title("Download Manager")
    app = AppGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()