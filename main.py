import tkinter as tk
from gui import BartenderGUI

def main():
    root = tk.Tk()
    app = BartenderGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
