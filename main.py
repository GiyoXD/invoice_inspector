
import tkinter as tk
from ui.app_window import AppWindow

def main():
    root = tk.Tk()
    # Optional: Set icon, theme
    try:
        root.state("zoomed")
    except: pass
    
    app = AppWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
