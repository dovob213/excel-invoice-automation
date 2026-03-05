import sys
import os
import tkinter as tk

# Ensure the project root is in the path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.append(project_root)

from src.gui import App

def main():
    try:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
    except Exception as e:
        print(f"Error launching application: {e}")
        input("Press Enter to close...")

if __name__ == "__main__":
    main()
