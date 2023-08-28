import tkinter as tk
from controller import Controller
from app import App




def main():
    root = tk.Tk()
    controller = Controller()
    app = App(root, controller)
    root.mainloop()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()

