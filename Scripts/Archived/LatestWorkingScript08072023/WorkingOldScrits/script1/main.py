# main.py

import tkinter as tk
from tkinter.ttk import Combobox, Treeview
from tkinter_gui import App
from controller import  update_excel_list, get_folders_in_same_level


if __name__ == "__main__":
    root = tk.Tk()  # Create the root window
    folder_options = get_folders_in_same_level()
    app = App(root, update_excel_list, folder_options)
    app.folder_dropdown.bind("<<ComboboxSelected>>", lambda event: update_excel_list(app))  # Bind the event with app instance
    root.mainloop()