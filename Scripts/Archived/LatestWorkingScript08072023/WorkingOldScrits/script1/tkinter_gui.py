# tkinter_gui.py

import tkinter as tk
from tkinter.ttk import Combobox, Treeview
import requests
import openpyxl
import csv
import tkinter as tk
import os
from bs4 import BeautifulSoup
from urllib.parse import urljoin, unquote
from tkinter import messagebox, filedialog
from tkinter.ttk import Progressbar
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor, Inches, Pt
from docx import Document as WordDocument
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from threading import Thread
from controller import check_broken_links, save_report_to_folder, create_word_document_from_excel




class CustomButton(tk.Button):
    def __init__(self, master=None, cnf={}, page_progress_label=None, link_progress_label=None, **kw):
        super().__init__(master, cnf, **kw)
        self.page_progress_label = page_progress_label
        self.link_progress_label = link_progress_label
        self.configure(foreground='white', background='blue', font=('Helvetica', 12))

    def update_progress_labels(self, current_page, total_pages, current_link, total_links):
        if self.page_progress_label:
            self.page_progress_label.config(text=f"Page {current_page}/{total_pages}")
        if self.link_progress_label:
            self.link_progress_label.config(text=f"Link {current_link}/{total_links}")

class App:
    def __init__(self, root, update_excel_list, folder_options):
        self.root = root
        self.root.title("Broken Link Checker")
        self.root.geometry("600x600")  # Set the minimum width and height

        # Main Frame
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Select Folder Frame
        self.folder_frame = tk.Frame(self.main_frame)
        self.folder_frame.pack(fill=tk.BOTH, pady=10)

        self.folder_var = tk.StringVar()
        self.folder_options = folder_options
        if self.folder_options:
            self.folder_var.set(self.folder_options[0])

        self.folder_label = tk.Label(self.folder_frame, text="Select Folder:", font=('Helvetica', 14))
        self.folder_label.pack(side=tk.LEFT)

        self.folder_dropdown = Combobox(self.folder_frame, textvariable=self.folder_var, values=self.folder_options)
        self.folder_dropdown.pack(side=tk.LEFT)

        # Excel List Frame
        self.excel_list_frame = tk.Frame(self.main_frame)
        self.excel_list_frame.pack(fill=tk.BOTH, expand=True, padx=10)

        self.excel_list_label = tk.Label(self.excel_list_frame, text="Excel Files in Selected Folder:", font=('Helvetica', 14))
        self.excel_list_label.pack()

        # Create the Treeview widget
        self.excel_listbox = Treeview(self.excel_list_frame, selectmode=tk.BROWSE, columns=("Link", "Page Link", "Section"),
                                      show="headings", height=10)
        self.excel_listbox.pack(fill=tk.BOTH, expand=True)

        # Configure tags for even and odd rows
        self.excel_listbox.tag_configure("evenrow", background="white")
        self.excel_listbox.tag_configure("oddrow", background="lightgray")

        # Bind the update_excel_list function to the Combobox selection event
        self.folder_dropdown.bind("<<ComboboxSelected>>", update_excel_list)

        # Progress Labels for Page and Link Progress Bars
        self.page_progress_label = tk.Label(self.excel_list_frame, text="", font=('Helvetica', 12))
        self.page_progress_label.pack()

        self.link_progress_label = tk.Label(self.excel_list_frame, text="", font=('Helvetica', 12))
        self.link_progress_label.pack()


        self.start_button = CustomButton(self.main_frame, text="Check for Broken Links",
                                         command=self.check_links_thread,
                                         page_progress_label=self.page_progress_label,
                                         link_progress_label=self.link_progress_label)
        self.start_button.pack()

    
    def get_excel_files(self, folder_path):
        return [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

    # Method to update the page progress label
    def update_page_progress(self, current_page, total_pages):
        self.page_progress_label.config(text=f"Page {current_page}/{total_pages}")

    # Method to update the link progress label
    def update_link_progress(self, current_link, total_links):
        self.link_progress_label.config(text=f"Link {current_link}/{total_links}")

    # Method to check broken links
    def check_links_thread(self):
        folder_path = self.folder_var.get()
        selected_items = self.excel_listbox.selection()

        if not selected_items:
            messagebox.showinfo("Error", "Please select an Excel file.")
            return

        selected_excel_with_extension = self.excel_listbox.item(selected_items[0], "values")[0]
        selected_excel_name, _ = os.path.splitext(selected_excel_with_extension)
        excel_file_path = os.path.join(folder_path, selected_excel_with_extension)
        print(excel_file_path)
        broken_links_report = check_broken_links(excel_file_path)

        if broken_links_report:
            print("Broken Links Report:")
            total_links = len(broken_links_report)
            for i, (link, page_link, section, relative) in enumerate(broken_links_report):
                print(f"Broken Link on page '{link}' in '{section}' section: {page_link}, relative: {relative}")
                self.update_link_progress(i + 1, total_links)
            self.link_progress_label.config(text="Link checking completed.")

            # Save the report to the appropriate folder
            try:
                report_folder = save_report_to_folder(folder_path, selected_excel_name, broken_links_report)
                print(f"Broken links report saved to '{report_folder}'.")

                # Generate Word document from Excel data
                word_output_file = os.path.join(report_folder, "broken_links_report.docx")
                create_word_document_from_excel(broken_links_report, word_output_file)
                print(f"Word document created at '{word_output_file}'.")
            except NotADirectoryError as e:
                print(e)
        else:
            print("No broken links found.")

        self.update_page_progress("", "")
        self.update_link_progress("", "")

        return broken_links_report    
