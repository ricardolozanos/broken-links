import requests
import openpyxl
#import csv
import tkinter as tk
import os
#import shutil
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime
from threading import Thread
from tkinter import messagebox
from tkinter.ttk import Progressbar, Treeview,  Combobox
from controller import Controller
from PIL import Image, ImageTk


######################################################################
######################################################################
######################################################################
######################################################################


class CustomButton(tk.Button):
    def __init__(self, master=None, cnf={}, **kw):
        self.page_progress_label = kw.pop("page_progress_label", None)
        self.link_progress_label = kw.pop("link_progress_label", None)
        super().__init__(master, cnf, **kw)
        self.configure(foreground='white', background='blue', font=('Helvetica', 12))

    def update_progress_labels(self, current_page, total_pages, current_link, total_links):
        if self.page_progress_label:
            self.page_progress_label.config(text=f"Page {current_page}/{total_pages}")
        if self.link_progress_label:
            self.link_progress_label.config(text=f"Link {current_link}/{total_links}")

######################################################################
######################################################################
######################################################################
######################################################################

class App:
    def __init__(self, root, controller=None):
        self.root = root
        self.root.title("Broken Link Checker")
        self.root.geometry("600x600")  # Set the minimum width and height
        self.controller=controller

        self.current_page = 0
        self.total_pages = 0

        self.current_link = 0
        self.total_links = 0

        # Load the image
        image_path = "icon/icon.png"
        image = Image.open(image_path)

        max_width = 100
        max_height = 100
        image.thumbnail((max_width, max_height))

        # Create a PhotoImage object from the image
        self.photo_image = ImageTk.PhotoImage(image)

        # Main Frame
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Select Folder Frame
        self.folder_frame = tk.Frame(self.main_frame)
        self.folder_frame.pack(fill=tk.BOTH, pady=10)

        self.folder_var = tk.StringVar()
        self.folder_options = self.controller.get_folders_in_same_level()
        if self.folder_options:
            self.folder_var.set(self.folder_options[0])

        # Image Label
        self.image_label = tk.Label(self.folder_frame, image=self.photo_image)
        self.image_label.pack(side=tk.LEFT)

        # Select Folder Label
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
        self.folder_dropdown.bind("<<ComboboxSelected>>", self.update_excel_list)

        # Progress Labels for Page and Link Progress Bars
        self.page_progress_label = tk.Label(self.excel_list_frame, text="", font=('Helvetica', 12))
        self.page_progress_label.pack()

        self.link_progress_label = tk.Label(self.excel_list_frame, text="", font=('Helvetica', 12))
        self.link_progress_label.pack()


        self.start_button = CustomButton(self.main_frame, text="Check for Broken Links",
                                         command=self.start_link_checking_thread,
                                         page_progress_label=self.page_progress_label,
                                         link_progress_label=self.link_progress_label)
        self.start_button.pack()
    

    # Get the list of Excel files in the selected folder
    def get_excel_files(self, folder_path):
        return [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

        # Function to update the Excel files list based on the selected folder
    def update_excel_list(self, event):
        selected_folder = self.folder_var.get()
        excel_files = self.get_excel_files(selected_folder)

        # Clear the listbox before updating
        self.excel_listbox.delete(*self.excel_listbox.get_children())

        # Populate the listbox
        for i, excel_file in enumerate(excel_files):
            self.excel_listbox.insert("", "end", values=(excel_file, "", ""), tags=("evenrow",) if i % 2 == 0 else ("oddrow",))
    
    def start_link_checking_thread(self):
        thread = Thread(target=self.check_links_thread)
        thread.start()

        self.current_page = 1
        
        self.update_page_progress(self.current_page, self.total_pages)


    def check_links_thread(self):
        folder_path = self.folder_var.get()
        
        # Check if any item is selected in the Treeview
        selected_items = self.excel_listbox.selection()
        if not selected_items:
            messagebox.showinfo("Error", "Please select an Excel file.")
            return

        # Retrieve the selected item and proceed with link checking
        selected_excel_with_extension = self.excel_listbox.item(selected_items[0], "values")[0]
        selected_excel_name, _ = os.path.splitext(selected_excel_with_extension)
        excel_file_path = os.path.join(self.folder_var.get(), selected_excel_with_extension)
        print(excel_file_path)


        broken_links_report = self.check_broken_links(excel_file_path)

        if broken_links_report:
            print("Broken Links Report:")
            total_links = len(broken_links_report)
            for i, (link, page_link, section, relative) in enumerate(broken_links_report):
                print(f"Broken Link on page '{link}' in '{section}' section: {page_link}, relative: {relative}")
                self.update_link_progress(i + 1, total_links)
            self.link_progress_label.config(text="Link checking completed.")
            
            # Save the report to the appropriate folder
            try:
                report_folder = self.save_report_to_folder(folder_path, selected_excel_name, broken_links_report)
                print(f"Broken links report saved to '{report_folder}'.")

                # Generate Word document from Excel data
                word_output_file = os.path.join(report_folder, "broken_links_report.docx")
                self.controller.create_word_document_from_excel(broken_links_report, word_output_file)
                print(f"Word document created at '{word_output_file}'.")
            except NotADirectoryError as e:
                print(e)
        else:
            print("No broken links found.")

        self.update_page_progress("", "")
        self.update_link_progress("", "")
    
    def check_broken_links(self, excel_file):
        try:
            print(f"Opening file: {excel_file}...")
            wb = openpyxl.load_workbook(excel_file)
            sheet = wb.active
            column_with_links = 'A'
            column_with_templates = 'B'

            broken_links_report = []
            total_links_checked = 0
            total_links = 0

            for cell in sheet[column_with_links]:
                if cell.value:
                    total_links += 1

            print(f"Checking {total_links} links inside {excel_file}...")
            self.total_pages = len([cell for cell in sheet[column_with_links] if cell.value])

            # Create a Progressbar widget for the overall progress
            print(self.total_pages)
            progress_bar = Progressbar(self.root, orient='horizontal', length=300, maximum=self.total_pages, mode='determinate')
            progress_bar.pack()
            
            for link_cell, template_cell in zip(sheet[column_with_links], sheet[column_with_templates]):
                if link_cell.value:
                    link = link_cell.value
                    template = template_cell.value

                    total_links_checked += 1
                    page_links = self.controller.get_links_from_page(link)
                    
                    # Calculate the total number of links for the current page
                    self.total_links = len(page_links)
                    print(self.total_links)
                    # Page pro  gress bar starts here
                    page_progress_var = tk.DoubleVar()
                    page_progress = Progressbar(self.root, variable=page_progress_var, length=300, maximum=self.total_links, mode='determinate')
                    page_progress.pack()

                    # Update page progress label
                    self.update_page_progress(total_links_checked, self.total_pages)

                    for i, page_link in enumerate(page_links):
                        absolute_link = urljoin(link, page_link)
                        final_url = self.controller.get_final_url(absolute_link)
                        section = self.controller.identify_section(link, page_link, template)
                        
                        try:
                            if final_url is not None:
                                try:
                                    response = requests.get(final_url)
                                    if response.status_code != 200:
                                        if (link, final_url, section, page_link) not in broken_links_report:  # Check for duplicates
                                            broken_links_report.append((link, final_url, section, page_link))
                                            print(f"Broken link found on page '{link}' in '{section}' section: {final_url}", end='\r')
                                        else:
                                            print(f"Duplicate broken link found on page '{link}': {final_url}", end='\r')
                                    else:
                                        pass  # No need for this message
                                except requests.exceptions.RequestException:
                                    if (link, final_url, section, page_link) not in broken_links_report:  # Check for duplicates
                                        broken_links_report.append((link, final_url, section, page_link))
                                        print(f"Connection error occurred for link on page '{link}' in '{section}' section: {final_url}", end='\r')
                                    else:
                                        print(f"Duplicate connection error occurred for link on page '{link}': {final_url}", end='\r')
                            else:
                                # Include invalid links in the broken links report
                                broken_links_report.append((link, page_link, section, "Invalid URL"))
                                print(f"Invalid URL: {absolute_link}", end='\r')
                            page_progress_var.set(i + 1)
                            self.update_link_progress(i + 1, self.total_links)
                            self.root.update_idletasks()
                        except TypeError as e:
                            print(f"Error: Unexpected value in cell ({link_cell.row}, {link_cell.column}). {e}")
                            print(f"Skipping the link and continuing...")
                            continue
                        self.current_page += 1
                        page_progress['value']=self.current_page
                    page_progress.destroy()

                    progress_bar['value'] = total_links_checked
                    self.root.update_idletasks()

            print("\nLink checking completed.")
            print(f"Total pages checked: {total_links_checked}")
            print(f"Total broken links found: {len(broken_links_report)}")
                
            # Move the return statement here to ensure it is outside the loop
            progress_bar.destroy()
            return broken_links_report

        except openpyxl.utils.exceptions.InvalidFileException:
            print(f"Error: Unable to open the file '{excel_file}'. Please check if it's a valid Excel file.")
            return []
        except Exception as e:
            print(f"Error: An unexpected error occurred while processing the file '{excel_file}': {e}")
            return []
    
    #NO SELFS
    def save_report_to_folder(self, folder_path, excel_name, report_data):
        # Create a folder with the name of the Excel file (if it doesn't exist)
        excel_folder = os.path.join(folder_path, excel_name)
        if not os.path.exists(excel_folder):
            os.makedirs(excel_folder)

        # Check if the excel_folder is a valid directory
        if not os.path.isdir(excel_folder):
            raise NotADirectoryError(f"'{excel_folder}' is not a valid directory.")

        # Get the number of existing reports in the folder
        report_count = len([f for f in os.listdir(excel_folder) if f.startswith("Report_")])

        # Create a folder for the current report with the format "Report_x_{time}"
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_folder_name = f"Report_{report_count + 1}_{now}"
        report_folder = os.path.join(excel_folder, report_folder_name)
        os.makedirs(report_folder)

        # Save the broken links report to a CSV file inside the report folder
        csv_file_path = os.path.join(report_folder, "broken_links_report.csv")
        header = ["Page", "Broken Link", "Section", "Relative Link"]
        self.controller.save_to_csv(csv_file_path, report_data, header)

        return report_folder
    
    def update_page_progress(self,current_page, total_pages):
        self.page_progress_label.config(text=f"Page {current_page}/{total_pages}")

    def update_link_progress(self,current_link, total_links):
        self.link_progress_label.config(text=f"Link {current_link}/{total_links}")
