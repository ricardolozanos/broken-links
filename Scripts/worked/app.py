import requests
import openpyxl
import tkinter as tk
import os
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime
from threading import Thread
from tkinter import messagebox
from tkinter.ttk import Progressbar, Treeview,  Combobox
from controller import Controller
from PIL import Image, ImageTk
import concurrent.futures
import threading
import time  
import queue  
from concurrent.futures import ThreadPoolExecutor, TimeoutError
import requests
from requests.adapters import HTTPAdapter, Retry
import urllib3
import threading
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from urllib3.exceptions import TimeoutError





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

        #Save Links to prevent multiple requests
        self.workingLinks = []
        self.brokenLinks = []

        # Load the image
        image_path = "icon/icon.png"
        image = Image.open(image_path)

        self.reports_folder = "Reports"
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
        self.folder_options = self.controller.get_files_in_excel_folder()
        if self.folder_options:
            self.folder_var.set(self.folder_options[0])


        # Image Label
        self.image_label = tk.Label(self.folder_frame, image=self.photo_image)
        self.image_label.pack(side=tk.LEFT)


        
        
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

        self.update_excel_list()
        
        # Progress Labels for Page and Link Progress Bars
        self.page_progress_label = tk.Label(self.excel_list_frame, text="", font=('Helvetica', 12))
        self.page_progress_label.pack()

        self.page_progress_var = tk.StringVar()
        self.page_progress = tk.Label(self.root, textvariable=self.page_progress_var)
        self.page_progress.pack()

        self.files_name=' '

        # Start a worker thread
        self.worker_thread = threading.Thread(target=self.worker_function)
        self.worker_thread.start()

        self.queue = queue.Queue()

        self.link_progress_label = tk.Label(self.excel_list_frame, text="", font=('Helvetica', 12))
        self.link_progress_label.pack()


        self.start_button = CustomButton(self.main_frame, text="Check for Broken Links",
                                         command=self.start_link_checking_thread,
                                         page_progress_label=self.page_progress_label,
                                         link_progress_label=self.link_progress_label)
        self.start_button.pack()

        # Start a periodic task to update the progress bar
        self.root.after(100, self.update_progress)
    
    def worker_function(self):
        for i in range(1, 11):
            # Simulate some work
            time.sleep(1)
            self.queue.put((i, 10))  # Put progress update in the queue

    def update_progress(self):
        try:
            while True:
                current_value, total = self.queue.get_nowait()  # Get updates from the queue
                self.page_progress_var.set(f"Page {current_value}/{total}")
        except queue.Empty:
            pass
        self.root.after(100, self.update_progress)

    # Get the list of Excel files in the selected folder
    def get_excel_files(self, folder_path):
        return [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and not f.startswith('.')]

        # Function to update the Excel files list based on the selected folder
    
    def update_excel_list(self, event=None):
        excel_files = self.controller.get_files_in_excel_folder()

        # Clear the listbox before updating
        self.excel_listbox.delete(*self.excel_listbox.get_children())

        # Populate the listbox
        for i, excel_file in enumerate(excel_files):
            self.excel_listbox.insert("", "end", values=(excel_file, "", ""), tags=("evenrow",) if i % 2 == 0 else ("oddrow",))




    def start_link_checking_thread(self):
        self.folder_var.set("Documents/Excel_Files")
        selected_items = self.excel_listbox.selection()
        if not selected_items:
            messagebox.showinfo("Error", "Please select an Excel file.")
            return

        selected_excel_with_extension = self.excel_listbox.item(selected_items[0], "values")[0]
        selected_excel_name, _ = os.path.splitext(selected_excel_with_extension)

        excel_folder = os.path.join("Documents", "Excel_Files")
        excel_file_path = os.path.join(excel_folder, selected_excel_with_extension)

        thread = Thread(target=self.check_links_thread)
        thread.start()

        self.current_page = 1
        self.update_page_progress(self.current_page, self.total_pages)



    def check_links_thread(self):
        
        # Check if any item is selected in the Treeview
        selected_items = self.excel_listbox.selection()
        if not selected_items:
            messagebox.showinfo("Error", "Please select an Excel file.")
            return

        # Retrieve the selected item and proceed with link checking
        selected_excel_with_extension = self.excel_listbox.item(selected_items[0], "values")[0]
        selected_excel_name, _ = os.path.splitext(selected_excel_with_extension)
        self.files_name= selected_excel_name
        excel_file_path = os.path.join(self.folder_var.get(), selected_excel_with_extension)
        


        broken_links_report = self.check_broken_links(excel_file_path)

        if broken_links_report:
            print("Broken Links Report:")
            total_links = len(broken_links_report)
            for i, (link, page_link, section, relative, linkname) in enumerate(broken_links_report):
                print(f"Broken Link on page '{link}' in '{section}' section: {page_link}, relative: {relative}")
                self.update_link_progress(i + 1, total_links)
            self.link_progress_label.config(text="Link checking completed.")
            
            # Save the report to the appropriate folder
            try:
                print(selected_excel_name)
                report_folder = self.save_report_to_folder(self.folder_var.get(), selected_excel_name, broken_links_report)
                print(f"Broken links report saved to '{report_folder}'.")

                # Generate Word document from Excel data
                current_datetime = datetime.now()

                # Combine year, month, and day as a single integer
                today = int(f"{current_datetime.year:04}{current_datetime.month:02}{current_datetime.day:02}")

                word_output_file = os.path.join(report_folder, f"{self.files_name}_broken_links_report_{today}.docx")
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

            threads = []
            
            for link_cell, template_cell in zip(sheet[column_with_links], sheet[column_with_templates]):
                if link_cell.value:
                    
                    link = link_cell.value
                    template = template_cell.value
                    print(f'Checking page: {link}')
                    total_links_checked += 1
                    page_links = self.controller.get_links_from_page_concurrently(link)
                    
                    # Calculate the total number of links for the current page
                    self.total_links = len(page_links)
                    # Page pro  gress bar starts here
                    page_progress_var = tk.DoubleVar()
                    page_progress = Progressbar(self.root, variable=page_progress_var, length=300, maximum=self.total_links, mode='determinate')
                    page_progress.pack()

                    # Update page progress label
                    self.update_page_progress(total_links_checked, self.total_pages)

                    thread_pool=[]
                    
                    for i, page_link in enumerate(page_links):
                        thread = threading.Thread(target=self.check_link, args=(page_progress_var, i, link, page_link, template, broken_links_report, link_cell, page_progress))
                        thread.start()
                        thread_pool.append(thread)

                        if len(thread_pool) >= 30:
                            for thread in thread_pool:
                                thread.join()
                            thread_pool = []

                    threads.extend(thread_pool)
                    self.current_page += 1
                    page_progress['value']=self.current_page
                page_progress.destroy()

                progress_bar['value'] = total_links_checked
                self.root.update_idletasks()

            
                # Wait for all threads to finish
            for thread in threads:
                thread.join()

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
    
    def check_link(self, page_progress_var, i, link, page_link, template, broken_links_report, link_cell, page_progress):
        
        absolute_link = urljoin(link, page_link)
        final_url = self.controller.get_final_url(absolute_link)
        section, link_name = self.controller.identify_section(link, page_link, template)
        broken_link_info = (link, final_url, section, page_link, link_name)  # Define broken_link_info once
        if section == 'Footer':
            return

        if final_url in self.brokenLinks:
            print("Broken Link already found before...")
            if broken_link_info not in broken_links_report:  # Check for duplicates
                broken_links_report.append(broken_link_info)
                print(f"Broken link found on page '{link}' in '{section}' section: {final_url}", end='\r')
            else:
                # broken_links_report.append(broken_link_info)
                print(f"Duplicate broken link found on page '{link}': {final_url}", end='\r')
            return

        if final_url in self.workingLinks:
            return
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
            }
            http = urllib3.PoolManager()
            
            response = http.request('GET', final_url, headers=headers, timeout=3.0)
            
            if response.status != 200:
                self.brokenLinks.append(final_url)
                if broken_link_info not in broken_links_report:  # Check for duplicates
                    broken_links_report.append(broken_link_info)
                    print(f"Broken link found on page '{link}' in '{section}' section: {final_url}", end='\r')
                else:
                    # broken_links_report.append(broken_link_info)
                    print(f"Duplicate broken link found on page '{link}': {final_url}", end='\r')
            else:
                self.workingLinks.append(final_url)
                pass  # No need for this message

            page_progress_var.set(i + 1)
            self.update_link_progress(i + 1, self.total_links)
            self.root.update_idletasks()

        except urllib3.exceptions.TimeoutError:
            self.brokenLinks.append(final_url)

            if broken_link_info not in broken_links_report:
                broken_links_report.append(broken_link_info)
            else:
                print("Error...")
                # broken_links_report.append(broken_link_info)
            print(f"Error occurred for link Timeout: {link}")
        except urllib3.exceptions.SSLError as ssl_error:
            # Attempt the request again with verify=False in case of SSL verification error
            print(f"SSL verification error occurred for link: {link}, Error: {ssl_error}")
            try:
                http = urllib3.PoolManager()
                print(f"{final_url}")
                response = http.request('GET', final_url, headers=headers, timeout=3.0, verify=False)
                
                if response.status != 200:
                    self.brokenLinks.append(final_url)

                    if broken_link_info not in broken_links_report:  # Check for duplicates
                        broken_links_report.append(broken_link_info)
                        print(f"Broken link found on page '{link}' in '{section}' section: {final_url}", end='\r')
                    else:
                        # broken_links_report.append(broken_link_info)
                        print(f"Duplicate broken link found on page '{link}': {final_url}", end='\r')
                else:
                    self.workingLinks.append(final_url)
                    pass
            except urllib3.exceptions.TimeoutError:
                self.brokenLinks.append(final_url)

                if broken_link_info not in broken_links_report:
                    broken_links_report.append(broken_link_info)
                else:
                    print("Error...")
                    # broken_links_report.append(broken_link_info)
                print(f"Error occurred for link Timeout: {link}")
        except Exception as e:
            self.brokenLinks.append(final_url)

            if broken_link_info not in broken_links_report:
                broken_links_report.append(broken_link_info)
            else:
                print("Error...")
                # broken_links_report.append(broken_link_info)
            print(f"Error occurred for link: {link}, Error: {e}")

        except TypeError as e:
            print(f"Error: Unexpected value in cell ({link_cell.row}, {link_cell.column}). {e}")
            print(f"Skipping the link and continuing...")
        
        finally:
            # Regardless of whether an exception was caught or not,
            # update the progress and UI
            page_progress_var.set(i + 1)
            self.update_link_progress(i + 1, self.total_links)
            self.root.update_idletasks()

        



    #NO SELFS
    def save_report_to_folder(self, excel_folder_name, excel_name, report_data):
        # Create a folder for reports if it doesn't exist
        if not os.path.exists(self.reports_folder):
            os.makedirs(self.reports_folder)

        # Create a folder with the name of the Excel file inside the reports folder (if it doesn't exist)
        excel_report_folder = os.path.join(self.reports_folder, excel_name)
        if not os.path.exists(excel_report_folder):
            os.makedirs(excel_report_folder)

        # Get the number of existing reports in the folder
        report_count = len([f for f in os.listdir(excel_report_folder) if f.startswith("Report_")])

        # Create a folder for the current report with the format "Report_x_{time}"
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_folder_name = f"Report_{report_count + 1}_{now}"
        report_folder = os.path.join(excel_report_folder, report_folder_name)
        os.makedirs(report_folder)

        # Generate Word document from Excel data
        current_datetime = datetime.now()

        # Combine year, month, and day as a single integer
        today = int(f"{current_datetime.year:04}{current_datetime.month:02}{current_datetime.day:02}")
        # Save the broken links report to a CSV file inside the report folder
        csv_file_path = os.path.join(report_folder, f"{self.files_name}_broken_links_report_{today}.csv")
        header = ["Page", "Broken Link", "Section", "Relative Link", "Link Name"]
        self.controller.save_to_csv(csv_file_path, report_data, header)

        return report_folder



    
    def update_page_progress(self, current_page, total_pages):
        self.current_page = current_page
        self.total_pages = total_pages

        # Schedule the update of the progress bar on the main thread
        self.root.after(0, self.update_page_progress_gui)

    def update_page_progress_gui(self):
        # Update the progress bar on the main thread
        self.page_progress['value'] = self.current_page
        self.page_progress_label.config(text=f"Page {self.current_page} of {self.total_pages}")

    def update_link_progress(self, current_link, total_links):
        self.current_link = current_link
        self.total_links = total_links

        # Schedule the update of the link progress label on the main thread
        self.root.after(0, self.update_link_progress_gui)

    def update_link_progress_gui(self):
        # Update the link progress label on the main thread
        self.link_progress_label.config(text=f"Link {self.current_link}/{self.total_links}")