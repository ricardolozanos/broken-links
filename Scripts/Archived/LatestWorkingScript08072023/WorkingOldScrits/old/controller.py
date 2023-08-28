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

class Controller:
    def __init__(self):
        self.created = True
        
    def get_final_url(self,url):
        try:
            response = requests.get(url, allow_redirects=True)
            return response.url
        except requests.exceptions.RequestException:
            return None

        
    def clean_link(self, link):
        if link and not link.startswith("mailto:") and not link.startswith("#"):
            link = unquote(link).strip()
            # CHANGED TO WHILE LOOP
            while link.endswith("%20"):
                link = link[:-1]
            return link
        return None

    # Get all the links from page
    def get_links_from_page(self, url):
        try:
            response = requests.get(url)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                links = [self.clean_link(link.get('href')) for link in soup.find_all('a', href=True)]
                return [link for link in links if link is not None]  # Filter out None values
            else:
                return []
        except requests.exceptions.RequestException:
            return []
