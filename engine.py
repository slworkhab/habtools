# -*-coding:utf-8 -*
import os
from os import path
import inspect
from datetime import datetime
import shutil

import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from xlsx_extractor import XlsxExtrator
#mes libs
import utils.str_utils as str_utils
import utils.file_utils as file_utils
from utils.urls import Urls
from utils.mydecorators import _error_decorator

class Engine:

        def __init__(self, root_app, trace, log, jsprms):
                self.root_app = root_app
                self.trace = trace
                self.log = log
                self.jsprms = jsprms                
                self.root_app = os.getcwd()         
                self.source_folder = self.jsprms.prms['path']['source']
                self.destination_folder = self.jsprms.prms['path']['dest']                          


        def read_str_file_first_line(self, file_path):        
            with open(file_path , "r", encoding="cp1252", errors="ignore") as file:
                return file.readline()
        
        def str_to_textfile (self, filename, str_to_write):
            text_file = open(filename, "w", encoding="cp1252")
            text_file.write(str_to_write)
            text_file.close()

        @_error_decorator()
        def copy_xlsx(self):                
            # Create the destination folder if it doesn't exist
            os.makedirs(self.destination_folder, exist_ok=True)        
            excel_extensions = (".xls", ".xlsx")
            # Walk through all subdirectories and files
            print (self.source_folder)
            flag_file_path = f"{self.jsprms.prms['path']['flag_copy']}"
            if not os.path.exists(flag_file_path):
                self.str_to_textfile(flag_file_path, "nope")
            all_xlsx_files = [
            os.path.join(root, file)
            for root, dirs, files in os.walk(self.destination_folder)
            for file in files if file.lower().endswith(excel_extensions)]
            total_files = len(all_xlsx_files)  
            processed_files = 1
            found_file = False
            for root, dirs, files in os.walk(self.source_folder):
                for file in files:                    
                    if file.lower().endswith(excel_extensions):
                        try:
                            if self.read_str_file_first_line(flag_file_path) == file :
                                found_file = True
                            if not found_file :                             
                                source_path = os.path.join(root, file)
                                destination_path = os.path.join(self.destination_folder, file)
                                # If a file with the same name already exists, rename it
                                if os.path.exists(destination_path):
                                    base, ext = os.path.splitext(file)
                                    counter = 1
                                    while os.path.exists(destination_path):
                                        new_name = f"{base}_{counter}{ext}"
                                        destination_path = os.path.join(self.destination_folder, new_name)
                                        counter += 1
                                if not os.path.exists(destination_path):
                                    shutil.copy2(source_path, destination_path)
                                    self.str_to_textfile(flag_file_path, file)
                                    processed_files += 1
                                    print(f"✅ Copied: {file}")
                            else:
                                print(f"Found file : {file}")
                        except Exception as e:
                                print(f"Computer malfunction : {file} {e}")  
            print("✅ All .xlsx files have been copied.")