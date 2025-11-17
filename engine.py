# -*-coding:utf-8 -*
import os
from os import path
import inspect
from datetime import datetime
import shutil
import re
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

        
        @_error_decorator(False)
        def convert_to_db2(self, sql):
            # Remplacements simples
            for old, new in conversion_rules.items():
                sql = sql.replace(old, new)            
            # Exemple : convertir LIMIT n en FETCH FIRST n ROWS ONLY
            sql = re.sub(r"LIMIT\s+(\d+)", r"FETCH FIRST \1 ROWS ONLY", sql)
            
            return sql

  
        ##############################################
        @_error_decorator()
        def browse_xlsx_for_sql(self):
            # Walk through all subdirectories and files
            for root, dirs, files in os.walk(self.destination_folder):
                for file in files:
                    print(f"Processing: {file}")
                    xl_file = os.path.join(root, file)
                    # self.get_sql_from_xlsx(xl_file)
                    xlsx_extractor = XlsxExtrator(self.root_app, self.trace, self.log, self.jsprms)                      
                    xlsx_extractor.main(xl_file)
                   
                    



        @_error_decorator()
        def copy_xlsx(self):                
            # Create the destination folder if it doesn't exist
            os.makedirs(self.destination_folder, exist_ok=True)        
            excel_extensions = (".xls", ".xlsx")
            # Walk through all subdirectories and files
            print (self.source_folder)
            for root, dirs, files in os.walk(self.source_folder):
                # print(files)
                for file in files:
                    if file.lower().endswith(excel_extensions):
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
                            print(f"✅ Copied: {file}")
            print("✅ All .xlsx files have been copied.")