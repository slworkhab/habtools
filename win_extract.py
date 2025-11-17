
import win32com.client
import os
from os import path
import inspect
from datetime import datetime
import xml.etree.ElementTree as ET
from xlsx_extractor import XlsxExtrator
#mes libs
import utils.str_utils as str_utils
import utils.file_utils as file_utils
from utils.urls import Urls
from utils.mydecorators import _error_decorator
import utils.jsonprms as jsonprms
import utils.mylog as mylog

class XlsxExtrator:

    def __init__(self):
        self.root_app = os.getcwd()
      

    @_error_decorator()
    def remove_logs(self):
        # self.trace(inspect.stack())
        keep_log_time = self.jsprms.prms['log_keep']['time']
        keep_log_unit = self.jsprms.prms['log_keep']['unit']
        self.log.lg(f"=>clean logs older than {keep_log_time} {keep_log_unit}")
        file_utils.remove_old_files(f"{self.root_app}{os.path.sep}log", keep_log_time, keep_log_unit)

    def init_main(self, jsonfile="default"):
        try:
            self.root_app = os.getcwd()
            self.log = mylog.Log(self.root_app)
            self.log.init(jsonfile)           
            # self.trace(inspect.stack())
            jsonFn = f"{self.root_app}{os.path.sep}data{os.path.sep}conf{os.path.sep}{jsonfile}.json"
            self.jsprms = jsonprms.Prms(jsonFn)
            self.time_out = self.jsprms.prms['time_out']         
            self.log.lg("=HERE WE GO=")
            self.remove_logs()
            self.source_folder = self.jsprms.prms['path']['source']
            self.destination_folder = self.jsprms.prms['path']['dest']

        except Exception as e:
            self.log.errlg(f"Wasted ! : {e}")
            raise


    def extract_powerquery_queries(self, excel_file):
        # Lancer Excel via COM
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Ne pas afficher Excel

        # Ouvrir le fichier Excel
        wb = excel.Workbooks.Open(excel_file)

        # Vérifier si des requêtes existent
        if hasattr(wb, "Queries"):
            fname = f"{self.jsprms.prms['path']['result']}\\{file_utils.get_filename_without_extension(os.path.basename(excel_file).replace(' ',''))}.sql"
            with open(fname, "w", encoding="utf-8") as f:
                for query in wb.Queries:
                    if "select" in query.Formula.lower():
                        f.write(f"Nom de la requête : {query.Name}\n")
                        f.write("Code M :\n")
                        f.write(query.Formula)  # Code M complet
                        f.write("\n" + "-"*80 + "\n")
            print(f"Extraction terminée. Les requêtes ont été sauvegardées dans {fname}.")
        else:
            print("Aucune requête Power Query trouvée dans ce fichier.")

        # Fermer le fichier et Excel
        wb.Close(False)
        excel.Quit()


    def browse_xlsx_for_sql(self):
        # Walk through all subdirectories and files
        # input(self.destination_folder)
        for root, dirs, files in os.walk(self.destination_folder):       
            for file in files:
                print(f"Processing: {file}")
                xl_file = os.path.join(root, file)                
                xlsx_extractor.extract_powerquery_queries(xl_file)

    def main(self):
        # Exemple d'utilisation
        # powerquery_code = xlsx_extractor.extract_powerquery_from_xlsx(file_path) 
        self.init_main()
        self.browse_xlsx_for_sql()
        


####
xlsx_extractor = XlsxExtrator()  
xlsx_extractor.main()