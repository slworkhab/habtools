
import win32com.client
import os
from os import path
import inspect
from datetime import datetime
import re
import xml.etree.ElementTree as ET
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
            self.destination_folder = self.jsprms.prms['path']['dest']

        except Exception as e:
            self.log.errlg(f"Wasted ! : {e}")
            raise


    def convert_sql_to_db2(self, sql: str) -> str:

        converted = sql

        # Backticks et crochets MS SQL / MySQL
        converted = re.sub(r'`([^`]*)`', r'\1', converted)
        converted = re.sub(r'\[([^\]]*)\]', r'\1', converted)

        # Convertir TRUE/FALSE
        converted = re.sub(r'\bTRUE\b', '1', converted, flags=re.IGNORECASE)
        converted = re.sub(r'\bFALSE\b', '0', converted, flags=re.IGNORECASE)

        # Retirer le ; final
        converted = converted.rstrip(';')

        return converted

    def extract_sql_from_powerquery(self, query: str, pattern) -> str:
        

        match = re.search(pattern, query.formula, re.DOTALL)

        if not match:
            print("Impossible de trouver une requête SQL dans Odbc.Query().")
            return ""            

        sql_raw = match.group(1)

        # Nettoyage PowerQuery
        sql_clean = sql_raw \
            .replace('\\"', '"') \
            .replace('""', '"') \
            .replace('#(lf)', os.linesep) \
            .replace('#lf', os.linesep) \
            .replace('#(cr)', os.linesep) \
            .replace('#(tab)', ' ')

        # Retirer les espaces multiples
        sql_clean = re.sub(r'\s+', ' ', sql_clean).strip()

        return sql_clean

    
    def str_to_textfile(self, filename: str, str_to_write: str, encoding: str = "cp1252") -> None:
        try:
            with open(filename, "w", encoding=encoding) as text_file:
                text_file.write(str_to_write)
            print(f"Fichier '{filename}' écrit avec encodage {encoding}.")
        except UnicodeEncodeError:
            print(f"Erreur d'encodage : impossible d'écrire avec {encoding}.")
        except Exception as e:
            print(f"Erreur lors de l'écriture du fichier : {e}")


        
    def extract_powerquery_queries(self, excel, excel_file):
        try: 
            # Lancer Excel via COM
            fck = ''
            query_found = False
            try: 
                # Ouvrir le fichier Excel
                wb = excel.Workbooks.Open(excel_file, UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
                # Vérifier si des requêtes existent
                if hasattr(wb, "Queries"):            
                    fname = f"{self.jsprms.prms['path']['result']}{os.path.sep}{file_utils.get_filename_without_extension(os.path.basename(excel_file).replace(' ',''))}.sql"                     
                    
                    for query in wb.Queries:
                                
                        # query.RefreshOnOpen = False       
                        if "select " in query.Formula.lower():
                            with open(fname, "a", encoding="cp1252", errors="ignore") as f:
                                                        
                                # f.write(self.convert_powerquery_to_db2_sql(query.Formula))  # Code M complet
                                # f.write(query.Formula)  # Code M complet
                                pattern1 = r'Odbc\.Query\s*\(\s*"[^"]*"\s*,\s*"((?:[^"\\]|\\.)*)"\s*\)'
                                pattern2 = r'Odbc\.Query\s*\(\s*"[^"]*"\s*,\s*"([\s\S]*?)"\s*\)'
                                db2_query = self.extract_sql_from_powerquery(query, pattern1)                                 
                                fck+=f"{db2_query}{os.linesep}"
                                if (db2_query is None):
                                    input("NONE")
                                if (db2_query.strip()!=""):
                                   db2_query = self.convert_sql_to_db2(db2_query)        
                                   print(f"db2_query--{db2_query}--")
                                else:
                                    print("############################")
                                    print("SEARCHING WITH OTHER PATTERN")
                                    print("############################")
                                    db2_query = self.extract_sql_from_powerquery(query, pattern2) 
                                    db2_query = self.convert_sql_to_db2(db2_query)
                                    fck+=f"{db2_query}{os.linesep}"
                                if (db2_query.strip()!=""):
                                    f.write(f"{query.Name}{os.linesep}")   
                                    f.write(db2_query)                                
                                    f.write(f'{os.linesep} {"-"*80} {os.linesep}')                                                         
                                    query_found=True      
                            print(f"Extraction terminée. Requête sauvegardee dans {fname}.")
                        else:
                            print("Pas de SELECT dans la requête.")
                       # Fermer le classeur si ouvert
                self.close_tab(wb)
            except Exception as e:
                # Gérer l'erreur sans afficher de popup
                print("Erreur ouverture :", e)  
            else:
                print(f"--{fck}--")
                # print(".")
            # timer.cancel()
        except Exception as e:
            # Gérer l'erreur sans afficher de popup
            print("Une erreur est survenue :", e)
        return query_found                    
     
     
           
    def close_tab(self, wb):
        if wb:
            wb.Close(SaveChanges=False)
      
    def read_str_file_first_line(self, file_path):        
        with open(file_path , "r", encoding="cp1252") as file:
            return file.readline()



    def browse_xlsx_for_sql(self):
        # Walk through all subdirectories and files
        # input(self.destination_folder)
        # file_utils.clean_dir(self.jsprms.prms['path']['result'])
        excel_extensions = (".xls", ".xlsx")
        excel = win32com.client.Dispatch("Excel.Application")            
        excel.Visible = True
        excel.DisplayAlerts = False  # Pas de pop-up
        excel.Visible = True  # Ne pas afficher Excel
        flag_file_path = f"{self.jsprms.prms['path']['flag_sql']}"        
        if not os.path.exists(flag_file_path):
            self.str_to_textfile(flag_file_path, "nope")         
        
        all_xlsx_files = [
            os.path.join(root, file)
            for root, dirs, files in os.walk(self.destination_folder)
            for file in files
            if file.lower().endswith(excel_extensions)
        ]
        total_files = len(all_xlsx_files)        
        found_file = False
        processed_files = 1
        for root, dirs, files in os.walk(self.destination_folder):
            for file in files:

                print(f"Processing: {processed_files} / {total_files} - {file} ")
                if self.read_str_file_first_line(flag_file_path) == file :
                    found_file = True
                self.str_to_textfile(flag_file_path, file)                
                if not found_file :  
                    print(f"Found file : {file}")                  
                    xl_file = os.path.join(root, file)                       
                    q_found = xlsx_extractor.extract_powerquery_queries(excel, xl_file)
                    processed_files += 1
                    if not q_found:
                        print ("NO REQUEST FOUND")
                    #file_utils.str_to_textfile(flag_file_path, file)
                else:
                    print(f"Found file : {file}")                  
          # Quitter Excel si lancé
        if excel:
            excel.Quit()

    def main(self):
        # Exemple d'utilisation
        self.init_main()
        self.browse_xlsx_for_sql()
        


####
xlsx_extractor = XlsxExtrator()  
xlsx_extractor.main()