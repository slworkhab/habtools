
import win32com.client
import os
from os import path
from datetime import datetime
from pathlib import Path

#mes libs
import utils.str_utils as str_utils
import utils.file_utils as file_utils
from utils.urls import Urls
from utils.mydecorators import _error_decorator
import utils.jsonprms as jsonprms
import utils.mylog as mylog


class StreetCleaner:

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
            self.result_folder = self.jsprms.prms['path']['result']

        except Exception as e:
            self.log.errlg(f"Wasted ! : {e}")
            raise

    def clean(self):
        # Nettoyage des dossiers
        for folder in [self.destination_folder, self.result_folder]:
            file_utils.clean_dir(folder)

        # Suppression des fichiers de flag
        for key in ["flag_sql", "flag_copy"]:
            flag_path = Path(self.jsprms.prms["path"][key])
            if flag_path.exists():
                flag_path.unlink()

    def main(self):
        # Exemple d'utilisation
        self.init_main()
        self.clean()
        


####
street_cleaner = StreetCleaner()  
street_cleaner.main()