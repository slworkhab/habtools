import os
import shutil
import inspect
from datetime import datetime
import utils.mylog as mylog
import utils.jsonprms as jsonprms
import utils.file_utils as file_utils

import re
from utils.mydecorators import _error_decorator
from string import Template

from engine import Engine


class Bot:

    def trace(self, stck):
        #  print (f"{stck.function} ({ stck.filename}-{stck.lineno})")
        self.log.lg(f"{stck[0].function} ({ stck[0].filename}-{stck[0].lineno})")

    def handler(self, signum, frame):
        raise Exception("Timeout")

    @_error_decorator()
    def remove_logs(self):
        self.trace(inspect.stack())
        keep_log_time = self.jsprms.prms['log_keep']['time']
        keep_log_unit = self.jsprms.prms['log_keep']['unit']
        self.log.lg(f"=>clean logs older than {keep_log_time} {keep_log_unit}")
        file_utils.remove_old_files(f"{self.root_app}{os.path.sep}log", keep_log_time, keep_log_unit)

    def init_main(self, jsonfile="default"):
        try:
            self.root_app = os.getcwd()
            self.log = mylog.Log(self.root_app)
            self.log.init(jsonfile)           
            self.trace(inspect.stack())
            jsonFn = f"{self.root_app}{os.path.sep}data{os.path.sep}conf{os.path.sep}{jsonfile}.json"
            self.jsprms = jsonprms.Prms(jsonFn)
            self.time_out = self.jsprms.prms['time_out']         
            self.log.lg("=HERE WE GO=")
            self.remove_logs()
        except Exception as e:
            self.log.errlg(f"Wasted ! : {e}")
            raise
   



    def main(self, command="", jsonfile="", param1="", param2=""):       
        self.init_main()
        t1 = datetime.now()
        try:
                self.trace(inspect.stack())               
                if command == "":
                        nb_args = len(sys.argv)
                        command = "test" if (nb_args == 1) else sys.argv[1]
                        # fichier json en param
                        jsonfile = "default" if (nb_args < 3) else sys.argv[2].lower()                                
                        param1 = "default" if (nb_args < 4) else sys.argv[3].lower()
                        param2 = "default" if (nb_args < 5) else sys.argv[4].lower()
                        #param3 = "default" if (nb_args < 6) else sys.argv[5].lower()      
                        print("params=", command, jsonfile, param1, param2)                                                                                    
                engine = Engine(self.root_app, self.trace, self.log, self.jsprms)                                                                        
                self.log.lg("=Here I am=")   
                if (command == "copyxl"):         
                        engine.copy_xlsx()
                        wk = input("waiting : ")
                if (command == "browse_xlsx_for_sql"):         
                        engine.browse_xlsx_for_sql()
                        wk = input("Done : ")                                                      
        except Exception as e:  
                print("GLOBAL MAIN EXCEPTION")
                self.log.errlg(e)
        finally:
            t2 = datetime.now()
            dt = t2 - t1
            self.log.lg("Done (elapse : %s)" % dt)                
            print("This is the end")            