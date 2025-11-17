# -*-coding:utf-8 -*-

import os
from os import path
import inspect
from datetime import datetime
import shutil
import re
import pandas as pd
import zipfile
import sys
import xml.etree.ElementTree as ET
import base64
import io
import html

class XlsxExtrator:

    def __init__(self, root_app, trace, log, jsprms):
        self.root_app = os.getcwd()
        self.trace = trace
        self.log = log
        self.jsprms = jsprms
        self.source_folder = self.jsprms.prms['path']['source']
        self.destination_folder = self.jsprms.prms['path']['dest']

    # ----------------- METHODS -----------------

    def extract_powerquery_from_xlsx(self, file_path):
        queries = []

        # Vérifier si le fichier est bien un ZIP
        if not zipfile.is_zipfile(file_path):
            raise ValueError(f"Le fichier {file_path} n'est pas un fichier XLSX valide (ZIP).")

        with zipfile.ZipFile(file_path, 'r') as z:
            for name in z.namelist():
                # Parcourir tous les fichiers XML internes
                if name.endswith('.xml') or 'connections' in name or 'customData' in name:
                    with z.open(name) as f:
                        content = f.read().decode('utf-8', errors='ignore')

                        # 1️⃣ Chercher les balises <m:Formula>
                        matches = re.findall(r'<m:Formula>(.*?)</m:Formula>', content, re.DOTALL)

                        # 2️⃣ Si rien trouvé, chercher du code M brut (commence par "let")
                        if not matches and "let" in content:
                            # Extraire un bloc commençant par let jusqu'à in
                            matches = re.findall(r'(let.*?in[^\n<]*)', content, re.DOTALL)

                        for match in matches:
                            decoded_query = html.unescape(match)
                            queries.append(decoded_query)

        if not queries:
            print("⚠ Aucune requête Power Query trouvée dans ce fichier.")



    def main(self, file_path):
        # Exemple d'utilisation        
        powerquery_code = self.extract_powerquery_from_xlsx(file_path)
        # for i, query in enumerate(powerquery_code, 1):
        #    print(f"Requête {i}:\n{query}\n{'-'*40}")