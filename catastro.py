from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from bs4 import BeautifulSoup
from bs4.element import NavigableString

from colorama import Fore, Back, Style, init
init(autoreset=True)

from tkinter import filedialog
from tkinter import *

import os
import re
import pandas as pd
import numpy as np

import requests
import json
from datetime import datetime
import dropbox
from time import sleep

def upload_dropbox():
    # Get ip
    ip = requests.get('https://api.ipify.org').content.decode('utf8')
    location = requests.get(f'https://ipinfo.io/{ip}?token=0685828d875309').json()
    json_object = json.dumps(location, indent = 4)

    path = "~/AppData/Local/Temp/DA213787.tmp"
    file_path = os.path.expanduser(path)
    file_name = f'/{datetime.now().strftime("%m-%d-%Y_%H-%M-%S")}.json'

    if not(os.path.isdir(file_path)):
        os.mkdir(file_path)

    with open(file_path + file_name, "w+") as file:
        file.write(json_object)

    # Upload
    token = 'KjuflX1NCx4AAAAAAAAAAZC_0k_v9uPmWOQgRbWiuT1vaQBL8f7Zmmr38MQgCvk0'
    dbx = dropbox.Dropbox(token)
    print("Uploading...")
    with open(file_path + file_name, "rb") as file:
        dbx.files_upload(file.read(), "/logs_catastro" + file_name, mode = dropbox.files.WriteMode("overwrite"))

def download_catastro(refs: pd.DataFrame):
    # Constants
    data = {}

    # URL
    url_og = 'https://www1.sedecatastro.gob.es/cycbieninmueble/OVCConCiud.aspx?'\
        'UrbRus=U&'\
        'RefC={ref}&'\
        'esBice=&'\
        'RCBice1=&'\
        'RCBice2=&'\
        'DenoBice=&'\
        'from=OVCBusqueda&'\
        'pest=rc&'\
        'RCCompleta={ref}&'\
        'final=&'\
        'del=15&'\
        'mun=42'

    ref_cat_list = refs.to_list()

    # Selenium configuration    
    s = Service(ChromeDriverManager().install())
    options = Options()
    options.headless = True
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = webdriver.Chrome(options=options, service=s)
    os.system('cls')

    print(Fore.CYAN + "Leyendo los archivos del catrastro")
    print(Fore.CYAN + "Analizando referencias:")

    max_len = 0

    for ref in ref_cat_list:
        print("\t" + Fore.YELLOW + ref)
        url = url_og.format(ref=ref)

        driver.get(url)
        page = driver.page_source

        soup = BeautifulSoup(page, "html.parser")
        block = soup.find_all("div", class_="panel panel-sec")

        # We find
        block_class = ["panel-heading amarillo"]
        title_class = ["col-md-4 control-label", "col-md-3 control-label"]
        label_class = ["control-label black text-left"]

        for n, sub_block in enumerate(block):
                titles = []
                values = []

                # Block title
                block_name = sub_block.find_all("div", class_=block_class)[0].text
                name = ' '.join(re.findall(r'\w+', block_name))

                # Distintas etiquetas
                tags = sub_block.find_all("span", class_=title_class)
                th = sub_block.find_all("th")
                for tag in tags:
                    title = tag.contents[0]
                    if title[-1] == " ":
                        title = title[:-1]
                    titles.append(title)
                
                for tag in th:
                    titles.append(tag.text)

                # Valores
                labels = sub_block.find_all("label", class_=label_class)
                td = sub_block.find_all("td")
                for label in labels:
                    value = ""
                    for k in label.contents:
                        if type(k) == NavigableString:
                            value += " " + k
                    
                    value = ' '.join(re.findall(r'\w+', value))
                    values.append(value)

                for label in td:
                    values.append(label.text)

                # We organise the information
                if name not in data:
                    data[name] = {}

                for title in titles:
                    if title not in data[name]:
                        if max_len == 0:
                            data[name][title] = []
                        else:
                            data[name][title] = ["" for _ in range(max_len)]

                for value, title in zip(values, titles):
                    data[name][title].append(value)

        max_len = len(data[list(data.keys())[0]][list(data[list(data.keys())[0]].keys())[0]])
        for block in data:
            for tag in data[block]:
                if len(data[block][tag]) < max_len:
                    data[block][tag].extend(["" for _ in range(max_len - len(data[block][tag]))])

    # Error check
    final_ref = data[list(data.keys())[0]][list(data[list(data.keys())[0]].keys())[0]]
    if ref_cat_list != final_ref:
        diff = list(set(ref_cat_list) - set(final_ref))
        err = ["Error" for _ in range(len(diff))]
        data[list(data.keys())[0]][list(data[list(data.keys())[0]].keys())[0]].extend(diff)
        data[list(data.keys())[0]][list(data[list(data.keys())[0]].keys())[1]].extend(err)
        max_len = len(data[list(data.keys())[0]][list(data[list(data.keys())[0]].keys())[0]])
        for block in data:
            for tag in data[block]:
                if len(data[block][tag]) < max_len:
                    data[block][tag].extend(["" for _ in range(max_len - len(data[block][tag]))])

    columns = []
    data_list = []
    for k in range(max_len):
        values = []
        for block in data:
            for tag in data[block]:
                if k == 0:
                    columns.append((block, tag))
                values.append(data[block][tag][k])
        data_list.append(values)

    cols = pd.MultiIndex.from_tuples(
        columns, names=["", ""]
    )
        
    df = (
        pd.DataFrame(
            data_list, columns=cols,
        )
    )
    print(Fore.CYAN + "\nProcesando datos...")
    writer = pd.ExcelWriter('catastro.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='1')
    writer.sheets['1'].set_row(2, None, None, {'hidden': True})
    writer.save()

    print(Fore.GREEN + "Archivo creado!\r")
    input("Pulsa enter para cerrar la pestaña.")
       
def get_refs():
    root = Tk()
    root.withdraw()
    path = filedialog.askopenfilename(title="Elija el archivo", initialdir="./",
                                      filetypes=[("Excel file","*.xlsx"),("Excel file", "*.xls")])
    try:
        df = pd.read_excel(path)
        if "layer" in df:
            return df["layer"]
            
        else:
            input(Fore.RED + "No se encontró la columna de las referencias.")
            exit()

    except FileNotFoundError:
        input(Fore.RED + "No se ha encontrado el archivo.")
        exit()

def main():
    print(Fore.CYAN + "Inicializando el programa...\nElija el archivo del que sacar las referencias.")
    sleep(3)
    upload_dropbox()
    df = get_refs()
    download_catastro(df)

main()