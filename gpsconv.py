import os
import numpy as np
import pandas as pd
import math
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from timer import Timer

# Initiate the browser
browser = webdriver.Chrome(ChromeDriverManager().install())

t = Timer()

pixlsx = "excel/input/"
poxlsx = "excel/output/"
ixlsx = os.listdir(pixlsx)

for x in ixlsx:
    xcoor = pd.read_excel(pixlsx+x, 0, usecols="A,B,J:L").to_numpy()
    datajlist = []

    for i in xcoor:
        t.start()
        xzon = str(i[2]).strip().replace(" ", "")
        if xzon and not math.isnan(i[3]) and not math.isnan(i[4]):
            xid = i[0]
            xname = i[1]
            xest = str(round(i[3]))
            xnord = str(round(i[4]))

            # sostituire con i valori da web (https://www.ultrasoft3d.net/Conversione_Coordinate.aspx)
            # Apro il sito per fare crawling
            browser.get(
                "https://www.ultrasoft3d.net/Conversione_Coordinate.aspx")
            browser.find_element_by_id(
                "Body_Coord_Conversion1_lst_source_1").click()
            browser.find_element_by_id(
                "Body_Coord_Conversion2_lst_source_4").click()

            # Seleziono l'elemento dell'est
            est = browser.find_element_by_id(
                "Body_Coord_Conversion1_txt_Longitudine")
            # Seleziono l'elemento del nord
            nord = browser.find_element_by_id(
                "Body_Coord_Conversion1_txt_Latitudine")
            # Seleziono l'elemento del fuso
            fuso = browser.find_element_by_id(
                "Body_Coord_Conversion1_txt_Fuso")

            # Pulisco i campi
            est.clear()
            nord.clear()
            fuso.clear()

            # Scrivo i valori nel campi
            est.send_keys(xest)
            nord.send_keys(xnord)
            fuso.send_keys(xzon)

            # conversione
            browser.find_element_by_id("Body_btn_Converti").click()

            # copia risultati
            est = browser.find_element_by_id(
                "Body_Coord_Conversion2_txt_Longitudine")
            est = est.get_attribute("value")
            nord = browser.find_element_by_id(
                "Body_Coord_Conversion2_txt_Latitudine")
            nord = nord.get_attribute("value")

            # browser.close()

            est = round(float(est), 6)
            nord = round(float(nord), 6)

            eltime = time.perf_counter() - eltime

            print("Codice: " + str(xid) + "| Nome:" + xname +
                  "| Est : " + str(est) + "| Nord : " + str(nord) + "| Time (s): " + str(eltime))

            di = {"Codice": i[0], "Nome": i[1], "Est": est, "Nord": nord}
            datajlist.append(di.copy())
            t.stop()

    xdf = pd.json_normalize(datajlist)
    filename = x.replace(".xlsx", "")+"_conv.xlsx"
    print("Salvataggio ["+filename+"] in corso")
    xdf.to_excel(poxlsx+filename)

browser.close()
