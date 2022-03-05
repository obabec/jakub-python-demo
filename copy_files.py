import pandas as pd
import shutil
import os
import glob
# Nastaveni souboru a listu
vstupni_tabulka = "Seznam-dilu.xlsx"
listerino = "Sheet1"

vstupni_slozka = "input"
vystupni_slozka = "output"

soubory_na_presun = []

# Nacteni excel souboru
dfs = pd.read_excel(vstupni_tabulka, sheet_name=listerino)

for i in dfs.index:
    soubory_na_presun.append(dfs['Part_number'][i])

soubory_na_presun.append("170-105")

soubory_na_presun = set(soubory_na_presun)

soubory_aktualni = os.listdir(vstupni_slozka)
full_files = []

for file in soubory_na_presun:
    for act_soubor in soubory_aktualni:
        if file == act_soubor.split('.')[0]:
            full_files.append(act_soubor)


for file in full_files:
    shutil.copyfile(vstupni_slozka + "/" + file, vystupni_slozka + "/" + file)

print("xx")