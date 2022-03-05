import pandas as pd
from openpyxl import load_workbook

class Folie:
    def __init__(self, pytel1, pytel2, pytel3, index):
        self.pytel1 = pytel1
        self.pytel2 = pytel2
        self.pytel3 = pytel3
        self.index = index


# Parsujeme xlsx soubor
# Deklarace poli
folie = []
pytel_1 = []
pytel_2 = []
pytel_3 = []

# Nastaveni souboru a listu
soubor = "/Users/obabec/development/personal/jakub-python-demo/Data_Folie.xlsx"
listerino = "List1"

# Nacteni excel souboru
dfs = pd.read_excel(soubor, sheet_name=listerino)

for i in dfs.index:
    folie.append(Folie(dfs['1_Pytel'][i], dfs['2_Pytel'][i], dfs['3_Pytel'][i], i))

for i in dfs.index:
    pytel_1.append(dfs['1_Pytel'][i])
    pytel_2.append(dfs['2_Pytel'][i])
    pytel_3.append(dfs['3_Pytel'][i])

# Zapisu zmeny do pole

folie[2].pytel1 = "Ahoj"
folie[2].pytel2 = "Svete"
folie[2].pytel3 = "!"

# Zapisu zmeny do datoveho ramce
for i in dfs.index:
    dfs['1_Pytel'][i] = folie[i].pytel1
    dfs['2_Pytel'][i] = folie[i].pytel2
    dfs['3_Pytel'][i] = folie[i].pytel3


FilePath = soubor
ExcelWorkbook = load_workbook(soubor)
with pd.ExcelWriter(FilePath, engine = 'openpyxl') as writer:
    writer.book = ExcelWorkbook
    dfs.to_excel(writer, sheet_name = 'List2')

print("XXX");