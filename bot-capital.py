import time
from openpyxl import Workbook
import openpyxl.styles as opStyles
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# ------ DEFINIENDO DIFERENTES MANERAS DE BUSCAR LA CAPITAL -------

paths_capital = [
    ["Capital de ",'//*[@id="rso"]/div[1]/div/block-component/div/div[1]/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div[1]/a'],
    ["Capital ", '//*[@id="rso"]/div[1]/div/block-component/div/div[1]/div[1]/div/div/div[1]/div/div/div[2]/div/div/div/div[1]/a'],
    ["",'//*[@id="_U35nZbOPKaOMwbkPr6WSsAw_66"]/div/div/div[2]/div/div/div/div[2]/div/div/div/span[2]/span/a'],
]

# ------ CREANDO NUEVO EXCEL ------ 

newExcel = Workbook()
newSheet = newExcel.active
newSheet.title = "Paises - capitales"

newSheet.append(("País", "Capital"))

#Se le da estilos a las celdas que contienen los titulos
newSheet["A1"].font = opStyles.Font(name="Arial",bold=True)
newSheet["B1"].font = opStyles.Font(name="Arial",bold=True)

#Se crea una lista vacia para luego guardar los datos que se obtienen
newData = []

# ------ ABRIENDO EL NAVEGADRO ------ 

driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://google.com/")

# ------ LEYENDO EXCEL -------------

data_countries = pd.read_excel('paises.xlsx')

for row in range(len(data_countries)):
    countries = data_countries.loc[row]
    for country in countries:
        data = ()
        
        # Se recorre las diferentes rutas para buscar la capital
        for path in paths_capital:
            # Por defecto la capital es: "No encontrada"
            capital = "No encontrada"
            
            search = driver.find_element(By.XPATH, '//*[@id="APjFqb"]')
            search.clear()
            search.send_keys(f"{path[0]}{country}", Keys.ENTER)
            
            #En caso de encontrar la capital se para el ciclo
            # Si no se encuentra, se sigue el ciclo
            try:
                capital = driver.find_element(By.XPATH, path[1]).text
                break
            except:
                pass
            
        # Se agregan los datos en una tupla y se añaden a la lista "newData"
        data = (country, capital)
        newData.append(data)

for nD in newData:
    newSheet.append(nD)
    
newExcel.save("paises-capital.xlsx")
print(f"El resultado se ha guardado en 'paises-capital.xslx'")

time.sleep(3)
driver.close()