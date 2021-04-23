import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
import json
import win32com.client as win32
import os

#WEB SCRAPING:

pair = #tupla com email e senha

#Para rodar em bd: options = option
option = Options()
option.headless = True
driver = webdriver.Firefox(options=option)

# Pegar conteúdo html a partir da url
def login(logPair):
    url1 = "https://coinmarketcap.com/"
    driver.get(url1)
    time.sleep(2)
    #Logando:
    driver.find_element_by_xpath(
        "//div[@class='sc-1q2q8hd-0 kqdKYQ']//button[@class='sc-1ejyco6-0 eQMwpO']").click()
    email_area = driver.find_element_by_xpath(
        "//div[@class='f0b7mj-3 enpzGp']//input[@class='cxm5lu-0 bOgmnN']")
    password_area = driver.find_element_by_xpath(
        "//div[@class='f0b7mj-3 enpzGp last']//div[@class='sc-273q29-0 fnwcbv']//input[@class='cxm5lu-0 bOgmnN']")
    email_area.send_keys(logPair[0])
    password_area.send_keys(logPair[1])
    driver.find_element_by_xpath(
    "//div[@class='f0b7mj-6 dDaimk']//button[@class='sc-1ejyco6-0 iGZjcz']").click()
    time.sleep(3)


#Entrando na Watchlist e pegando tabela:
def get_table():
    url2 = "https://coinmarketcap.com/watchlist/"
    driver.get(url2)
    #Fecha pop-up
    driver.find_element_by_xpath(
        "//div[@class='sc-1u4r6ia-38 gYxhkk']//button[@class='sc-1ejyco6-0 czBWYA']").click()        
    #Pega a tabela
    element = driver.find_element_by_xpath("//div[@class='tableWrapper___3utdq cmc-table-watchlist-wrapper']//table")
    html_content = element.get_attribute('outerHTML')

    #Formatação:
    soup = BeautifulSoup(html_content, 'html.parser') #parseia a tabela e estrutura
    table = soup.find (name='table')
    df_full = pd.read_html(str(table))[0]
    #Arruma as colunas do Data Frame
    df = df_full[['Name', 'Price', '24h', 'Circulating Supply', 'Volume']] 
    df.columns = ['Nome', 'Preço', '24h %', 'Supply', 'Volume 24h']

    #Transforma dados em dicionário
    criptos = {}
    criptos = df.to_dict('records')

    #Arruma volume
    for coin in criptos:
        splitted = coin['Volume 24h'].split(',')
        formatted1 = ''

        for index, nums in enumerate(splitted):
            if len(nums) > 3 and '$' not in nums:
                nums1 = nums[:3]
                nums2 = nums[3:]
                formatted1 += nums1
                formatted2 = nums2
                for cripto_num in splitted[index +1:]:
                    formatted2 += cripto_num
                break
            else:
                formatted1 += nums

        formatted1 = "$" + "{:,}".format(int(formatted1[1:]))
        formatted2 = "{:,}".format(int(formatted2[:-4])) + formatted2[-4:]
        volume = formatted1 + '  /  ' + formatted2
        coin['Volume 24h'] = volume

    #Arruma nome
    for coin in criptos:
        splitted = list(coin['Nome'])
        formatName = ''
        for index, char in enumerate(splitted):
            if char.isdigit():
                withoutNumbers = splitted[:index]
                formatName = ''.join(withoutNumbers)
                break
        coin['Nome'] = formatName

    #Fecha
    driver.quit()
    return criptos

login(pair)
criptos = get_table()

#Salva os dados como json
#json = json.dumps(criptos)
#fp = open('tabela.json', 'w')
#fp.write(json)
#fp.close()


#WRITE TO EXCEL:

#json.loads(open('tabela.json').read())

#Formatando dados para passar pro excel
rows = []

for detail in criptos:
    name = detail['Nome']
    supply = detail['Supply']
    volume_24 = detail['Volume 24h']
    percentage_24 = detail['24h %']
    preco = detail['Preço']
    rows.append([name, preco, volume_24, percentage_24, supply])

#Open excel file
ExcelApp = win32.gencache.EnsureDispatch('Excel.Application')
ExcelApp.Visible = True
path =  os.getcwd().replace('\'','\\') + '\\'
workBook = ExcelApp.Workbooks.Open(path+'Portfólio.xlsx')
workSheet = workBook.Worksheets(1)

#Insert records
row_tracker = 2
column_size = 5

for row in rows:
    workSheet.Range(
        workSheet.Cells(row_tracker, 1),
        workSheet.Cells(row_tracker, column_size)
    ).Value = row
    row_tracker += 1

workBook.Save() 

#Para fechar arquivo:
#workBook.Close()
#ExcelApp.Quit()
#ExcelApp = None

#Pra salvar em outro arquivo
#workBook.SaveAs(os.path.join(os.getcwd(), 'Portfólio.xlsx'), 51) 
