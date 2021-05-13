import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
import json
import win32com.client as win32
import os
import PySimpleGUI as sg
import schedule

#WEB SCRAPING:
class WebScraping():
    def __init__(self): #futuramente opções de outros navegadores       
        #Para rodar em bd: options = option
        option = Options()
        option.headless = True
        self.driver = webdriver.Firefox()

# Pegar conteúdo html a partir da url
    def ws_login(self, logPair):
        url1 = "https://coinmarketcap.com/"
        self.driver.get(url1)
        time.sleep(3)
        #Logando:
        self.driver.find_element_by_xpath(
            "//div[@class='sc-1q2q8hd-0 kqdKYQ']//button[@class='sc-1ejyco6-0 eQMwpO']").click()
        email_area = self.driver.find_element_by_xpath(
            "//div[@class='f0b7mj-3 enpzGp']//input[@class='cxm5lu-0 bOgmnN']")
        password_area = self.driver.find_element_by_xpath(
            "//div[@class='f0b7mj-3 enpzGp last']//div[@class='sc-273q29-0 fnwcbv']//input[@class='cxm5lu-0 bOgmnN']")
        email_area.send_keys(logPair[0])
        password_area.send_keys(logPair[1])
        self.driver.find_element_by_xpath(
            "//div[@class='f0b7mj-6 dDaimk']//button[@class='sc-1ejyco6-0 iGZjcz']").click()
        time.sleep(3)


    #Entrando na Watchlist e pegando tabela:
    def ws_get_table(self):
        url2 = "https://coinmarketcap.com/watchlist/"
        self.driver.get(url2)
        time.sleep(3)
        #Fecha pop-up
        try:
            self.driver.find_element_by_xpath(
                "//div[@class='sc-1u4r6ia-38 gYxhkk']//button[@class='sc-1ejyco6-0 czBWYA']").click()     
        except:
            pass   
        #Pega a tabela
        element = self.driver.find_element_by_xpath("//div[@class='tableWrapper___3utdq cmc-table-watchlist-wrapper']//table")
        html_content = element.get_attribute('outerHTML')

        #Formatação:
        soup = BeautifulSoup(html_content, 'html.parser') #parseia a tabela e estrutura
        table = soup.find (name='table')
        df_full = pd.read_html(str(table))[0]
        #Arruma as colunas do Data Frame
        df = df_full[['Name', 'Price', '24h %', 'Circulating Supply', 'Volume(24h)']] 
        df.columns = ['Nome', 'Preço', '24h %', 'Supply', 'Volume 24h']

        #Transforma dados em dicionário
        criptos = {}
        criptos = df.to_dict('records')

        #Arruma volume
        try:
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
                for i, alg in enumerate(formatted2):
                    if alg.isalpha():
                        formatted2 = "{:,}".format(int(formatted2[:i])) + " " + formatted2[i:]
                        break
                volume = formatted1 + '  /  ' + formatted2
                coin['Volume 24h'] = volume
        except:
            pass
                
        #Arruma nome
        try:
            for coin in criptos:
                splitted = list(coin['Nome'])
                formatName = ''
                for index, char in enumerate(splitted):
                    if char.isdigit():
                        withoutNumbers = splitted[:index]
                        formatName = ''.join(withoutNumbers)
                        break
                coin['Nome'] = formatName
        except:
            pass
        
        #Fecha
        self.driver.quit()
        return criptos

#WRITE TO EXCEL:
class ToExcel(): 
    def __init__(self, cripto_dict):
        self.rows = []
        self.cripto_dict = cripto_dict

    #Formatando dados para passar pro excel
    def format_data(self):
        for detail in self.cripto_dict:
            name = detail['Nome']
            supply = detail['Supply']
            volume_24 = detail['Volume 24h']
            percentage_24 = detail['24h %']
            preco = detail['Preço']
            self.rows.append([name, preco, volume_24, percentage_24, supply])

    #Inserindo dados na planilha
    def insert_data(self, file):
        #Open excel file
        ExcelApp = win32.gencache.EnsureDispatch('Excel.Application')
        ExcelApp.Visible = True
        #Path local
        path =  os.getcwd().replace('\'','\\') + '\\'
        workBook = ExcelApp.Workbooks.Open(path+file)
        workSheet = workBook.Worksheets(1)
        workBook.Save() 
        
        #Insert records
        row_tracker = 2
        column_size = 5

        for row in self.rows:
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

#GRAPHICAL INTERFACE
class GUI():
    def __init__(self):
        #Pega dados salvos
        data = ["", "", "False", "", ""]
        with open('data.json') as json_file:
            data = json.load(json_file)
        #formata para booleano
        if data[2] == 'True':
            data[2] = True
        else:
            data[2] = False
        sg.theme('Reddit')
        #Layout
        layout = [
            [sg.Text('Cripto Update Excel')],
            [sg.Text('Email: ', size = (5,0)), sg.InputText(data[0], key='email')],
            [sg.Text('Senha: ', size = (5,0)), sg.InputText(data[1], key='senha', password_char = '*')],
            [sg.Checkbox('Lembrar de mim', key='remember', default = data[2])], #default = data[2])
            [sg.Text('Frequência: '), sg.InputText(data[3], size = (6,0),
                                                key='update'), sg.Text('min')],                                   
            [sg.Button('Iniciar ciclo de atualizações')],
            [sg.Output(size=(40,5))]
        ]
        #Janela
        self.window = sg.Window('Cripto Update Excel').layout(layout)


    def Iniciar(self, file):
        while True:
            self.event, self.values = self.window.Read()
            #Botão x
            if self.event is None:
                break
            #Troubleshooting
            if notNone(self.values):
                email = self.values['email']
                senha = self.values['senha']
                remember = self.values['remember']
                frequencia = self.values['update']

                def update():
                    web_scraping = WebScraping()
                    web_scraping.ws_login((email, senha))
                    if remember == True:
                        with open ('data.json', 'w') as f:
                            json.dump([email, senha, 'True', frequencia], f)
                    else:
                        with open ('data.json', 'w') as f:
                            json.dump(['','', 'False', ''], f)
                    try:
                        criptos = web_scraping.ws_get_table()
                        print('Atualização')
                        to_excel = ToExcel(criptos)
                        to_excel.format_data()
                        to_excel.insert_data(file)
                        del web_scraping
                    except:
                        print('Erro')
                        quit_web(web_scraping)
                
                def quit_web(obj):
                    obj.driver.quit()
                
                update()
                if frequencia.isnumeric():
                    schedule.every(int(frequencia)).minutes.do(update)
                    while True:
                        schedule.run_pending()
                        time.sleep(1)
                else:
                    print('Input Inválido')
            else:
                print('Input Inválido')

            

def notNone(obj):
    for key, value in obj.items():
        if value == '':
            return False
    return True    

user_interface = GUI()
user_interface.Iniciar('Portfólio.xlsx')

#browse e debugg



