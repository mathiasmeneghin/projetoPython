from selenium.common.exceptions import *
from selenium.webdriver.support import expected_conditions as CondicaoExperada
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import openpyexcel
import os
import pyautogui
import pyperclip




def obter_precos():

  def iniciar_driver():
    chrome_options = Options()
    arguments = ['--lang=pt-BR', '--window-size=1300,1000', '--incognito']
    for argument in arguments:
        chrome_options.add_argument(argument)

    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_setting_values.automatic_downloads': 1,

    })
    driver = webdriver.Chrome(service=ChromeService(
        ChromeDriverManager().install()), options=chrome_options)

    wait = WebDriverWait(
        driver,
        10,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )
    return driver, wait
    
  driver, wait = iniciar_driver()
  driver.get('https://www.pichau.com.br/search?q=memoria%20ram%2016&product_category=6450&sort=price-asc')
  precos_1 = wait.until(CondicaoExperada.visibility_of_all_elements_located((By.XPATH,"//div[@class='jss83']")))
  memoria_ram_1 = float(precos_1[0].text.split(' ')[1].replace(',','.'))

  driver.get('https://www.magazineluiza.com.br/busca/memoria+ram+16gb/?page=1&sortOrientation=asc&sortType=price')
  precos_2 = wait.until(CondicaoExperada.visibility_of_any_elements_located((By.XPATH,"//p[@data-testid='price-value']")))
  memoria_ram_2 = float(precos_2[0].text.split(' ')[1].replace(',','.'))

  return memoria_ram_1,memoria_ram_2  


def gerar_planilha_margem_de_lucro():
  memoria_ram_1,memoria_ram_2 = obter_precos()
  custo = 200
  site_1='https://www.pichau.com.br'
  site_2='https://www.magazineluiza.com.br'
  workbook = openpyexcel.Workbook()
  del workbook['Sheet']
  workbook.create_sheet('margem_lucro')
  sheet_margem_lucro = workbook['margem_lucro']
  sheet_margem_lucro.append(['Site','Custo','Preço','Lucro'])
  sheet_margem_lucro.append([site_1,custo,memoria_ram_1,memoria_ram_1 - custo])
  sheet_margem_lucro.append([site_2,custo,memoria_ram_2,memoria_ram_2 - custo])
  workbook.save('margem de lucro.xlsx')

  workbook_margem_lucro = openpyexcel.load_workbook('margem de lucro.xlsx')
  sheet_margem_lucro = workbook_margem_lucro['margem_lucro']
  margem_de_lucro = ''
  for linha in sheet_margem_lucro.iter_rows(min_row=1):
    margem_de_lucro += f'{linha[0].value},{linha[1].value},{linha[2].value},{linha[3].value}{os.linesep}'

  with open('margem_lucro.tx.','w',newline='',encoding='utf-8') as arquivo:
    arquivo.write(margem_de_lucro) 

  return margem_de_lucro 



def enviar_margem_lucro_no_whatsapp():

  margem_de_lucro = gerar_planilha_margem_de_lucro()
  
  pyautogui.moveTo(20,1062,duration=4)
  sleep(5)
  pyautogui.click()
  sleep(5)
  pyautogui.moveTo(170,460,duration=4)
  sleep(5)
  pyautogui.scroll(-5000)
  sleep(5)
  pyautogui.click(148,708,duration=4)
  sleep(5)
  botao_editar = pyautogui.locateCenterOnScreen('icone_editar.PNG')
  pyautogui.moveTo(botao_editar[0],botao_editar[1],duration=2)
  sleep(5)
  pyautogui.moveTo(678,186)
  sleep(5)
  pyautogui.click()
  sleep(5)
  pyautogui.write('teste')
  sleep(5)
  pyautogui.click(646,246)
  sleep(5)
  pyperclip.copy(margem_de_lucro)
  pyautogui.hotkey('ctrl','v')
  pyautogui.hotkey('enter')


enviar_margem_lucro_no_whatsapp()

  



