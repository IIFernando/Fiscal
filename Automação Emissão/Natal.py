import pandas as pd
from selenium import webdriver
from time import sleep as slp
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
servico = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service = servico)

browser.implicitly_wait(125)
cont = 0

df = pd.read_excel(r'C:\Users\ferna\Downloads\Sefaz RN.xlsx', dtype=str)

browser.maximize_window()
browser.get('https://uvt2.set.rn.gov.br/#/services/icms-antecipado')

slp(20)

for n in df['Nota']:
        
    c =  df['CNPJ']

    browser.find_element('xpath', '/html/body/section/div/div[4]/div[2]/form/div/div[1]/div[1]/div/input').send_keys(n)
    slp(1)
    browser.find_element('xpath', '/html/body/section/div/form/div/div[1]/div[1]/div/input').send_keys(c[cont])
    slp(1)
    browser.find_element('xpath', '/html/body/section/div/div[4]/div[2]/form/div/div[2]/div/input').send_keys(c[cont])
    slp(1)
    browser.find_element('xpath', '/html/body/section/div/form/div/div[1]/div[2]/div/div/input').send_keys('15122022')
    slp(1)
    browser.find_element('xpath', '/html/body/section/div/div[6]/div/input').click()
    slp(1)
    vDifal = browser.find_element('xpath', '/html/body/section/div/div[3]/div/div[3]/div/div/input').get_attribute('value')
    
    with open('Log_RN.csv' , 'a', newline= '', encoding='UTF-8') as arquivo:
        
        if vDifal != '0,00':
            browser.find_element('xpath', '/html/body/section/div/div[3]/div/div[4]/div/div[2]/label/input').click()
            slp(1)
            browser.find_element('xpath', '/html/body/section/div/div[4]/input').click()
            slp(1)
            browser.find_element('xpath', '/html/body/div[1]/div/div/div[3]/button').click()
            slp(1)
            
            browser.back()
            slp(1)
            browser.back()
            slp(1)
            cont += 1
            
            arquivo.write(n + '_' + '_' + str(vDifal) + '_' + str('Gerada'))
            arquivo.write(str('\n'))
            
        elif vDifal == '0,00':
            arquivo.write(n + '_' + str('Não Gerada'))
            arquivo.write(str('\n'))
            
            cont += 1
            browser.back()
            slp(1)

browser.quit()