import pandas as pd
from selenium import webdriver
from time import sleep as slp
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

servico = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service = servico)

browser.implicitly_wait(125)

df = pd.read_excel(r'C:\Users\ferna\Downloads\2022.000008306688-87\2022.000008306688-87.xlsx', dtype=str)

browser.get('https://efisco.sefaz.pe.gov.br/sfi_trb_cmt/PRConsultarNFTermoFielDepositario')
#browser.maximize_window()
slp(30)

with open('Log_PE.csv' , 'a', newline= '', encoding='UTF-8') as arquivo:

    for c in df['REGISTRO']:
        browser.find_element('xpath', '//*[@id="primeiro_campo"]').clear()
        browser.find_element('xpath', '//*[@id="primeiro_campo"]').send_keys(c)
        browser.find_element('xpath', '//*[@id="btt_localizar"]').click()
        browser.find_element('xpath', '//*[@id="btt_detalhar"]').click()
        chaveA = browser.find_element('xpath', '/html/body/form/table/tbody/tr[2]/td/div/table/tbody/tr/td/input').get_attribute('value')
        browser.back()
        browser.find_element('xpath', '//*[@id="bt_emitirgnre"]').click()
        browser.find_element('xpath', '//*[@id="DtPagamentoGNRE"]').click()
        browser.find_element('xpath', '//*[@id="DtPagamentoGNRE"]').send_keys('15122022')

        vDifal = browser.find_element('xpath', '//*[@id="VL_DIFAL_ABERTO_' + c + '"]').get_attribute('value')
        rSocial = browser.find_element('xpath', '//*[@id="PESSOA_NM_RAZAO_SOCIAL"]').get_attribute('value')
        cnjp = browser.find_element('xpath', '//*[@id="NuDocumentoEmitente"]').get_attribute('value')
        nNota = browser.find_element('xpath', '//*[@id="table_tabeladados"]/tbody/tr[3]/td[3]').get_attribute('class')

        if vDifal != '0,00':
            browser.find_element('xpath', '//*[@id="btt_confirmaremitirgnre"]').click()
            browser.find_element('xpath', '/html/body/div/form/div[2]/div/div/div/div[2]/p[2]/input').click()
            browser.back()
            browser.back()

        else:
            browser.find_element('xpath', '//*[@id="btt_calcaulardifalzerado"]').click()
            browser.find_element('xpath', '//*[@id="btt_confirmaremitirgnre"]').click()
            browser.find_element('xpath', '/html/body/div/form/div[2]/div/div/div/div[2]/p[2]/input').click()
            browser.back()
            browser.back()
            continue

        arquivo.write(c + '_' + str(cnjp) + '_' + str(rSocial) + '_' + str(chaveA) + '_' + '_' + str(vDifal))
        arquivo.write(str('\n'))

browser.quit()