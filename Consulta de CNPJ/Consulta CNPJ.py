import pandas as pd
import requests
import json
from time import sleep as slp

df = pd.read_excel(input(r'Informe o caminho do arquivo: '), dtype=str)

print('Consultando...')

with open('B_CNPJ.csv' , 'a', newline= '', encoding='UTF-8') as arquivo:
    
    for c in df['CNPJ']:
        
        browser = requests.get('https://www.receitaws.com.br/v1/cnpj/' + c)
        slp(3)
        
        resp = json.loads(browser.text)
        
        cep = resp['cep']
        nome = resp['nome']
        logradouro = resp['logradouro']
        numero = resp['numero']
        tipo = resp['tipo']
        
        arquivo.write(c + '_' + str(cep) + '_' + str(tipo) + '_' + str(nome) + '_' + str(logradouro) + '_' + str(numero))
        arquivo.write(str('\n'))
        slp(20)

print('Finalizado.')