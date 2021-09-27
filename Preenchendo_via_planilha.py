#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from openpyxl import load_workbook

import os
from selenium import webdriver as opSeleneium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoEspera

caminhoArquivo = "C:\\Users\\User\\Desktop\\RPA1\\Formulario 2\\DadosFormulario.xlsx"

plan_open = load_workbook(filename=caminhoArquivo)

sheet_selecionada = plan_open['Dados']

for i in range(2, len(sheet_selecionada['A']) + 1):
    
    nome=sheet_selecionada['A%s'% i].value
    email=sheet_selecionada['B%s'% i].value
    telefone=sheet_selecionada['C%s'% i].value
    opcao=sheet_selecionada['D%s'% i].value
    esporte=sheet_selecionada['E%s'% i].value
    
    tempoEspera.sleep(6)
    
    navegadorform = opSeleneium.Chrome()
    navegadorform.get("https://pt.surveymonkey.com/r/H2NDTB7")


    tempoEspera.sleep(4)

    navegadorform.find_element_by_name("708837289").send_keys(nome)

    navegadorform.find_element_by_name("708837290").send_keys(email)

    navegadorform.find_element_by_name("708837291").send_keys(telefone)

    navegadorform.find_element_by_name("708837293").send_keys(esporte)
    
    if opcao == "Sim":
        navegadorform.find_element_by_id("708837292_4661988045_label").click()
    else:
        navegadorform.find_element_by_id("708837292_4661988046_label").click()

    tempoEspera.sleep(6)
    navegadorform.find_element_by_xpath('//*[@id="patas"]/main/article/section/form/div[2]/button').click()

