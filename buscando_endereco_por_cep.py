from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

#utilizar comandos usando teclas
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.by import By

#usar para esperar entre as ações definidas
import pyautogui as tempoPausa
import pyautogui as atalhoTeclasTeclado

#imports do excel
from openpyxl import load_workbook
import os


#instala versão atual do webdriver do navegador(no caso, chrome)
servico = Service(ChromeDriverManager().install())

navegador = webdriver.Chrome(service = servico)


#-------------- abre o navegador, acessa o site, coleta as informacoes -----------------------
navegador.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

tempoPausa.sleep(5)

navegador.find_element(By.NAME, "endereco").send_keys("13215791")

tempoPausa.sleep(5)

navegador.find_element(By.NAME, "btn_pesquisar").click()

tempoPausa.sleep(5)

rua = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]').text

bairro = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[2]').text

cidade = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[3]').text

cep = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[4]').text

print(rua)
print(bairro)
print(cidade)
print(cep)

#-------------------------- Salvando no excel ---------------------------

#abre o arquivo .xlsx do caminho indicado e nome expecificado
nome_arquivo_cep = "E:\\dados_cep2.xlsx"
planilhaCriada = load_workbook(nome_arquivo_cep)

#variável que armazena a folha da planilha do excel onde os dados vão ser salvos
sheet_selecionada = planilhaCriada["Dados"]

#inicia na próxima linha vazia
linha = len(sheet_selecionada["A"]) + 1
colunaA = "A" + str(linha) #concatenando coluna e linha
colunaB = "B" + str(linha)
colunaC = "C" + str(linha)
colunaD = "D" + str(linha)


sheet_selecionada[colunaA] = rua
sheet_selecionada[colunaB] = bairro
sheet_selecionada[colunaC] = cidade
sheet_selecionada[colunaD] = cep

planilhaCriada.save(filename=nome_arquivo_cep)

os.startfile(nome_arquivo_cep)



