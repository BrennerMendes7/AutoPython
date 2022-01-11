# Link da documentação do pyautogui: https://pyautogui.readthedocs.io/en/latest/quickstart.html
# Passo a passo para resolver o nosso desafio
# pip install pyautogui, pyperclip, pandas, numpy e openpyxl
import pyautogui
import pyperclip
import pandas as pd
import time
# Passo 1: Entrar no link do drive
pyautogui.PAUSE = 1 # tempo de intervalo entre as ações

pyautogui.press("win") # pressiona a tecla windows
pyautogui.write("chrome") # escreve o nome do app chrome
time.sleep(2) # espera 2 segundos
pyautogui.press("enter") # pressiona a tecla enter
time.sleep(8) # espera 8 segundos
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga") # copiar link do site pelo pyperclip por causa de caracteres especiais
pyautogui.hotkey("ctrl", "v") # usa atalho de colar
pyautogui.press("enter") # pressiona a tecla enter
time.sleep(15) # espera 15 segundos

# Passo 2: Navegar no sistema (até a pasta Exportar)
pyautogui.click(x=366, y=267, clicks=2) # clica na posição do mouse na tela duas vezes
time.sleep(10) # espera 10 segundos

# Passo 3: Fazer o download da base de vendas
pyautogui.click(x=374, y=320) # clica na posição do mouse na tela
pyautogui.click(x=1073, y=163) # clica na posição do mouse na tela
pyautogui.click(x=872, y=560) # clica na posição do mouse na tela
time.sleep(17) # espera 17 segundos

# Passo 4: Importar a base de vendas pro Python
tabela = pd.read_excel(r"C:\Users\Windows\Downloads\Vendas - Dez.xlsx") # ler arquivo excel no diretório que está localizado

# Passo 5: Calcular o faturamento e quantidade de produtos vendidos (de indicadores)
faturamento = tabela["Valor Final"].sum() # somar todos os valores totais da planilha, usar colchetes para acessar a coluna da planilha
qtde_produtos = tabela["Quantidade"].sum() # somar a quantidade de todos os produtos da planilha

# Passo 6: Enviar email
pyautogui.hotkey("ctrl", "t") # usa atalho para criar nova aba no navegador
time.sleep(8) # espera 8 segundos
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox") # copia link do gmail pelo pyperclip por causa de caracteres especiais
pyautogui.hotkey("ctrl", "v") # usa atalho de colar
pyautogui.press("enter") # pressiona a tecla enter
time.sleep(15) # espera 15 segundos
pyautogui.click(x=34, y=168) # clica na posição do mouse na tela
pyautogui.write("algumemail@gmail.com") # escreve o destino do email
pyautogui.press("tab") # pressiona a tecla tab
pyautogui.press("tab") # pressiona a tecla tab
pyperclip.copy("Relatório de Vendas") # escreve o assunto do email
pyautogui.hotkey("ctrl", "v") # colar o assunto do email
pyautogui.press("tab") # pressiona a tecla tab
# escrever o email
texto = """ 
Prezados, bom dia

O faturamento de ontem foi de R$: {faturamento:,.2f}
A quantidade de produtos foi de: {qtde_produtos:,}

Abs
Brenner
"""
# {} - especificar variável dentro de uma string   : - expressão para formatar   , - separador de milhar   . - separador de casas decimais (1f, 2f, 3f ...)
pyperclip.copy(texto) # copiar a variável texto pelo pyperclip por causa dos caracteres especiais
pyautogui.hotkey("ctrl", "v") # colar a variável texto no campo de conteúdo do email
pyautogui.hotkey("ctrl", "enter") # atalho para enviar o email