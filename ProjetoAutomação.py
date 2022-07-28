#Passo 1: Entrar no sistema da empresa (o sistema nesse caso sera entrar no link do drive)
#Passo 2: Navegar no sistema e encontrar a base de dados (Entrar na pasta exportar)
#Passo 3: Realizar download da base de dados
#Passo 4: Calcular indicadores (fatturamento, quantidade de produtos vendidos)
#Passo 5: Entrar no email
#Passo 6: Mandar um email para diretoria com os indicadores

import time
import pyautogui
import pyperclip
import pandas as pd

#Passo 1: Entrar no sistema da empresa (o sistema nesse caso sera entrar no link do drive)
pyautogui.PAUSE = 1.2
 
pyautogui.press('win') 
pyautogui.write('opera') #esses três comandos irão abrir o aplicativo do chrome.
pyautogui.press('enter')

pyperclip.copy('https://drive.google.com/drive/folders/14oLE59U1RqyRqlBbKpsyymW-mitvbtoh')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
#Passo 2: Navegar no sistema e encontrar a base de dados (Entrar na pasta exportar)
time.sleep(1)
pyautogui.click(x=403, y=350) #adiciona duplo click na posição informada
#Passo 3: Realizar download da base de dados
pyautogui.click(x=1717, y=164) #Clica nos tres pontinhos
pyautogui.click(x=1491, y=553) #clica para fazer download
time.sleep(7) #tempo para esperar o download
#Passo 4: Calcular indicadores (fatturamento, quantidade de produtos vendidos)
tabela = pd.read_excel('C:/Users/Funari/Downloads/Vendas - Dez.xlsx')

quantidade = tabela['Quantidade'].sum() #Soma os valores da coluna Quantidade
faturamento = tabela['Valor Final'].sum() #Soma os valores da coluna Valor Final

#Passo 5: Entrar no email
pyautogui.hotkey('ctrl','t')

pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(5)
#Passo 6: Mandar um email para diretoria com os indicadores
pyautogui.click(x=102, y=176)
pyautogui.write('felipefunari2015@gmail.com')
pyautogui.press('enter')
pyautogui.press('tab')
pyautogui.write('Indicadores Faturamento e Quantidade')
pyautogui.press('tab')
texto = f"""
Prezados, bom dia!

Segue Quantidade e Faturamento
Quantidade: {quantidade}
Faturamento: R${faturamento}

Att,
Felipe Funari.
"""
pyautogui.write(texto)
pyautogui.press('tab')
pyautogui.press('enter')