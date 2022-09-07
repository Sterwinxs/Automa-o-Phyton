from sys import displayhook
import pyautogui
import pyperclip
import time

pyautogui.click(x=610, y=740) #iniciando o navegador.
time.sleep(3)
pyautogui.click(x=589, y=45) #selecionando barra de pesquisa.
pyperclip.copy("https://docs.google.com/spreadsheets/d/1HSFFI4v7ju8qmy4eW12hdl1_BwkEImSw/edit#gid=1634402548")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter") #acessando relatório de vendas
time.sleep(5)
pyautogui.click(x=101, y=115)
time.sleep(2)
pyautogui.click(x=215, y=390)
time.sleep(2)
pyautogui.click(x=467, y=393)
time.sleep(2)
pyautogui.click(x=513, y=446) #realizando download
time.sleep(3)

import pandas as pd
tabela= pd.read_excel(r"C:\Users\gabri\Downloads\Vendas - Dez.xlsx") #importando os dados
displayhook(tabela)

faturamento= tabela["Valor Final"].sum()
quantidade= tabela["Quantidade"].sum()

#Enviando Email!
pyautogui.click(x=335, y=48)
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=85, y=159)
time.sleep(3)
pyperclip.copy("emailficticio@gmail.com")
pyautogui.hotkey('ctrl', "v")
pyautogui.press("tab")
time.sleep(2)
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

texto= f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:}

Abs
Gabrielly Castro"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
time.sleep(2)
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=1349, y=9)