from asyncore import read
import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 0.5

# 1° passo: entrar no google drive:
pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v', 'enter')

# 2° passo: entrar na pasta das vendas:
time.sleep(3)
pyautogui.click(x=322, y=274, clicks = 2)

# 3° passo: baixar as vendas:
time.sleep(3)
pyautogui.click(x=322, y=274)
pyautogui.click(x=1715, y=162)
pyautogui.click(x=1504, y=555)

# 4° passo: calcular os indicadores (faturamento e quantidade de produtos vendidos):
time.sleep(5)
tabela = pd.read_excel(r'C:\Users\Samuel Lucas\Downloads\Vendas - Dez.xlsx')

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

# 5° passo: enviar para a empresa por meio de emails:
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/?tab=rm&ogbl#all')
pyautogui.hotkey('ctrl', 'v', 'enter')
time.sleep(5)
pyautogui.click(x=89, y=179)
pyperclip.copy('pythonimpressionador+diretoria@gmail.com')
pyautogui.hotkey('ctrl', 'v', 'enter')
pyautogui.click(x=1312, y=468)
pyperclip.copy('Faturamento e Quantidade de produtos vendidos em dezembro.')
pyautogui.hotkey('ctrl', 'v', 'enter')
pyautogui.click(x=1326, y=512)
pyperclip.copy(f'Olá!\n\n Aqui está o relatório de vendas que você solicitou:\n\n As vendas de dezembro concluíram com {quantidade:} de produtos vendidos e com um faturamento de R${faturamento:,.2f}.\n\nAbraços!\nSamuel Lucas')
pyautogui.hotkey('ctrl', 'v', 'enter')
pyautogui.hotkey('ctrl', 'enter')