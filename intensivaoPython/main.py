import pandas
import pyautogui
import pyperclip
from time import sleep
import pandas as pd

pyautogui.PAUSE = 1.5

pyautogui.alert('Começando o programa')
pyautogui.press('win')
pyautogui.write('opera')
pyautogui.press('enter')

link = r'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'

pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
sleep(3)
pyautogui.doubleClick(394, 298)
sleep(2)
pyautogui.click(394, 298, button='right')
pyautogui.click(574, 680)


tabela = pandas.read_excel(r'C:\Users\Desktop\Downloads\Vendas - Dez.xlsx')
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()


pyautogui.hotkey('ctrl', 't')

link = r'mail.google.com/mail/u/0/#inbox'

pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
sleep(2)
pyautogui.click(x=145, y=199)
sleep(2)
pyautogui.write('joaovmt1905+diretoria@gmail.com')
pyautogui.press('tab') # confirmar email
pyautogui.press('tab') # passar para o assunto
sleep(2)

msg = r'Relatório de Vendas'

pyperclip.copy(msg)
pyautogui.hotkey('ctrl', 'shift', 'v')
sleep(2)
pyautogui.press('tab') # passar para o corpo do email

mensagem = f'''Relatório de Vendas de Ontem

Prezados, bom dia

O faturamento de ontem foi de R${faturamento:,.2f}
A quantidade de produtos foi de {quantidade:,}

Abs
John'''


pyperclip.copy(mensagem)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')
pyautogui.alert('Encerrando, o controle do computador voltou a ser seu')
