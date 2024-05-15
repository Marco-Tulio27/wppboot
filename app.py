import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
webbrowser.open('https://web.whatsapp.com')
sleep(10)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Olá {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://pagarboleto.com'
    link_mensagem_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_wpp)
    sleep(10)
    try:
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w', interval=0.1)
        sleep(5)
    except:
        print(f'nao foi possivel enviar mensagem para {nome}')
        with open ('erros.csv','a',newline='',encoding='utf8') as arquivo:
            arquivo.write(f'{nome},{telefone}')
        pyautogui.hotkey('ctrl', 'w', interval=0.1)