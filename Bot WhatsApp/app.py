# https://web.whatsapp.com/send?phone=&text&
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
sleep(25)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Planilha1']
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    aniversario = linha[2].value
    mensagem = f'Olá {nome}, tudo bem? Vimos que está chegando uma data muito especial {aniversario.strftime(
        '%d/%m/%Y')}, por isso decidimos te presentear com o cupom: MEUNIVER, corra e aproveite.'

    try:
        mensagem_wp = f'https://web.whatsapp.com/send?phone={
            telefone}&text={quote(mensagem)}'
        webbrowser.open(mensagem_wp)
        sleep(30)
        seta = pyautogui.locateCenterOnScreen('enviar .png')
        sleep(10)
        pyautogui.click(seta[0], seta[1])
        sleep(10)
        pyautogui.hotkey('ctrl', 'w')
        sleep(10)
    except:
        print(f'Não foi possível enviar a mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{telefone},{nome}')
