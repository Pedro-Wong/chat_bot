
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import pywhatkit as kit

# Ler planilhas e guardar informações
# webbrowser.open('https://web.whatsapp.com/')
# sleep(5)
workbook = openpyxl.load_workbook('clientes.xlsx')
clientes = workbook['Planilha1']

 
for linha in clientes.iter_rows(min_row=2):
    
    nome = linha[5].value
    telefone1 = linha[0].value
    telefone2 = linha[1].value
    telefone3 = linha[2].value
    telefone4 = linha[3].value

    if nome == None:
        break

    print(nome)
    # print('55'+str(telefone1))
    # print(telefone1)
    # print(telefone2)
    # print(telefone3)    
    
    
    
    # print(telefone4)
    # mensagem = f'Olá {nome}, boa tarde !'
    conv_tel = '+55'+str(telefone1)
    print(conv_tel)
    sleep(30) 
 
 
    kit.sendwhats_image(conv_tel, "anuncio.png", "Olá! Me chamo Pedro e estou mandando essa mensagem para te ajudar. Está precisando de dinheiro rápido, mas não tem margem?? Posso te ajudar a conseguir esse dinheiro e ainda benefícios. Só responder essa mensagem que entro em contato com as informações detalhadas. ", 7 )
    sleep(4)
    pyautogui.hotkey('ctrl', 'w')

    # try:
    #     link_wpp = f'https://web.whatsapp.com/send?phone={'55'+str(telefone1)}&text={quote(mensagem)}'
    #     webbrowser.open(link_wpp)
    #     sleep(11)
    #     seta = pyautogui.locateCenterOnScreen('seta.png')
    #     sleep(5)
    #     pyautogui.click(seta[0], seta[1]) 
    #     sleep(5)
    #     pyautogui.hotkey('ctrl', 'w')
    #     sleep(5)
    # except:
    #     print(f'não foi possível enviar a mensagem para {nome} e telefone {telefone1} ')
    #     with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
    #         arquivo.write(f'{nome}, {telefone1}')






# https://web.whatsapp.com/send?phone=55126445&text="dsfsd"
