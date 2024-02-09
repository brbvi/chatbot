"""
# Descrever os passos manuais e depois transformar isso em código
# Ler planilha e guardar infos de nome, tel e vencimento.
# Criar links personalizados do whats e enviar mensagens para cada cliente com base nos dados da planilha

"""

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
#webbrowser.open('https://web.whatsapp.com/')
#sleep(10)
#carregar planilha
workbook = openpyxl.load_workbook('clientes.xlsx')
#chamar a página do arquivo
pagina_clientes = workbook['Planilha1']
#Lendo e capturando linha por linha a partir da linha 2
for linha in pagina_clientes.iter_rows(min_row=2):
    #nome, telefone e vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    valor = linha[3].value
    status = linha[4].value

    if status == 'NÃO-PAGO':
        mensagem = f'Olá, {nome} o deposito no valor de R${valor} com o vencimento para {vencimento.strftime('%d/%m/%Y')} já está disponivel. Favor fazer o pix pelo link xxxx-xxxx-xxxx-xxxx'
        #criando link personalizado 
  
        try:
            link_mensagem_whatsapp= f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(10)
            seta = pyautogui.locateCenterOnScreen('button.png')
            sleep(4)
            pyautogui.click(seta[0], seta[1])
            sleep(4)
            pyautogui.hotkey('ctrl','w')
            sleep(4)
        except:
            print(f'Não foi possivel enviar mensagem para {nome}')
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                arquivo.write(f'{telefone},{nome}')
    

   