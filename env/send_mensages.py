import pandas as pd 
import pywhatkit as kit 
import time 
import pyautogui as pg

tel_rafa = "https://wa.me/5511973029907"
tel_isa = "https://wa.me/5511975894895"


tpe_android = "https://play.google.com/store/apps/details?id=e.jcarlos.tpebh"
tpe_iphone = "testemunhometropolitano.com.br"
tpe_computador = "testemunhometropolitano.com.br"
tpe_formulario = "https://forms.gle/yPdDFCYZe2mVkcRU6"

#Carregar a tabela do Excel 
file_path = 'C:\\developer\\WhatsappMsg\\env\\Data\\teste.xlsx'
df = pd.read_excel(file_path)  # Lê a planilha Excel e armazena os dados em um DataFrame


tempo_espera = 5
#Iterar sobre cada linha da tabela 

for index, row in df.iterrows():

    nome_completo = row['Nome Completo']  # Obtém o nome da pessoa
    primeiro_nome = nome_completo.split()[0]  # Divide o nome completo e pega o primeiro nome
    numero = row['Numero']  # Obtém o número de telefone da pessoa
    prefixo = "+5511"
    # Converter o número para string e adicionar o prefixo
    numero_str = str(numero)  # Converte o número para string
    numero_com_prefixo = prefixo + numero_str.lstrip('+')  # Remove '+' do número se já estiver presente
   # Criar a mensagem com os vCards
    mensagem = f"""Oi, {primeiro_nome}, tudo bem? Estou dando uma força aqui para os irmãos do TPE....
Nós vimos aqui que sua petição consta como aprovada :D 
Te avisaram que existe um aplicativo do  "TPE" ??  Lá você pode escolher em que lugar quer trabalhar, o horário e inclusive pode fazer  arranjo com irmãos de diferentes circuitos :) 

Vou te passar os links para facilitar, blza? 
Se precisar  de qualquer ajuda para usar o site ou aplicativo, pode me chamar aqui.

FcJ
Rafael Ramires

Aplicativo para Android:
{tpe_android}
Se você estiver usando iPhone ou computador, pode acessar por este link:
{tpe_iphone}
"""
    
    

# Enviar mensagem pelo WhatsApp
    try:
        kit.sendwhatmsg_instantly(numero_com_prefixo, mensagem, 30, True, 10)  # Envia a mensagem pelo WhatsApp'
        
        pg.click(x=2831, y=703)  # Adjust the coordinates based on your screen
        time.sleep(1)
        pg.press("Enter")
        time.sleep(1)
        pg.hotkey('ctrl', 'w')
        
 
        
        print(f"Mensagem enviada para {nome_completo} ({numero_com_prefixo})")  # Exibe uma mensagem de sucesso

        time.sleep(10) 
        
    except Exception as e:
        print(f"Falha ao enviar mensagem para {nome_completo}: {e}")  # Exibe uma mensagem de erro



# print("Posicione o cursor do mouse na posição desejada em 5 segundos...")
# time.sleep(5)  # Tempo para você posicionar o cursor do mouse
# x, y = pg.position()  # Obtém a posição do cursor do mouse
# print(f"As coordenadas são: x={x}, y={y}")




