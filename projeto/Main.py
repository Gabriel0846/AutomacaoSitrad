import pyautogui
import subprocess
import time
import os
from datetime import datetime, timedelta

# facilidar codar o caminho para os arquivos
def caminho_imagem(imagem_nome):
    return os.path.join(os.path.dirname(__file__), 'img', imagem_nome)

# retorna a data inicial do relatório
def data_inicial():
    hoje = datetime.now()
    if hoje.weekday() == 0:
        data_inicial = hoje - timedelta(days=2)
    else:
        data_inicial = hoje - timedelta(days=1)
    return data_inicial.strftime('%d%m%Y')

# retorna a data final de um dia atrás
def data_final():
    ontem = datetime.now() - timedelta(days=1)
    return ontem.strftime('%d%m%Y')

# função que insere as datas
def inserir_data():
    data_inicio = data_inicial()
    pyautogui.write(data_inicio)
    pyautogui.press('tab')
    time.sleep(1)

    data_fim = data_final()
    pyautogui.write(data_fim)
    time.sleep(1)

def gerar_relatorio(nome_sala, sensores):
    localizar_clicar(caminho_imagem('arquivo.png'))


# Abrir o programa
caminho_app = r'C:\Program Files (x86)\Full Gauge\SitradRemote\SitradRemote.exe'
subprocess.Popen(caminho_app)
time.sleep(10)

# digitar o login e senha
imagem_servidor = caminho_imagem('servidor.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_servidor)
x, y = imagem_localizada
pyautogui.doubleClick(x, y)
time.sleep(1)
pyautogui.write('admin123')
pyautogui.press('enter')
time.sleep(8)

# clica em uma parte do programa para carregar os termometros
imagem_carregar = caminho_imagem('carregar.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_carregar)
x, y = imagem_localizada
pyautogui.click(x, y)
time.sleep(1)

# entrar na tela que puxa os relatórios
imagem_arquivo = caminho_imagem('arquivo.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_arquivo)
x, y = imagem_localizada
pyautogui.click(x,y)
imagem_gerador = caminho_imagem('gerador_de_relatorios.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_gerador)
x, y = imagem_localizada
pyautogui.click(x, y)
time.sleep(10)

# puxar relátorio da anticamara
imagem_anticamara = caminho_imagem('anti_camara.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_anticamara)
x, y = imagem_localizada
pyautogui.doubleClick(x, y)
time.sleep(2)
inserir_data()
imagem_temperaturaAmbiente = caminho_imagem('temperatura_ambiente.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_temperaturaAmbiente)
x, y = imagem_localizada
pyautogui.click(x, y)
imagem_dispositivo = caminho_imagem('dispositivo.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_dispositivo)
x, y = imagem_localizada
pyautogui.click(x, y)
imagem_adicionar = caminho_imagem('adicionar.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_adicionar)
x, y = imagem_localizada
pyautogui.click(x, y)
imagem_gerar = caminho_imagem('gerar.png')
imagem_localizada = pyautogui.locateCenterOnScreen(imagem_gerar)
x, y = imagem_localizada
pyautogui.click(x, y)







