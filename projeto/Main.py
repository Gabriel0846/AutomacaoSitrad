import pyautogui
import subprocess
import time
import os

def caminho_imagem(imagem_nome):
    return os.path.join(os.path.dirname(__file__), 'img', imagem_nome)

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
