import pyautogui
import subprocess
import time
import os
from datetime import datetime, timedelta

# Facilita o caminho para os arquivos
def caminho_imagem(imagem_nome):
    return os.path.join(os.path.dirname(__file__), 'img', imagem_nome)

# Retorna a data inicial do relatório
def data_inicial():
    hoje = datetime.now()
    data_inicial = hoje - timedelta(days=2) if hoje.weekday() == 0 else hoje - timedelta(days=1)
    return data_inicial.strftime('%d%m%Y')

# Retorna a data final de um dia atrás
def data_final():
    ontem = datetime.now() - timedelta(days=1)
    return ontem.strftime('%d%m%Y')

# Função que insere as datas
def inserir_data():
    pyautogui.write(data_inicial())
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write(data_final())
    time.sleep(1)

# Função para clicar em uma imagem
def clicar_imagem(imagem, duplo_clique=False):
    localizacao = pyautogui.locateCenterOnScreen(caminho_imagem(imagem))
    if localizacao:
        x, y = localizacao
        if duplo_clique:
            pyautogui.doubleClick(x, y)
        else:
            pyautogui.click(x, y)
        time.sleep(1)

# Abrir o programa
caminho_app = r'C:\Program Files (x86)\Full Gauge\SitradRemote\SitradRemote.exe'
subprocess.Popen(caminho_app)
time.sleep(10)

# Digitar o login e senha
clicar_imagem('servidor.png', duplo_clique=True)
pyautogui.write('admin123')
pyautogui.press('enter')
time.sleep(8)

# Clica em uma parte do programa para carregar os termômetros
clicar_imagem('carregar.png')

# Entrar na tela que puxa os relatórios
clicar_imagem('arquivo.png')
clicar_imagem('gerador_de_relatorios.png')
time.sleep(10)

# Puxar relatório da anticâmara
clicar_imagem('anti_camara.png', duplo_clique=True)
time.sleep(2)
inserir_data()
clicar_imagem('temperatura_ambiente.png')
clicar_imagem('dispositivo.png')
clicar_imagem('adicionar.png')
clicar_imagem('gerar.png')
time.sleep(10)
clicar_imagem('ferramentas.png')
clicar_imagem('relatoriotexto.png')
clicar_imagem('anticamaratexto.png')
time.sleep(10)

# Selecionar e copiar o texto gerado
pyautogui.hotkey('ctrl', 'a')
pyautogui.hotkey('ctrl', 'c')

# Caminho para o arquivo da planilha existente
caminho_planilha = r'\\Srv-araovos\d\Meus documentos\INFO\INFO\Controles\sitrad controladores.xlsx'

# Abrir Excel com a planilha específica
subprocess.Popen([r'C:\Program Files\Microsoft Office\Office16\EXCEL.EXE', caminho_planilha])
time.sleep(10)

# Navegar para a aba "ANTICAMARA (13)"
pyautogui.hotkey('ctrl', 'g')  # Atalho para 'Ir Para'
time.sleep(1)
pyautogui.write('anticamara')
pyautogui.press('enter')
time.sleep(2)
pyautogui.hotkey('Ctrl', 'down')
time.sleep(0.5)
pyautogui.press('down')

# Colar os dados copiados e exclui o indice
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('up')
pyautogui.press('down')
pyautogui.hotkey('Shift', 'space')
pyautogui.hotkey('Ctrl', '-')
