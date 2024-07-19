import pyautogui
import subprocess
import time
import os
from datetime import datetime, timedelta
import pygetwindow as gw
from pyscreeze import ImageNotFoundException

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

#retorna data com /
def data():
    ontem = datetime.now() - timedelta(days=1)
    return ontem.strftime('%d/%m/%Y')

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

# --------------------------------------------------------------------------------

# Puxar relatório da anticâmara
clicar_imagem('anti_camara.png', duplo_clique=True)
time.sleep(2)
inserir_data()
clicar_imagem('temperatura_ambiente.png')
clicar_imagem('dispositivo.png')
clicar_imagem('adicionar.png')
clicar_imagem('gerar.png')
time.sleep(10)

# Verifica se a mensagem de erro apareceu na tela
imagem_erro = caminho_imagem('erro.png')
try:
    erro_localizado = list(pyautogui.locateAllOnScreen(imagem_erro))
except ImageNotFoundException:
    erro_localizado = []

if erro_localizado:
    pyautogui.hotkey('alt', 'f4')
    print('Erro relatório Anticamara')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)
else:
    clicar_imagem('ferramentas.png')
    clicar_imagem('relatoriotexto.png')
    clicar_imagem('anticamaratexto.png')
    time.sleep(10)

    # Selecionar e copiar o texto gerado
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    # Caminho para o arquivo da planilha de controles
    caminho_planilha = r'\\Srv-araovos\d\Meus documentos\INFO\INFO\Controles\sitrad controladores.xlsx'

    # Abrir Excel com a planilha específica
    subprocess.Popen([r'C:\Program Files\Microsoft Office\Office16\EXCEL.EXE', caminho_planilha])
    time.sleep(20)

    # Navegar para a aba "ANTICAMARA (13)"
    pyautogui.hotkey('ctrl', 'g')  # Atalho para 'Ir Para'
    time.sleep(1)
    pyautogui.write('anticamara')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('Ctrl', 'down')
    time.sleep(0.5)
    pyautogui.press('down')

    # Colar os dados copiados e excluir o índice
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('up')
    time.sleep(0.5)
    pyautogui.press('down')
    time.sleep(0.5)
    pyautogui.hotkey('Shift', 'space')
    time.sleep(0.5)
    pyautogui.hotkey('Ctrl', '-')
    time.sleep(2)

    # Voltar para a janela do Sitrad e fechar as abas para o próximo relatório
    pyautogui.hotkey('alt', 'tab')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)

# --------------------------------------------------------------------------------

# Puxar relatório da Camara fria
clicar_imagem('camara_fria.png', duplo_clique=True)
time.sleep(2)
clicar_imagem('temperatura_ambiente.png')
clicar_imagem('temperatura_evaporador.png')
clicar_imagem('dispositivos.png')
clicar_imagem('adicionar.png')
clicar_imagem('gerar.png')
time.sleep(8)

# Verifica se a mensagem de erro apareceu na tela
imagem_erro = caminho_imagem('erro.png')
try:
    erro_localizado = list(pyautogui.locateAllOnScreen(imagem_erro))
except ImageNotFoundException:
    erro_localizado = []

if erro_localizado:
    pyautogui.hotkey('alt', 'f4')
    print('Erro relatório Camara fria')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)
else:
    print('Camara fria voltou, por favor ajuste o código')

# --------------------------------------------------------------------------------

# Puxar relatório do Pasteurizador
clicar_imagem('pasteurizador_1.png', duplo_clique=True)
time.sleep(2)
clicar_imagem('temperatura_ambiente.png')
clicar_imagem('dispositivo.png')
clicar_imagem('adicionar.png')
clicar_imagem('gerar.png')
time.sleep(10)

# Verifica se a mensagem de erro apareceu na tela
imagem_erro = caminho_imagem('erro.png')
try:
    erro_localizado = list(pyautogui.locateAllOnScreen(imagem_erro))
except ImageNotFoundException:
    erro_localizado = []

if erro_localizado:
    pyautogui.hotkey('alt', 'f4')
    print('Erro relatório Pasteurizador 1')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)
else:
    clicar_imagem('ferramentas.png')
    clicar_imagem('relatoriotexto.png')
    clicar_imagem('pasteurizador_1texto.png')
    time.sleep(10)

    # Selecionar e copiar o texto gerado
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    # Voltar para a janela do Excel
    janela_excel = gw.getWindowsWithTitle("sitrad controladores.xlsx")[0]
    janela_excel.activate()

    # Navegar para a aba "PASTEURIZADOR 1"
    pyautogui.hotkey('ctrl', 'g')  # Atalho para 'Ir Para'
    time.sleep(1)
    pyautogui.write('pasteurizador1')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('Ctrl', 'down')
    time.sleep(0.5)
    pyautogui.press('down')

    # Colar os dados copiados e excluir o índice
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('up')
    time.sleep(0.5)
    pyautogui.press('down')
    time.sleep(0.5)
    pyautogui.hotkey('Shift', 'space')
    time.sleep(0.5)
    pyautogui.hotkey('Ctrl', '-')
    time.sleep(2)

    # Voltar para a janela do Sitrad e fechar as abas para o próximo relatório
    pyautogui.hotkey('alt', 'tab')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)

# --------------------------------------------------------------------------------

# Puxar relatório do Tunel de congelamento
clicar_imagem('tunel_de_congelamento.png', duplo_clique=True)
time.sleep(2)
clicar_imagem('temperatura_ambiente.png')
clicar_imagem('temperatura_evaporador.png')
clicar_imagem('temperatura_sensor_3.png')
clicar_imagem('dispositivos.png')
clicar_imagem('adicionar.png')
clicar_imagem('gerar.png')
time.sleep(10)

# Verifica se a mensagem de erro apareceu na tela
imagem_erro = caminho_imagem('erro.png')
try:
    erro_localizado = list(pyautogui.locateAllOnScreen(imagem_erro))
except ImageNotFoundException:
    erro_localizado = []

if erro_localizado:
    pyautogui.hotkey('alt', 'f4')
    print('Erro relatório tunel de congelamento')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)
else:
    clicar_imagem('ferramentas.png')
    clicar_imagem('relatoriotexto.png')
    clicar_imagem('tunel_congelamentotexto.png')
    time.sleep(10)

    # Selecionar e copiar o texto gerado
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    # Voltar para a janela do Excel
    janela_excel = gw.getWindowsWithTitle("sitrad controladores.xlsx")[0]
    janela_excel.activate()

    # Navegar para a aba "PASTEURIZADOR 1"
    pyautogui.hotkey('ctrl', 'g')  # Atalho para 'Ir Para'
    time.sleep(1)
    pyautogui.write('tunelcongelamento')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('Ctrl', 'down')
    time.sleep(0.5)
    pyautogui.press('down')

    # Colar os dados copiados e excluir o índice
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('up')
    time.sleep(0.5)
    pyautogui.press('down')
    time.sleep(0.5)
    pyautogui.hotkey('Shift', 'space')
    time.sleep(0.5)
    pyautogui.hotkey('Ctrl', '-')
    time.sleep(2)

    # Voltar para a janela do Sitrad e fechar as abas para o próximo relatório
    pyautogui.hotkey('alt', 'tab')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)

# -------------------------------------------------------------------------------

# Puxar relatório do Tunel de congelado
clicar_imagem('tunel_congelado1.png')
time.sleep(1)
clicar_imagem('tunel_congelado2.png', duplo_clique=True)
time.sleep(2)
clicar_imagem('temperatura_sensor_1.png')
clicar_imagem('temperatura_sensor_2.png')
clicar_imagem('temperatura_sensor_3.png')
clicar_imagem('temperatura_diferencial.png')
clicar_imagem('temperatura_media.png')
clicar_imagem('adicionar.png')
clicar_imagem('gerar.png')
time.sleep(10)

# Verifica se a mensagem de erro apareceu na tela
imagem_erro = caminho_imagem('erro.png')
try:
    erro_localizado = list(pyautogui.locateAllOnScreen(imagem_erro))
except ImageNotFoundException:
    erro_localizado = []

if erro_localizado:
    pyautogui.hotkey('alt', 'f4')
    print('Erro relatório tunel de congelado')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)
else:
    clicar_imagem('ferramentas.png')
    clicar_imagem('relatoriotexto.png')
    clicar_imagem('tunel_congeladotexto.png')
    time.sleep(10)

    # Selecionar e copiar o texto gerado
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    # Voltar para a janela do Excel
    janela_excel = gw.getWindowsWithTitle("sitrad controladores.xlsx")[0]
    janela_excel.activate()

    # Navegar para a aba "PASTEURIZADOR 1"
    pyautogui.hotkey('ctrl', 'g')  # Atalho para 'Ir Para'
    time.sleep(1)
    pyautogui.write('tunelcongelado')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('Ctrl', 'down')
    time.sleep(0.5)
    pyautogui.press('down')

    # Colar os dados copiados e excluir o índice
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('up')
    time.sleep(0.5)
    pyautogui.press('down')
    time.sleep(0.5)
    pyautogui.hotkey('Shift', 'space')
    time.sleep(0.5)
    pyautogui.hotkey('Ctrl', '-')
    time.sleep(2)

    # Voltar para a janela do Sitrad e fechar as abas para o próximo relatório
    pyautogui.hotkey('alt', 'tab')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)

# ------------------------------------------------------------------------------

# Puxar relatório do Tunel de resfriamento
clicar_imagem('tunel_resfriamento.png', duplo_clique=True)
time.sleep(2)
clicar_imagem('temperatura_ambiente.png')
clicar_imagem('temperatura_evaporador.png')
clicar_imagem('dispositivos.png')
clicar_imagem('adicionar.png')
clicar_imagem('gerar.png')
time.sleep(10)

# Verifica se a mensagem de erro apareceu na tela
imagem_erro = caminho_imagem('erro.png')
try:
    erro_localizado = list(pyautogui.locateAllOnScreen(imagem_erro))
except ImageNotFoundException:
    erro_localizado = []

if erro_localizado:
    pyautogui.hotkey('alt', 'f4')
    print('Erro relatório tunel de resfriamento')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)
else:
    clicar_imagem('ferramentas.png')
    clicar_imagem('relatoriotexto.png')
    clicar_imagem('tunel_resfriamentotexto.png')
    time.sleep(10)

    # Selecionar e copiar o texto gerado
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

    # Voltar para a janela do Excel
    janela_excel = gw.getWindowsWithTitle("sitrad controladores.xlsx")[0]
    janela_excel.activate()

    # Navegar para a aba "PASTEURIZADOR 1"
    pyautogui.hotkey('ctrl', 'g')  # Atalho para 'Ir Para'
    time.sleep(1)
    pyautogui.write('tunelresfriado')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('Ctrl', 'down')
    time.sleep(0.5)
    pyautogui.press('down')

    # Colar os dados copiados e excluir o índice
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('up')
    time.sleep(0.5)
    pyautogui.press('down')
    time.sleep(0.5)
    pyautogui.hotkey('Shift', 'space')
    time.sleep(0.5)
    pyautogui.hotkey('Ctrl', '-')
    time.sleep(2)

# --------------------------------------------------------------------------------------

# salva a planilha após o ultima relatorio ser lançado
pyautogui.hotkey('ctrl', 'b')
time.sleep(6)

# --------------------------------------------------------------------------------------


caminho_planilha = r'\\Srv-araovos\d\Meus documentos\INFO\INFO\Controles\Sitrad Tmp.xlsx'
subprocess.Popen([r'C:\Program Files\Microsoft Office\Office16\EXCEL.EXE', caminho_planilha])
time.sleep(5)

pyautogui.hotkey('ctrl', 'g')
time.sleep(1)
pyautogui.write('A')
pyautogui.press('enter')
time.sleep(1)
pyautogui.hotkey('ctrl', 'l')
time.sleep(2)
pyautogui.write(data())
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.hotkey('alt', 'f4')
time.sleep(1)
for _ in range(5):
    pyautogui.press('right')
    time.sleep(0.1)

# Caminho para o executável do AutoHotKey
ahk_executable = r'C:\Program Files\AutoHotkey\AutoHotkey.exe'
# Caminho para o script AutoHotKey
ahk_script = r'C:\Users\Terminal17\Documents\GitHub\AutomacaoSitrad\shift_select.ahk'

# Executa o script AutoHotKey para pressionar Shift + setas
subprocess.Popen([ahk_executable, ahk_script])
time.sleep(2)

# Copiar e colar a seleção
pyautogui.hotkey('ctrl', 'c')
time.sleep(1)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)

# Colar com somente valores e salvar a planilha
pyautogui.press('ctrl')
time.sleep(0.5)
pyautogui.press('e')
pyautogui.hotkey('ctrl', 'b')
time.sleep(6)

# Fechar o Excel
pyautogui.hotkey('alt', 'f4')
time.sleep(1)
pyautogui.press('enter')

# fecha a planilha de relatorios
pyautogui.hotkey('alt', 'f4')
time.sleep(2)

# Voltar para a janela do Sitrad e fechar as abas para o próximo relatório
for _ in range(3):
    pyautogui.hotkey('alt', 'f4')
    time.sleep(4)
time.sleep(2)