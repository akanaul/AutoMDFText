#IMPORTAR BIBLIOTECA E RECURSOS
import pyautogui
import time
import ctypes
import os
import pyperclip
import re
import sys
import tkinter as tk
import threading

#---------------------------------------------------------------
#PARA CANCELAR O CÓDIGO A QUALQUER MOMENTO
print("Para cancelar o código, mova o mouse para o canto superior esquerdo da tela.")
pyautogui.FAILSAFE = True

time.sleep(0.5)

#---------------------------------------------------------------
#PRONT DE ALERTA ANTES DE SEGUIR COM O CÓDIGO    
pyautogui.alert(
    'ANTES DE PROSSEGUIR:\n\n'
    '1. Mantenha 3 abas do Invoisys abertas e o site de averbação logado;\n'
    '2. Deixe o navegador como o primeiro app na barra do Windows;\n'
    '3. Mantenha apenas uma janela do navegador ativa.\n\n'
    'OBS: Para interromper o código, mova o mouse repetidamente para o canto superior direito da tela.'
)

time.sleep(1)

print("Código em execução...")

#---------------------------------------------------------------
# ABRIR NAVEGADOR
pyautogui.hotkey('winleft', '1')
time.sleep(1)
#GAP
for _ in range(2):
    pyautogui.press('esc')
    time.sleep(0.3)

#---------------------------------------------------------------
# ATUALIZAR 3ª PAGINA
pyautogui.hotkey('ctrl', '3')
time.sleep(0.5)
pyautogui.press('f5')
time.sleep(1)

#---------------------------------------------------------------
# IR PARA 1ª ABA DO NAVEGADOR
pyautogui.hotkey('ctrl', '1')
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'f')
time.sleep(1)
pyautogui.write('empresa', interval=0.10)
pyautogui.press('esc')
time.sleep(0.5)

#---------------------------------------------------------------
#CERTIFICAR DE QUE A ABA DO CTE ESTEJA ABERTA
# CONFIGURAÇÃO — ajuste conforme necessário
TARGET_TEXT = "notas emitidas: ct-e"
TEMPO_MAXIMO = 4       # segundos (aumente se seu site demora a carregar)
INTERVALO = 1           # segundos entre tentativas
COPY_ATTEMPTS = 2         # pequenas tentativas rápidas de Ctrl+C por iteração
SHORT_SLEEP = 0.12        # pausa curta entre ações de cópia
PRINT_CLIP_PREVIEW = 400  # quantos chars do clipboard mostrar no log final

def log(msg: str):
    ts = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
    print(f"[{ts}] {msg}")

def wait_for_text(target_text: str,
                  tempo_maximo: float = TEMPO_MAXIMO,
                  intervalo: float = INTERVALO,
                  copy_attempts: int = COPY_ATTEMPTS,
                  short_sleep: float = SHORT_SLEEP):
    inicio = time.monotonic()
    ultimo_conteudo = ""
    target_norm = re.sub(r"\s+", " ", target_text).strip().lower()

    log("Aguardando o formulário abrir...")
    attempt = 0

    while time.monotonic() - inicio < tempo_maximo:
        attempt += 1
        try:
            try:
                pyperclip.copy("")
            except Exception as e:
                log(f"Aviso: não foi possível limpar clipboard: {e}")

            for i in range(copy_attempts):
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(short_sleep)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(short_sleep)

            try:
                conteudo = pyperclip.paste() or ""
            except Exception as e:
                conteudo = ""
                log(f"Aviso: erro ao ler clipboard: {e}")

            ultimo_conteudo = conteudo
            conteudo_norm = re.sub(r"\s+", " ", conteudo).strip().lower()

            log(f"Tentativa {attempt}: comprimento clipboard = {len(conteudo)}")

            if target_norm in conteudo_norm:
                log("Formulário detectado!")
                return True, conteudo

            log(f"Não encontrado. Aguardando {intervalo}s antes da próxima tentativa.")
            time.sleep(intervalo)

        except Exception as e:
            log(f"Erro interno durante a tentativa: {e}")
            time.sleep(intervalo)

    return False, ultimo_conteudo


if __name__ == "__main__":
    encontrado, clipboard = wait_for_text(
        TARGET_TEXT,
        TEMPO_MAXIMO,
        INTERVALO,
        COPY_ATTEMPTS,
        SHORT_SLEEP
    )

    if not encontrado:
        log(f"Formulário não foi detectado dentro de {TEMPO_MAXIMO} segundos.")

        pyautogui.alert(
            title="",
            text=(
                "Formulário não encontrado. O processo será finalizado, tente novamente."
            )
        )

        if clipboard:
            log("Último conteúdo capturado (preview):")
            print(clipboard[:PRINT_CLIP_PREVIEW])

        raise SystemExit(1)

    log("Continuando com o restante da automação...")
    time.sleep(1)

#---------------------------------------------------------------
#PESQUISAR E BAIXAR CTE
pyautogui.hotkey('ctrl', 'f')
time.sleep(2)

pyautogui.write('serie final', interval=0.12)
pyautogui.press('esc')
time.sleep(0.5)

for _ in range(2):
    pyautogui.press('tab')
    time.sleep(1)

#---------------------------------------------------------------
#PROMPT DE COMANDO PARA DIGITAR A DT
codigo = pyautogui.prompt(
    text='Digite o número do DT:',
    title='DT'
)

if not codigo:
    pyautogui.alert('Nenhum código informado. O script foi pausado.')
    pyautogui.FAILSAFE = True
    exit()
pyautogui.write(codigo.upper(), interval=0.1)
pyautogui.press('enter')
time.sleep(0.3)
pyautogui.press('enter')
time.sleep(0.5)
#---------------------------------------------------------------

#---------------------------------------------------------------

#ALERTA - ORIENTAÇÕES ANTES DE SEGUIR COM O CÓDIGO
pyautogui.alert(
    'Baixe o arquivo XML antes de prosseguir.'
)
time.sleep(1)
#---------------------------------------------------------------


#DESATIVAR CAPS LOOK
VK_CAPITAL = 0x14  # código da tecla Caps Lock

# Obtém o estado atual do Caps Lock
caps_state = ctypes.windll.user32.GetKeyState(VK_CAPITAL)

# Se estiver ativo, desliga
if caps_state & 1:
    # Pressiona Caps Lock
    ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 0, 0)
    # Solta Caps Lock
    ctypes.windll.user32.keybd_event(VK_CAPITAL, 0, 2, 0)


#-----------------------------------------------------
#DADADOS DO MDF-E
# IR PARA 3ª PAGINA - MDF
pyautogui.hotkey('ctrl', '3')
time.sleep(0.5)

# ABRIR DADOS DO MDF-E
pyautogui.hotkey('ctrl', 'f')
pyautogui.write('EMITIR NOTA', interval=0.10)
pyautogui.press('esc')
pyautogui.press('enter')
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'f')
pyautogui.write('MDF-E', interval=0.10)
pyautogui.press('esc')
pyautogui.press('enter')
time.sleep(0.8)
#---------------------------------------------------------------

# AGUARDAR FORMULÁRIO ABRIR E DETECTAR "Emissor MDF-e"
# CONFIGURAÇÃO — ajuste conforme necessário
TARGET_TEXT = "Emissor MDF-e"
TEMPO_MAXIMO = 15.0       # segundos (aumente se seu site demora a carregar)
INTERVALO = 3.0           # segundos entre tentativas
COPY_ATTEMPTS = 3         # pequenas tentativas rápidas de Ctrl+C por iteração
SHORT_SLEEP = 0.12        # pausa curta entre ações de cópia
PRINT_CLIP_PREVIEW = 400  # quantos chars do clipboard mostrar no log final

def log(msg: str):
    ts = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
    print(f"[{ts}] {msg}")

def wait_for_text(target_text: str,
                  tempo_maximo: float = TEMPO_MAXIMO,
                  intervalo: float = INTERVALO,
                  copy_attempts: int = COPY_ATTEMPTS,
                  short_sleep: float = SHORT_SLEEP):
    inicio = time.monotonic()
    ultimo_conteudo = ""
    target_norm = re.sub(r"\s+", " ", target_text).strip().lower()

    log("Aguardando o formulário abrir...")
    attempt = 0

    while time.monotonic() - inicio < tempo_maximo:
        attempt += 1
        try:
            try:
                pyperclip.copy("")
            except Exception as e:
                log(f"Aviso: não foi possível limpar clipboard: {e}")

            for i in range(copy_attempts):
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(short_sleep)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(short_sleep)

            try:
                conteudo = pyperclip.paste() or ""
            except Exception as e:
                conteudo = ""
                log(f"Aviso: erro ao ler clipboard: {e}")

            ultimo_conteudo = conteudo
            conteudo_norm = re.sub(r"\s+", " ", conteudo).strip().lower()

            log(f"Tentativa {attempt}: comprimento clipboard = {len(conteudo)}")

            if target_norm in conteudo_norm:
                log("Formulário detectado!")
                return True, conteudo

            log(f"Não encontrado. Aguardando {intervalo}s antes da próxima tentativa.")
            time.sleep(intervalo)

        except Exception as e:
            log(f"Erro interno durante a tentativa: {e}")
            time.sleep(intervalo)

    return False, ultimo_conteudo


if __name__ == "__main__":
    encontrado, clipboard = wait_for_text(
        TARGET_TEXT,
        TEMPO_MAXIMO,
        INTERVALO,
        COPY_ATTEMPTS,
        SHORT_SLEEP
    )

    if not encontrado:
        log(f"Formulário não foi detectado dentro de {TEMPO_MAXIMO} segundos.")

        pyautogui.alert(
            title="Formulário não encontrado",
            text=(
                "Formulário não encontrado. O processo será finalizado, tente novamente."
            )
        )

        if clipboard:
            log("Último conteúdo capturado (preview):")
            print(clipboard[:PRINT_CLIP_PREVIEW])

        raise SystemExit(1)

    log("Continuando com o restante da automação...")
    time.sleep(0.8)
#---------------------------------------------------------------

## DADOS DO MDFE: PRESTADPR DE SERVIÇO
pyautogui.hotkey('ctrl', 'f')
pyautogui.write('SELECIONE...', interval=0.12)
pyautogui.press('esc')
time.sleep(0.2)
pyautogui.press('enter')
time.sleep(0.2)
pyautogui.write('PRESTADOR', interval=0.1)
time.sleep(0.5)
pyautogui.press('enter')
time.sleep(0.2)
pyautogui.press('tab')
time.sleep(0.2)
pyautogui.press('space')
time.sleep(0.2)
#---------------------------------------------------------------

## DADOS DO MDFE: EMITENTE
pyautogui.write('0315-60', interval=0.15)
pyautogui.press('enter')
time.sleep(0.5)

for _ in range(7):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.press('space')
time.sleep(0.2)
#---------------------------------------------------------------

## DADOS DO MDFE: UF CARREGAMENTO E DESCARREGAMENTO
pyautogui.write('SP', interval=0.20)
pyautogui.press('enter')
pyautogui.press('tab')
time.sleep(0.5)
pyautogui.press('space')
time.sleep(0.2)
pyautogui.write('SP', interval=0.20)
pyautogui.press('enter')
time.sleep(0.5)
#---------------------------------------------------------------

## DADOS DO MDFE: MUNICIPIO DE CARREGAMENTO
pyautogui.press('tab')
time.sleep(0.1)
pyautogui.write('ITU'.upper(), interval=0.15)
time.sleep(0.3)

for _ in range(4):
    pyautogui.press('down')
    time.sleep(0.1)
for _ in range(3):
    pyautogui.press('up')
    time.sleep(0.1)
pyautogui.press('enter')
time.sleep(0.2)
#---------------------------------------------------------------

from pathlib import Path

## UPLOAD DO ARQUIVO XML
for _ in range(2):
    pyautogui.press('tab')
pyautogui.press('space')
time.sleep(2)

downloads_path = Path.home() / "Downloads"
list_of_files = list(downloads_path.glob('*.xml'))

if not list_of_files:
    pyautogui.alert('Nenhum arquivo XML encontrado na pasta Downloads!')
    sys.exit()

latest_file = max(list_of_files, key=lambda f: f.stat().st_ctime)

pyautogui.write(str(latest_file), interval=0.03)
time.sleep(0.5)
pyautogui.press('enter')

time.sleep(0.5)
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.3)
pyautogui.press('enter')
time.sleep(2)

#---------------------------------------------------------------

## DADOS DO MDFE: UNIDADE DE MEDIDA
for _ in range(5):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.press('space')
time.sleep(0.1)
pyautogui.write('1', interval=0.1)
pyautogui.press('enter')
#---------------------------------------------------------------

## DADOS DO MDFE: TIPO DE CARGA
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.press('space')
pyautogui.press('tab')
pyautogui.press('space')
time.sleep(0.1)
pyautogui.write('05', interval=0.1)
pyautogui.press('enter')

## DADOS DO MDFE: DESCRIÇÃO DO PRODUTO
pyautogui.press('tab')
pyautogui.write('PA/PALLET', interval=0.1)
#---------------------------------------------------------------

## DADOS DO MDFE: CÓDIGO NCM
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.1)
opcao = pyautogui.confirm(
    text='Selecione o código NCM ou escolha "Outro código" para digitar manualmente:',
    title='Escolha de NCM',
    buttons=['19041000', '19059090', '20052000', 'Outro código', 'Cancelar']
)

if opcao == 'Cancelar':
    pyautogui.alert('Nenhum código NCM selecionado. O script foi pausado.')
    pyautogui.FAILSAFE = True
    exit()
elif opcao == 'Outro código':
    codigo = pyautogui.prompt('Digite o código NCM:')
    if not codigo:
        pyautogui.alert('Nenhum código NCM digitado. O script foi pausado.')
        pyautogui.FAILSAFE = True
        exit()
else:
    codigo = opcao
pyautogui.write(codigo.upper(), interval=0.1)
pyautogui.press('enter')
#---------------------------------------------------------------

## DADOS DO MDFE: CEP ORIGEM
pyautogui.press('tab')
pyautogui.press('space')
pyautogui.press('tab')
pyautogui.write('13300340', interval=0.1)

## DADOS DO MDFE: CEP DESTINO
for _ in range(3):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.write('13315000', interval=0.1)
time.sleep(1)


#-----------------------------------------------------#-----------------------------------------------------------------#
#-----------------------------------------------------#-----------------------------------------------------------------#
#						MODAL RODOVIÁRIO							#


# ABRIR DADOS DO MODAL RODOVIÁRIO
time.sleep(1)
pyautogui.hotkey('ctrl', 'f')
pyautogui.write('modal rodo', interval=0.10)
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.2)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('esc')
pyautogui.press('enter')
time.sleep(1)

#---------------------------------------------------------------


## RNTRC
pyautogui.hotkey('ctrl', 'f')
pyautogui.write('RNTRC', interval=0.10)
pyautogui.press('esc')
pyautogui.press('tab')
pyautogui.write('45501846', interval=0.10)
#---------------------------------------------------------------

## NOME DO CONTRATANTE
for _ in range(6):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.press('space')
pyautogui.press('tab')
pyautogui.write('PEPSICO ITU', interval=0.20)
time.sleep(0.1)

## CNPJ DO CONTRATATANTE
pyautogui.press('tab')
pyautogui.write('02957518000224', interval=0.12)
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.press('enter')
time.sleep(1)


#-----------------------------------------------------#-----------------------------------------------------------------#
#-----------------------------------------------------#-----------------------------------------------------------------#
#						INFORMAÇÕES OPCIONAIS							#

# ABRIR DADOS DE INFORMAÇÕES OPCIONAIS
pyautogui.hotkey('ctrl', 'f')
pyautogui.write('OPCIONAIS', interval=0.10)
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('esc')
time.sleep(0.5)
pyautogui.press('enter')
time.sleep(1)



## INFORMAÇÕES ADICIONAIS
pyautogui.hotkey('ctrl', 'f')
time.sleep(0.5)
pyautogui.write('ADICIONAIS', interval=0.10)
time.sleep(0.3)
pyautogui.press('esc')
time.sleep(0.3)
pyautogui.press('enter')
time.sleep(0.5)

pyautogui.hotkey('ctrl', 'f')
time.sleep(0.5)
pyautogui.write('CONTRIBUINTE', interval=0.10)
time.sleep(0.3)
pyautogui.press('esc')
time.sleep(0.3)
pyautogui.press('tab')
time.sleep(0.3)
pyautogui.press('tab')
time.sleep(0.3)
pyautogui.press('tab')
time.sleep(0.3)
pyautogui.write('04898488000177', interval=0.10)
pyautogui.press('tab')
time.sleep(0.3)
pyautogui.press('enter')

for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.2)
pyautogui.press('space')
time.sleep(0.2)
pyautogui.press('tab')
pyautogui.press('space')
time.sleep(0.2)
pyautogui.write('CONTRA', interval=0.10)
pyautogui.press('enter')
time.sleep(0.3)
pyautogui.press('tab')
pyautogui.write('02957518000224', interval=0.12)
pyautogui.press('tab')
pyautogui.write('SEGUROS SURA SA', interval=0.10)
pyautogui.press('tab')
pyautogui.write('33065699000127', interval=0.12)
pyautogui.press('tab')
pyautogui.write('5400035882', interval=0.10)
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(0.3)

for _ in range(4):
    pyautogui.press('tab')
    time.sleep(0.3)
pyautogui.press('space')
time.sleep(0.2)

pyautogui.press('tab')
time.sleep(0.2)
pyautogui.write('PEPSICO DO BRASIL', interval=0.12)
time.sleep(0.2)

pyautogui.press('tab')
time.sleep(0.2)
pyautogui.write('02957518000224', interval=0.12)
time.sleep(0.2)

for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.3)

pyautogui.write('1314.27', interval=0.12)
time.sleep(0.2)

pyautogui.press('tab')
time.sleep(0.2)
pyautogui.press('space')
time.sleep(0.2)
pyautogui.write('1', interval=0.12)
time.sleep(0.2)

pyautogui.press('enter')
time.sleep(0.3)
pyautogui.press('tab')
time.sleep(0.2)
pyautogui.write('237', interval=0.12)
time.sleep(0.2)

pyautogui.press('tab')
time.sleep(0.2)
pyautogui.write('2372/8', interval=0.12)
time.sleep(0.2)

for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.3)

pyautogui.press('enter')
time.sleep(0.3)

pyautogui.press('tab')
time.sleep(0.2)
pyautogui.press('enter')

pyautogui.hotkey('ctrl', 'f')
pyautogui.write('SELECIONE...', interval=0.10)
pyautogui.press('esc')
pyautogui.press('enter')
pyautogui.write('FRETE', interval=0.10)
pyautogui.press('enter')
time.sleep(0.2)
pyautogui.press('tab')
time.sleep(0.2)
pyautogui.write('1314.27', interval=0.12)
time.sleep(0.2)
pyautogui.press('tab')
time.sleep(0.2)
pyautogui.write('FRETE', interval=0.12)
pyautogui.press('tab')
time.sleep(0.2)
pyautogui.press('enter')
time.sleep(0.2)

pyautogui.hotkey('ctrl', 'f')
pyautogui.write('SELECIONE...', interval=0.10)
pyautogui.press('esc')
time.sleep(0.10)

for _ in range(5):
    pyautogui.press('tab')
    time.sleep(0.05)
pyautogui.write('1', interval=0.10)

for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.05)
pyautogui.press('space')
time.sleep(0.10)

for _ in range(7):
    pyautogui.press('tab')
    time.sleep(0.05)
pyautogui.press('space')
time.sleep(0.05)
pyautogui.press('space')
time.sleep(0.05)
pyautogui.press('tab')
time.sleep(0.05)
pyautogui.press('enter')
time.sleep(0.05)
pyautogui.press('tab')
pyautogui.write('1314.27', interval=0.10)
time.sleep(0.05)
pyautogui.press('tab')
time.sleep(0.05)
pyautogui.press('enter')
time.sleep(1)


#-----------------------------------------------------#-----------------------------------------------------------------#
#-----------------------------------------------------#-----------------------------------------------------------------#
#						AVERBAÇÃO								#

	
pyautogui.hotkey('ctrl', 'shift', 'a')
time.sleep(0.5)
pyautogui.write('ATM', interval=0.10)
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.2)
pyautogui.press('enter')
time.sleep(1)

pyautogui.hotkey('ctrl', 'f')
time.sleep(0.5)
pyautogui.write('OK', interval=0.10)
time.sleep(0.5)
pyautogui.press('esc')
time.sleep(0.2)
pyautogui.press('enter')
time.sleep(0.5)

pyautogui.hotkey('ctrl', 'f')
time.sleep(0.5)
pyautogui.write('XML', interval=0.10)
time.sleep(0.5)
pyautogui.press('esc')
time.sleep(0.2)
pyautogui.press('enter')
time.sleep(0.5)

pyautogui.hotkey('ctrl', 'f')
time.sleep(0.5)
pyautogui.write('ENVIAR', interval=0.10)
time.sleep(0.5)
pyautogui.press('esc')
time.sleep(0.2)
pyautogui.press('enter')
time.sleep(1)


#---------------------------------------------------------------

## UPLOAD DO ARQUIVO XML
downloads_path = Path.home() / "Downloads"
list_of_files = list(downloads_path.glob('*'))
if not list_of_files:
    pyautogui.alert('A pasta Downloads está vazia!')
else:
    latest_file = max(list_of_files, key=os.path.getctime)
    pyautogui.write(str(latest_file), interval=0.07)
    time.sleep(0.3)
    pyautogui.press('enter')
time.sleep(2)
#---------------------------------------------------------------
# Seleciona todo o texto da janela e copia
pyautogui.hotkey('ctrl', 'a')
time.sleep(0.2)
pyautogui.hotkey('ctrl', 'c')
time.sleep(0.2)

# Pega o texto da área de transferência
texto = pyperclip.paste()

# Extrai somente os números da linha "Número de Averbação"
match = re.search(r'Número de Averbação:\s*([\d]+)', texto)
if match:
    numero_averbacao = match.group(1)
    
    # Coloca somente os números da averbação de volta na área de transferência
    pyperclip.copy(numero_averbacao)
    print("Número de Averbação copiado:", numero_averbacao)
else:
    print("Número de Averbação não encontrado")

time.sleep(0.5)
pyautogui.hotkey('alt', 'tab')
time.sleep(1)

pyautogui.hotkey('ctrl', 'f')
time.sleep(0.5)
pyautogui.write('DETALHES', interval=0.10)
time.sleep(0.3)
pyautogui.press('esc')
pyautogui.press('enter')
time.sleep(0.5)
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.5)
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
time.sleep(0.5)
pyautogui.press('enter')
time.sleep(1)


# ---------- CÓDIGO ITU PARTE 5: COLETAR DT e CTE -------

pyautogui.hotkey('ctrl', 'shift', 'a')
time.sleep(0.5)
pyautogui.write('INVOISYS', interval=0.10)
for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.5)
pyautogui.press('enter')
time.sleep(1)

pyautogui.hotkey('ctrl', 'f')
pyautogui.write('e final', interval=0.20)
pyautogui.press('esc')
time.sleep(0.2)

for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.1)

time.sleep(0.5)

pyautogui.hotkey('ctrl', 'a')
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'c')
time.sleep(0.5)
pyautogui.hotkey('alt', 'tab')
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'f')
time.sleep(0.5)
pyautogui.write('CONTRIBUINTE', interval=0.10)
time.sleep(0.3)
pyautogui.press('esc')
time.sleep(0.5)
pyautogui.press('tab')
time.sleep(0.5)
pyautogui.write('DT: ', interval=0.10)
time.sleep(0.3)
pyautogui.hotkey('ctrl', 'v')
time.sleep(0.5)
pyautogui.write(' CTE: ', interval=0.10)
time.sleep(0.5)
pyautogui.hotkey('alt', 'tab')
time.sleep(0.5)
########################################################

for _ in range(2):
    pyautogui.press('tab')
    time.sleep(0.3)
# Seleciona tudo e copia
pyautogui.hotkey('ctrl', 'a')
time.sleep(0.3)
pyautogui.hotkey('ctrl', 'c')
time.sleep(1)

# Pega o conteúdo copiado
conteudo = pyperclip.paste()

# Divide o conteúdo em linhas
linhas = conteudo.splitlines()

# Procura a linha que contém "100 - Autorizado o uso do CT-e.N"
numero_cte = None
for linha in linhas:
    if "100 - Autorizado o uso do CT-e.N" in linha:
        # Tenta extrair o número de 6 dígitos após essa frase
        match = re.search(r"100\s*-\s*Autorizado o uso do CT-e\.N[^\d]*(\d{6})", linha)
        if match:
            numero_cte = match.group(1)
            break

# Exibe e copia o resultado
if numero_cte:
    print("Número do CT-e encontrado:", numero_cte)
    pyperclip.copy(numero_cte)
    print("Número do CT-e copiado para a área de transferência.")
else:
    print("Não foi possível localizar o número do CT-e.")
########################################################

pyautogui.hotkey('alt', 'tab')
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'v')
time.sleep(0.5)
pyautogui.write(' NF: ', interval=0.10)
time.sleep(0.5)


# ---------- FINALIZAÇÃO ----------
pyautogui.alert('Sucesso! Inclua a NF e os dados do motorista')
