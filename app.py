import time
import openpyxl
import pyperclip
import pyautogui

# ENTRAR NA PLANILHA
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# COPIAR INFORMAÇÃO DE UM CAMPO E COLAR NO SEU CORRESPONDENTE
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    time.sleep(3)
    pyautogui.hotkey('tab')
    time.sleep(0.5)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)

    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(3)
    pyautogui.hotkey('tab')

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        pyautogui.hotkey('enter')
        pyautogui.hotkey('enter')
    elif tamanho == 'Médio':
        pyautogui.hotkey('enter')
        pyautogui.hotkey('down')
        pyautogui.hotkey('enter')
    else:
        pyautogui.hotkey('enter')
        pyautogui.hotkey('down')
        pyautogui.hotkey('down')
        pyautogui.hotkey('enter')
    pyautogui.hotkey('tab')

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)

    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(3)
    pyautogui.hotkey('tab')

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.hotkey('Ctrl', 'shft', 'v')
    time.sleep(0.5)
    pyautogui.hotkey('tab')

    pyautogui.hotkey('enter')
    pyautogui.hotkey('enter')
    time.sleep(2)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(2)
# REPETIR ESSES PASSOS PARA OUTROS CAMPOS ATÉ PREENCHER CAMPOS DAQUELA PÁGINA
# CLICAR EM PRÓXIMO
# REPETIR OS MESMOS PASSOS E IR PARA A PROXIMA PÁGIN    A
# REPETIR OS MESMOS PASSOS E FINALIZAR
# CLICAR EM CONCLUIR