'''
Vídeo que deu essa ideia:
    https://www.youtube.com/watch?v=UtkPIpov6h8&t=1336s&ab_channel=DevAprender%7CJhonatandeSouza
Site usado como "sistema web de cadastro dos produtos":
    https://cadastro-produtos-devaprender.netlify.app/index.html

This code, it's comments and the Excel's workbook used are in Brazilian Portuguese. In future, I'll "translate" them to english.

'''

import pyautogui
import openpyxl
import pyperclip
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(852,187, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.press('tab')
    pyautogui.press('enter')
    sleep(5)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyperclip.copy(tamanho)
    pyautogui.press('tab')
    pyautogui.press('enter')
    if tamanho == "Pequeno":
        pyautogui.press('enter')
    elif tamanho == "Médio":
        pyautogui.press('down')
        pyautogui.press('enter')
    else:
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('enter')

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.press('tab')
    pyautogui.press('enter')
    sleep(5)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    localizacao_no_armazem = linha[16].value
    pyperclip.copy(localizacao_no_armazem)
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('enter')
    sleep(5)
    pyautogui.press('enter')

    pyautogui.press('tab')
    pyautogui.press('enter')
    sleep(5)

print('Cadastro de produtos da planilha concluídos com sucesso.')