import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1571,882, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1495,990, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1500,1094, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1444,1188, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1607,1275, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1442,1357, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1461,1424, duratuion=1)
    sleep(5)

    preco = linha[6].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1451,903, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1433,990, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1439,1077, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1515,1156, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(1424,1243,duatuion=1)

    if tamanho == 'Pequeno':
        pyautogui.click(1462,1278,duatuion=1)
    elif tamanho == 'Médio':
        pyautogui.click(1452,1296,duatuion=1)
    else:
        pyautogui.click(1462,1315,duatuion=1)

    material = linha[11].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1439,1337, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1477,1397,duatuion=1)
    sleep(5)

    fabricante = linha[12].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1461,926, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1436,1012, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1439,1091, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1441,1223, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1436,1318, duratuion=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1476,1388, duratuion=1)
    pyautogui.click(2255,720, duratuion=1)
    pyautogui.click(2120,1152, duratuion=1)