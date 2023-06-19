import pyautogui
import subprocess
from time import sleep
import pandas as pd
import pyperclip


def encontrar_imagem(imagem):
    while not pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9):
        sleep(1)
    encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9)
    return encontrou


def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')


def clicar_direita(posicao):
    return posicao[0] + posicao[2], posicao[1] + posicao[3] / 2


pyautogui.FAILSAFE = True

# abrindo o programa
subprocess.Popen([r'C:\Program Files\Fakturama2\Fakturama.exe'])

# esperando reconhecer o programa
encontrou = encontrar_imagem('logo.png')

# reconhecer a tabela de produtos
tabela_produtos = pd.read_excel('Produtos.xlsx')
for linha in tabela_produtos.index:
    nome = tabela_produtos.loc[linha, 'Nome']
    id = tabela_produtos.loc[linha, 'ID']
    categoria = tabela_produtos.loc[linha, 'Categoria']
    gtin = tabela_produtos.loc[linha, 'GTIN']
    supplier = tabela_produtos.loc[linha, 'Supplier']
    descricao = tabela_produtos.loc[linha, 'Descrição']
    imagem = tabela_produtos.loc[linha, 'Imagem']
    preco = tabela_produtos.loc[linha, 'Preço']
    custo = tabela_produtos.loc[linha, 'Custo']
    estoque = tabela_produtos.loc[linha, 'Estoque']

    # clicando em novo
    encontrou = encontrar_imagem('new.png')
    pyautogui.click(pyautogui.center(encontrou))

    # clicando em novo projeto
    encontrou = encontrar_imagem('projeto.png')
    pyautogui.click(pyautogui.center(encontrou))

    # clicar em numero do item
    encontrou = encontrar_imagem('item.png')
    pyautogui.click(clicar_direita(encontrou))

    # escrever o numero
    escrever_texto(str(id))

    # clicar em nome
    encontrou = encontrar_imagem('nome.png')
    pyautogui.click(clicar_direita(encontrou))

    # escrever o nome
    escrever_texto(str(nome))

    # clicar em categoria
    encontrou = encontrar_imagem('categoria.png')
    pyautogui.click(clicar_direita(encontrou))

    # escrever a categoria
    escrever_texto(str(categoria))

    # clicar em gtin
    encontrou = encontrar_imagem('gtin.png')
    pyautogui.click(clicar_direita(encontrou))

    # escrever o gtin
    escrever_texto(str(gtin))

    # clicar em supplier
    encontrou = encontrar_imagem('supplier.png')
    pyautogui.click(clicar_direita(encontrou))

    # escrever o supplier
    escrever_texto(str(supplier))

    # clicar em descrição
    pyautogui.moveTo(434, 327)
    pyautogui.click()

    # escrever a descrição
    escrever_texto(str(descricao))

    # clicar em preço
    encontrou = encontrar_imagem('preco.png')
    pyautogui.click(clicar_direita(encontrou))
    preco_texto = f'{preco:.2f}'.replace('.', ',')
    # escrever o preço
    escrever_texto(str(preco_texto))

    # clicar em custo
    encontrou = encontrar_imagem('custo.png')
    pyautogui.click(clicar_direita(encontrou))
    custo_texto = f'{custo:.2f}'.replace('.', ',')
    # escrever o custo
    escrever_texto(str(custo_texto))

    # descer a tela
    pyautogui.scroll(-100)

    # clicar em estoque
    encontrou = encontrar_imagem('estoque.png')
    pyautogui.click(clicar_direita(encontrou))
    estoque_texto = f'{estoque:.2f}'.replace('.', ',')
    # escrever o estoque
    escrever_texto(str(estoque_texto))

    # clicar em selecionar foto
    encontrou = encontrar_imagem('selecionar.png')
    pyautogui.click(pyautogui.center(encontrou))

    # selecionar a imagem
    encontrou = encontrar_imagem('arquivo.png')
    escrever_texto(rf'C:\Users\Samsung\Downloads\{str(imagem)}')
    pyautogui.press('enter')

    # salvar projeto
    encontrou = encontrar_imagem('save.png')
    pyautogui.click(pyautogui.center(encontrou))
