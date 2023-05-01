#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pyautogui
import pyperclip
import subprocess
import time
import pandas as pd


def encontrar_imagem(imagem): 
    while not pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9):
        time.sleep(1)
    encontrou_img = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9)  # reconhecer a imagem
    return encontrou_img


def direita_img(posicoes_imagem):
    # posicoes_imagem = (x, y, largura, altura)
    return posicoes_imagem[0] + posicoes_imagem[2], posicoes_imagem[1] + posicoes_imagem[3]/2


def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey("ctrl", "v")


pyautogui.FAILSAFE = True

# Abrir o ERP (Fakturama)
subprocess.Popen([r"D:\Fakturama2\Fakturama.exe"])

# fakturama está aberto
encontrou = encontrar_imagem("logo_fakturama.png") # reconhecer a imagem

# fazer em lista
tabela_produtos = pd.read_excel("Produtos.xlsx")

for linha in tabela_produtos.index:
    nome = tabela_produtos.loc[linha, "Nome"]
    id = tabela_produtos.loc[linha, "ID"]
    categoria = tabela_produtos.loc[linha, "Categoria"]
    gtin = tabela_produtos.loc[linha, "GTIN"]
    supplier = tabela_produtos.loc[linha, "Supplier"]
    descricao = tabela_produtos.loc[linha, "Descrição"]
    imagem = tabela_produtos.loc[linha, "Imagem"]
    preco = tabela_produtos.loc[linha, "Preço"]
    custo = tabela_produtos.loc[linha, "Custo"]
    estoque = tabela_produtos.loc[linha, "Estoque"]

    # Clicar no Menu New
    encontrou = encontrar_imagem('menu_new.png')
    pyautogui.click(pyautogui.center(encontrou))

    # Clicar no New Product
    encontrou = encontrar_imagem('botao_newproduct.png')
    pyautogui.click(pyautogui.center(encontrou))

    # Preencher os campos
    encontrou = encontrar_imagem('campo_itemnumber.png')
    pyautogui.click(direita_img(encontrou))
    escrever_texto(str(id))

    pyautogui.press("tab")
    escrever_texto(str(nome))
    pyautogui.press("tab")
    escrever_texto(str(categoria))
    pyautogui.press("tab")
    escrever_texto(str(gtin))
    pyautogui.press("tab")
    escrever_texto(str(supplier))
    pyautogui.press("tab")
    escrever_texto(str(descricao))
    pyautogui.press("tab")
    # prices, atenção muda um pouco. coloca um 0 a mais no numero
    preco_texto = f"{preco:.2f}".replace(".", ",")
    escrever_texto(str(preco))
    pyautogui.press("tab")
    custo_texto = f"{custo:.2f}".replace(".", ",")
    escrever_texto(str(custo))
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    estoque_texto = f"{estoque:.2f}".replace(".", ",")
    escrever_texto(str(estoque))
    # selecionar imagem
    encontrou = encontrar_imagem('botao_selecionarimagem.png')
    pyautogui.click(pyautogui.center(encontrou))
    # esperar encontrar a imagem campo Nome
    encontrou = encontrar_imagem('nome_arquivo.png')
    escrever_texto(rf"C:\Users\Gui\CURSO PYTHON\31- RPA com Python\Automação de ERPs (Totvs, SAP e outros)\Imagens Produtos-20230327T163037Z-001\Imagens Produtos\{str(imagem)}")
    pyautogui.press("enter")

    # clicar em salvar
    encontrou = encontrar_imagem('botao_save.png')
    pyautogui.click(pyautogui.center(encontrou))


# In[ ]:




