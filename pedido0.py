from asyncio.windows_events import NULL
from contextlib import nullcontext
from faulthandler import is_enabled
from xml.sax.xmlreader import Locator
import numpy as np
import pandas as pd
from playwright.sync_api import sync_playwright

excel = pd.read_excel("C:\megatron\pedidos\pedido0.xlsm", sheet_name=2)
user = "971" #usuário do representante
password = "sucesso" #senha do usuário do repre
colecao = "Pedidos - OUTONO 2023" #nome do campo dentro do WISE da coleçao
marca = "COCA-COLA - BE RED" #nome do campo dentro do WISE referente à marca para efetuar o pedido
cliente = excel.iloc[2,2] #Atribui o nome do cliente que esta na planilha para a variável cliente.

#print(excel)


with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto("https://wisesale-cocacola1.azurewebsites.net/WiseSale.aspx")
    page.fill("input[name='txtUsuario']", user)
    page.fill("input[name='txtSenha']", password)
    page.click("input[name='btnEntrar']")
    page.wait_for_timeout(1500)
    page.get_by_text(colecao).click()
    page.locator("a", has_text=marca).click()
    page.wait_for_timeout(2000)
    page.click("input[name='ImgFecharPan']")
    page.fill("input[name='txtRazao_Social']",cliente)
    page.click("input[name='btnLocalizarCliente']")
    page.locator("//*[@id='grdPedidos']/tbody/tr[2]/td[1]/a").click()
    page.click("input[name='ImgFecharPan']")


    linha = 14 #indexação onde começa as referencias do produto
    while linha <=63: #Mudar para 533 quando for rodar o teste, ou para o tamanho da da planilha de pedidos -2. ex: planilha tem 600 referencias mudar esse campo para 598.
        coluna = 7
        produto = str(excel.iloc[linha,1])
        campoCod = page.locator("input[name='txtCodProd']").get_attribute("title")
        page.fill("input[name='txtCodProd']", produto)
        page.click("input[name='btnLocalizarProduto']")
        popUp= page.locator("div[id=''panMsgWise]")
        if popUp != NULL:
            try:
                page.wait_for_selector("input[id='ImgFecharPan']",1000)
                page.click("input[id='ImgFecharPan']")
            except:
                pass

        digitoTam=1
        tamanhosPreenchidos = []
        while coluna <= 13:
            if not np.isnan(excel.iloc[linha,coluna]):
                corProdutoPlanilha = str(excel.iloc[linha,3]) #Extrai a informação de cor da planilha
                verificaCor = page.locator(f"span[title='{corProdutoPlanilha}']").is_visible(timeout=100) #Armazena true para se a cor existir e false se a cor não existir
                verificaTam = page.locator(f"span[id='lblTam{digitoTam}']").inner_text()#Guarda o texto do conteudo da tag HTML, se o tamanho no WISE não existir a tag vai estar por padrão assume um valor de vazio
                if verificaCor != False:
                    if verificaTam != "":
                        corProdutoWise = page.locator(f"span[title='{corProdutoPlanilha}']").inner_text()
                        idCor = page.locator(f"text={corProdutoWise}").get_attribute("id")
                        quantidadeTam = str(excel.iloc[linha,coluna])
                        tamanhosPreenchidos.append(quantidadeTam)
                        selectorTamanho = f"input[name='txtTam{digitoTam}_{idCor[len(idCor)-1]}']"
                        verificaGrade = page.locator(selectorTamanho).is_disabled()
                        if verificaGrade == False:
                            if corProdutoPlanilha in corProdutoWise:
                                page.fill(selectorTamanho, quantidadeTam)
            digitoTam+=1
            coluna+=1
        
        for i in range(len(tamanhosPreenchidos)):
            if int(tamanhosPreenchidos[i]) >= 10:
                page.click(f"input[id='chkSalvar10']")
                break
        #bloco para salvar as quantidades do modelo no pedido
        page.click("input[id='btnSalvar']")
        popUp= page.locator("div[id=''panMsgWise]")
        if popUp != NULL:
            try:
                page.wait_for_selector("input[id='ImgFecharPan']",1000)
                page.click("input[id='ImgFecharPan']")
            except:
                pass
        linha+=1

    page.wait_for_timeout(5000)