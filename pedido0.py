from asyncio.windows_events import NULL
from contextlib import nullcontext
from faulthandler import is_enabled
from xml.sax.xmlreader import Locator
import numpy as np
import pandas as pd
from playwright.sync_api import sync_playwright

excel = pd.read_excel("C:\Riot Games\Planilha_de_Pedido_JPC-_OUT23_-_MARIA_FILÓ.xlsm", sheet_name=0)
user = "Rodrigo" #usuário do representante
password = "r0d1" #senha do usuário do repre
colecao = "Pedidos - Coleção Atual" #nome do campo dentro do WISE da coleçao
marca = "Pedidos - INVERNO 2023" #nome do campo dentro do WISE referente à marca para efetuar o pedido
cliente = excel.iloc[2,2] #Atribui o nome do cliente que esta na planilha para a variável cliente

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto("http://soma-ws.compuwise.com.br/")
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
    produto = "inicio"

    while produto != "": #Mudar para 533 quando for rodar o script, para o tamanho da da planilha de pedidos -2. ex: plnilha tem 600 referencias mudar esse campo para 598.
        coluna = 7
        produto = str(excel.iloc[linha,1])
        campoCod = page.locator("input[name='txtCodProd']").get_attribute("title")

        if produto != str(excel.iloc[linha-1,1]):
            page.fill("input[name='txtCodProd']", produto)
            page.click("input[name='btnLocalizarProduto']")

            if page.locator("span[id='lblDescricao']").inner_html() == "":
                page.locator("input[name='btnLimpar']").click()
                linha +=1
                print(f'Referência {produto} não existe.')
                continue

            try:
                if page.locator("//*[@id='lblMsgWise']").is_visible():
                    page.wait_for_timeout(1000)
                    popUp = page.locator("//*[@id='lblMsgWise']")
                    if popUp.inner_html() == "Produto desabilitado!\n\n":
                        page.locator("input[id='ImgFecharPan']").click()
                        page.locator("input[name='btnLimpar']").click()
            except:
                pass
        else:
            if popUp != NULL:
                try:
                    page.wait_for_selector("input[id='ImgFecharPan']",1000)
                    page.click("input[id='ImgFecharPan']")
                except:
                    pass

        digitoTam=1
        tamanhosPreenchidos = []
        tamanhosPlanilha = []

        while coluna <= 13:
            if not pd.isnull(excel.iloc[linha,coluna]):
                corProdutoPlanilha = str(excel.iloc[linha,3]) #Extrai a informação de cor da planilha
                verificaCor = page.locator(f"span[title='{corProdutoPlanilha}']").is_visible(timeout=100) #Armazena true para se a cor existir e false se a cor não existir
                indexTamanho = 7

                if len(tamanhosPlanilha) == 0:
                    for i in range(1,8):
                        #tamanhosPlanilha.append(str(excel.iloc[i,coluna])
                        tamanhosPlanilha.append(str(excel.iloc[linha,indexTamanho]))
                        indexTamanho +=1
                
                digitoTam = tamanhosPlanilha.index(str(excel.iloc[linha,coluna]))+1
                tamanhosPlanilha[digitoTam-1] = ""
               
                
                for i in range(1,7):
                    try:
                        if page.locator(f"span[id='lblTam{digitoTam}']").is_visible():
                            verificaTam = page.locator(f"span[id='lblTam{digitoTam}']").inner_text()#Guarda o texto do conteudo da tag HTML, se o tamanho no WISE não existir a tag vai estar por padrão assume um valor de vazio
                            break
                        else:
                            digitoTam+=1
                    except:
                        break

                if verificaCor != False:
                    if verificaTam != "":
                        corProdutoWise = page.locator(f"span[title='{corProdutoPlanilha}']").inner_text()
                        idCor = page.locator(f"text={corProdutoWise}").get_attribute("id")
                        if np.char.isnumeric(str(excel.iloc[linha,coluna])):
                            quantidadeTam = str(excel.iloc[linha,coluna])
                            tamanhosPreenchidos.append(quantidadeTam)
                            if page.locator(f"span[id='lblTam1']").inner_text() == 'UN':
                                selectorTamanho = f"input[name='txtTam{digitoTam+1}_{idCor[len(idCor)-1]}']"
                            else:
                                selectorTamanho = f"input[name='txtTam{digitoTam}_{idCor[len(idCor)-1]}']"
                            verificaGrade = page.locator(selectorTamanho).is_disabled()
                            if verificaGrade == False and corProdutoPlanilha in corProdutoWise:
                                #if corProdutoPlanilha in corProdutoWise:
                                page.fill(selectorTamanho, quantidadeTam)
                            # else:
                            #     try:
                            #         auxInteiro = int(quantidadeTam) + int(tamanhosPlanilha[digitoTam])
                            #         tamanhosPlanilha[digitoTam] = str(auxInteiro)
                            #         excel.iloc[linha,coluna+1] = str(int(excel.iloc[linha,coluna+1])+int(excel.iloc[linha,coluna]))
                            #     except:
                            #         pass
                        # else:
                        #     temp = int(excel.iloc[linha,coluna+1]) + 2
                        #     excel.iloc[linha,coluna+1] = str(temp)
                else:
                    pass

            coluna+=1
        
        for i in range(len(tamanhosPreenchidos)):
            if int(tamanhosPreenchidos[i]) > 9:
                verificarCheckbox = page.locator(f"input[id='chkSalvar10']")
                if verificarCheckbox.is_checked():
                    break
                else:
                    page.click(f"input[id='chkSalvar10']")
                    break
            
        #bloco para salvar as quantidades do modelo no pedido
        try:
            if str(excel.iloc[linha+1,1]) != produto:
                if verificaCor == True:
                    page.locator("input[id='btnSalvar']").click()
                    popUp= page.locator("div[id=''panMsgWise]")
                    if popUp != NULL:
                        try:
                            page.wait_for_selector("input[id='ImgFecharPan']",1000)
                            page.click("input[id='ImgFecharPan']")
                        except:
                            pass
                else:
                    page.locator("input[name='btnLimpar']").click()
            else:
                pass
        except:
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