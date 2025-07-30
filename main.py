from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import pandas as pd
import win32com.client as win32

# Criar o navegador
driver = webdriver.Chrome()

# Importar a tabela busca que contém os dados para a pesquisa de preços
try:
    tabela_produtos = pd.read_excel("buscas.xlsx")
    print("Tabela de produtos carregada:")
    print(tabela_produtos)
except FileNotFoundError:
    print(
        "Erro: O arquivo 'buscas.xlsx' não foi encontrado. Certifique-se de que ele está no mesmo diretório do script.")
    driver.quit()
    exit()


def verificar_tem_termos_banidos(lista_termos_banidos, nome):
    '''
    A função verificar_tem_termos_banidos analisa se há termos não desejados no nome extraído na pesquisa.
    '''
    for palavra in lista_termos_banidos:
        if palavra in nome:
            return True
    return False


def verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome):
    '''
    A função verificar_tem_todos_termos_produto analisa se há todos os termos desejados no nome extraído na pesquisa.
    '''
    for palavra in lista_termos_nome_produto:
        if palavra not in nome:
            return False
    return True


def busca_bing_compras(driver, produto, termos_banidos, preco_minimo, preco_maximo):
    '''
    A função busca_bing_compras realiza a primeira pesquisa, no site Bing, para extrair o nome do produto, o preço e o
    link que são adicionados em uma lista de ofertas.
    '''
    # Tratar os valores
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nome_produto = produto.split(' ')
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    # Criar uma lista para adicionar as ofertas
    lista_ofertas = []

    # Abrir o navegador
    driver.get('https://www.bing.com/')

    # Pesquisar o produto desejado
    try:
        time.sleep(3)
        campo_pesquisa = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'q')))
        campo_pesquisa.send_keys(produto)
        campo_pesquisa.send_keys(Keys.ENTER)

        # Clicar em compras
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'b-scopeListItem-shop'))).click()

        # Mudar para a segunda aba após clicar em compras
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])

        # Obter as informações dos produtos
        WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="br-paidOffersGrid"]/li'))
        )
        lista_produtos = driver.find_elements(By.XPATH, '//*[@id="br-paidOffersGrid"]/li')

        for resultado in lista_produtos:
            try:
                nome_element = WebDriverWait(resultado, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'br-pdItemName-noHover'))
                )
                nome = nome_element.text.lower()

                tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
                tem_todos_termos_produtos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)

                if not tem_termos_banidos and tem_todos_termos_produtos:
                    try:
                        preco_element = WebDriverWait(resultado, 5).until(
                            EC.presence_of_element_located(
                                (By.CLASS_NAME, 'pd-price.br-standardPrice.promoted.br-dealPrice.nonOgColor'))
                        )
                        preco_str = preco_element.text.replace('R$', '').replace('.', '').replace(',', '.')
                        preco_float = float(preco_str)

                        if preco_minimo <= preco_float <= preco_maximo:
                            link_element = WebDriverWait(resultado, 5).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'br-compareSellers.b_hide.sj_spcls'))
                            )
                            link = link_element.get_attribute('href')
                            lista_ofertas.append((nome, preco_float, link))
                    except Exception as e:
                        continue
            except Exception as e:
                continue
    except Exception as e:
        print(f'Erro geral na busca do Bing para "{produto}": {e}')
        return []

    # Fechar a aba atual (Compras)
    driver.close()
    # Voltar para a aba original (Bing Search results)
    driver.switch_to.window(driver.window_handles[0])

    return lista_ofertas


def busca_buscape(driver, produto, termos_banidos, preco_minimo, preco_maximo):
    '''
    A função busca_buscape realiza a segunda pesquisa, no site Buscapé, para extrair o nome do produto, o preço e o
    link que também são adicionados em uma lista de ofertas.
    '''
    # Tratar os valores
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nome_produto = produto.split(' ')
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    # Criar uma lista para adicionar as ofertas
    lista_ofertas = []

    # Pesquisar o site Buscapé
    try:
        campo_pesquisa = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'q')))
        campo_pesquisa.clear()
        campo_pesquisa.send_keys('Buscapé')
        campo_pesquisa.send_keys(Keys.ENTER)
        # Esse tempo é para apertar no pop up, mas é preciso criar um código para isso
        time.sleep(5)

        # --- Clicar no site Buscapé ---
        link_buscape = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CLASS_NAME, 'b_adurl'))
                                                       )
        link_buscape.click()

        # Mudar para a segunda aba do Buscapé
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])

        # Esperar até a barra de busca do Buscapé ficar visível
        busca_buscape_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input'))
        )
        busca_buscape_input.send_keys(produto, Keys.ENTER)

        # Obter as informações dos produtos
        WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, 'Hits_ProductCard__Bonl_'))
        )
        lista_produtos = driver.find_elements(By.CLASS_NAME, 'Hits_ProductCard__Bonl_')

        for resultado in lista_produtos:
            try:
                nome_element = WebDriverWait(resultado, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'ProductCard_ProductCard_NameWrapper__45Z01'))
                )
                nome = nome_element.text.lower()

                tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
                tem_todos_termos_produtos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)

                if not tem_termos_banidos and tem_todos_termos_produtos:
                    try:
                        preco_element = WebDriverWait(resultado, 5).until(
                            EC.presence_of_element_located(
                                (By.CLASS_NAME, 'Text_Text__ARJdp.Text_MobileHeadingS__HEz7L'))
                        )
                        preco_str = preco_element.text.replace('R$', '').replace('.', '').replace(',', '.')
                        preco_float = float(preco_str)

                        if preco_minimo <= preco_float <= preco_maximo:
                            link_element = WebDriverWait(resultado, 5).until(
                                EC.presence_of_element_located((By.CLASS_NAME, 'ProductCard_ProductCard_Inner__gapsh'))
                            )
                            link = link_element.get_attribute('href')
                            lista_ofertas.append((nome, preco_float, link))
                    except Exception as e:
                        # print(f"Não foi possível encontrar o preço para o produto '{nome}' no Buscapé: {e}")
                        continue
            except Exception as e:
                # print(f"Erro ao obter nome do produto no Buscapé: {e}")
                continue

    except Exception as e:
        print(f'Erro geral na busca do Buscapé para "{produto}": {e}')
        return []

    # Fechar a segunda aba
    driver.close()
    # Mudar para a primeira aba
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(3)

    return lista_ofertas


# Criar uma tabela vazia
tabela_ofertas = pd.DataFrame(columns=['Produto', 'Preço', 'Link'])

for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']

    print(f"\nBuscando: {produto}")

    # Busca no Bing Compras
    lista_ofertas_bing_compras = busca_bing_compras(driver, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_bing_compras:
        print(f"Encontradas {len(lista_ofertas_bing_compras)} ofertas no Bing para {produto}")
        tabela_bing_compras = pd.DataFrame(lista_ofertas_bing_compras, columns=['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_bing_compras], ignore_index=True)
    else:
        print(f"Nenhuma oferta encontrada no Bing para {produto}")

    # Busca no Buscapé
    lista_ofertas_buscape = busca_buscape(driver, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        print(f"Encontradas {len(lista_ofertas_buscape)} ofertas no Buscapé para {produto}")
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape], ignore_index=True)
    else:
        print(f"Nenhuma oferta encontrada no Buscapé para {produto}")

# Exibir a tabela final de ofertas
print("\nTabela final de ofertas:")
print(tabela_ofertas)

# Exportar para Excel
tabela_ofertas.to_excel('Ofertas.xlsx', index=False)

# Enviar o resultado da tabela por email
if len(tabela_ofertas) > 0:
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = 'testando@gmail.com' # Adicionar um e-mail válido
    mail.Subject = 'Produto(s) encontrado(s) na faixa de preço desejada.'
    mail.HTMLBody = f'''
    <h2>Prezados,</h2>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada.</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Ficamos à disposição para sanar qualquer dúvida</p>
    <p>Att.</p>
    <p>Danilo Oliveira dos Santos</p>
    '''
    mail.Send()

# Fechar o navegador no final
driver.quit()