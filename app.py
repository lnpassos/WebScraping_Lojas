### PROJETO WEB SCRAPING DE PRODUTOS DO GOOGLE SHOPPING E BUSCAPÉ


## MAPA MENTAL:

# criar navegador
# importar/visualizar base de dados
# para cada item na base de dados:
    # 1 - procurar esse produto no google shopping
        # verificar se algum dos produtos do google shopping está dentro da margem de preço
    # 2 - procurar ese produto no buscapé
        #  verificar se algum dos produtos do buscapé está dentro da margem de preço
# Salvar as ofertas boas em um dataframe (tabela)
# exportar por excel
# enviar por email o resultado da tabela

######################################################################################################


## IMPORTS 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import win32com.client as win32

# pesquisar produto
tabela_produtos = pd.read_excel('buscas.xlsx')
tabela_ofertas = pd.DataFrame()

# função para verificar termos banidos
def verificar_tem_termos_banidos(lista_termos_banidos, nome):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome:
            tem_termos_banidos = True
    return tem_termos_banidos

# função para verificar se há todos termos banidos
def verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome):
    tem_todos_termos_produtos = True
    for palavra in lista_termos_nome_produto:
        if palavra not in nome:
            tem_todos_termos_produtos = False
    return tem_todos_termos_produtos

# Criando Navegador
servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome()

# Função para busca Google Shopping
def busca_google_shopping(driver, produto, termos_banidos, preco_minimo, preco_maximo):

    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    produto = produto.lower()
    lista_ofertas = []
    termos_banidos = termos_banidos.lower()
    lista_termos_nome_produto = produto.split(' ')
    lista_termos_banidos = termos_banidos.split(' ')

    # acessar google
    driver.get('https://shopping.google.com/')
    time.sleep(0.6)
                                            
    driver.find_element(By.XPATH, '//*[@id="REsRA"]').send_keys(produto, Keys.ENTER) # Passar os parâmetros diretos separados por 'virgula'
    time.sleep(0.5)

    lista_resultados = driver.find_elements(By.CLASS_NAME, 'i0X6df')

    for resultado in lista_resultados:
        #print(resultado)
        # Puxar apenas o resultado dentro do elemento selecionado e não no driver inteiro.
        nome = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
        nome = nome.lower()

        # Analise se não há termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        # Analisar se ele tem TODOS os termos banidos
        tem_todos_termos_produtos = verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome)

        # Tratar os dados do preço
        if not tem_termos_banidos and tem_todos_termos_produtos: # Tratação Boolean True e False
            preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
            preco = preco.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.').replace('+impostos', '')
            preco = float(preco)

            # Validar se o preço está atendendo os termos definidos
            if preco_minimo <= preco <= preco_maximo: # Basicamente Se o preço minimo < que preço e preço_minimo < que preço_máximo
                elemento_referencia = resultado.find_element(By.CLASS_NAME, 'bONr3b')
                elemento_pai = elemento_referencia.find_element(By.XPATH, '..')
                link = elemento_pai.get_attribute('href') # Puxando atributo
                lista_ofertas.append((nome, preco, link))

    return lista_ofertas

# Funcão de busca para o Buscapé
def busca_buscape(driver, produto, termos_banidos, preco_minimo, preco_maximo):

    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nome_produto = produto.split(' ')
    lista_ofertas = []

    driver.get('https://www.buscape.com.br/')
    time.sleep(0.6)

    driver.find_element(By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, Keys.ENTER)

    # Aguardar enquanto elemento único da página apareça
    while len(driver.find_elements(By.CLASS_NAME, 'SearchFilters_HitsCount__A0m37')) < 1:
        time.sleep(1)

    lista_resultados = driver.find_elements(By.CLASS_NAME, 'ProductCard_ProductCard_Inner__gapsh')

    for resultado in lista_resultados:
        #print(resultado)
        nome = resultado.find_element(By.CLASS_NAME, 'Text_DesktopLabelSAtLarge__wWsED').text
        nome = nome.lower()
        preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__HEz7L').text
        link = resultado.get_attribute('href')

        # Analise se não há termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        # Analisar se ele tem TODOS os termos banidos
        tem_todos_termos_produtos = verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome)

        #print(nome, preco, link)

        if not tem_termos_banidos and tem_todos_termos_produtos:
            preco = preco.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.').replace('+impostos', '')
            preco = float(preco)

            if preco_minimo <= preco <= preco_maximo:
                lista_ofertas.append((nome, preco, link))
                print(nome, preco, link)

    return lista_ofertas


# Separando os dados e armazenando-os na tabela de ofertas
for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(driver, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto', 'preco', 'link'])
        
        # Concatena as ofertas na tabela.
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])
    else:
        tabela_google_shopping = None
        print('Tabela Google vazia!')
    #print(tabela_google_shopping)

    lista_ofertas_buscape = busca_buscape(driver, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])

        # Concatena as ofertas na tabela.
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])
    else:
        tabela_buscape = None
        print('Tabela Buscape vazia!')
    #print(tabela_buscape)

#print(tabela_ofertas)
        
# Exportando para Excel
tabela_ofertas.to_excel('Ofertas.xlsx', index = False)


# enviar por e-mail o resultado da tabela

#verificando se existe alguma oferta dentro da tabela de ofertas
if len(tabela_ofertas.index) > 0:
   
    # Necessário ter o outlook instalado e atualizado no PC.
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'leo.nardo.360@hotmail.com' # Mudar aqui seu Email
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f"""
    <p>Prezados,</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada. Segue tabela com detalhes</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att, Leonardo.</p>
    """
    
    mail.Send()