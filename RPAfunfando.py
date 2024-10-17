import pandas as pd
from zenrows import ZenRowsClient
from lxml import html
import openpyxl
from openpyxl.styles import Alignment, numbers

# Defina sua chave da API ZenRows
api_key = "SUA CHAVE API"
client = ZenRowsClient(api_key)

# Define a URL da API da Casa dos Dados
url_api = 'https://api.casadosdados.com.br/v2/public/cnpj/search'

# Cria o DataFrame vazio para usarmos daqui a pouco
df_final = pd.DataFrame()

print('Iniciando raspagem de dados...')

# Loop que faz a paginação
for i in range(1, 2):
    data = {
        "query": {
            "termo": [],
            "atividade_principal": [],
            "natureza_juridica": [],
            "uf": ["SP"],
            "municipio": [],
            "bairro": [],
            "situacao_cadastral": "ATIVA",
            "cep": [],
            "ddd": []
        },
        "range_query": {
            "data_abertura": {
                "lte": None,
                "gte": "2023-06-01"
            },
            "capital_social": {
                "lte": None,
                "gte": None
            }
        },
        "extras": {
            "somente_mei": False,
            "excluir_mei": True,
            "com_email": True,
            "incluir_atividade_secundaria": False,
            "com_contato_telefonico": False,
            "somente_fixo": False,
            "somente_celular": False,
            "somente_matriz": False,
            "somente_filial": False
        },
        "page": i
    }

    print(f'Raspando página {i} ', end='')

    # Realiza a solicitação POST usando ZenRows
    response = client.post(url_api, json=data)

    # Verifica se a solicitação foi bem-sucedida
    if response.status_code == 200:
        resultado = response.json()
    else:
        print(f'Erro na solicitação (Código {response.status_code}): {response.text}')
        break

    df_provisorio = pd.json_normalize(resultado, ['data', 'cnpj'])
    df_final = pd.concat([df_final, df_provisorio], axis=0)

    print(f'- OK')

print('Raspagem inicial feita com sucesso!')

# Aqui inicia a busca por dados adicionais ---------------------------------------------------------------
print('Iniciando extração dos dados adicionais...')

# Vamos usar o DataFrame df_final como fonte dos dados
url = []
for razao, cnpj in zip(df_final['razao_social'], df_final['cnpj']):
    url.append('https://casadosdados.com.br/solucao/cnpj/' + razao.replace(' ', '-').replace('.', '').replace('&', 'and').replace('/', '').replace('*', '').replace('--', '-').lower() + '-' + cnpj)
    
# Inicia as listas vazias para receberem os dados
lista_email = []
lista_tel = []
lista_socio1 = []
lista_socio2 = []
lista_socio3 = []
lista_socio4 = []
lista_socio5 = []
lista_capital_social = []

# Função que verifica se uma variável é número - usaremos na validação do capital social
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

# Função auxiliar para ajustar o comprimento das listas
def ajustar_comprimento_listas(*listas):
    max_len = max(len(lista) for lista in listas)
    listas_ajustadas = []
    for lista in listas:
        while len(lista) < max_len:
            lista.append('')  # Preencher listas curtas com valores vazios
        listas_ajustadas.append(lista)
    return listas_ajustadas

# Itera sobre a lista de URLs e usa ZenRows para fazer a requisição
for indice, link in enumerate(url):
    print(f'Empresa {indice + 1}/{len(df_final)} ', end='')

    # Faz a requisição GET com ZenRows
    response = client.get(link)

    # Verificar se a requisição foi bem-sucedida
    if response.status_code == 200:
        # Parsear o conteúdo HTML da página
        page_content = html.fromstring(response.content)

        # Usar os XPath para encontrar os elementos desejados
        email = page_content.xpath('//a[contains(@href, "mailto:")]/text()')
        tel = page_content.xpath('//a[contains(@href, "tel:")]/text()')

        # Verificar se o elemento EMAIL foi encontrado
        lista_email.append(email[0].lower() if email else '')
        # Verificar se o elemento TEL foi encontrado
        lista_tel.append(tel[0].replace('Telefone: ', '') if tel else '')
        
        # Extraindo os sócios
        socio1 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[1]/text()')
        socio2 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[2]/text()')
        socio3 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[3]/text()')
        socio4 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[4]/text()')
        socio5 = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[24]/p[5]/text()')

        lista_socio1.append(socio1[0] if socio1 else '')
        lista_socio2.append(socio2[0] if socio2 else '')
        lista_socio3.append(socio3[0] if socio3 else '')
        lista_socio4.append(socio4[0] if socio4 else '')
        lista_socio5.append(socio5[0] if socio5 else '')

        # Verificando o capital social
        capital_social_elements = page_content.xpath('//*[@id="__nuxt"]/div/section[4]/div[2]/div[1]/div/div[10]/p')
        if capital_social_elements:
            capital_social = capital_social_elements[0].text_content().replace('R$ ', '').replace('.', '').replace(',', '')
            lista_capital_social.append(float(capital_social) if is_number(capital_social) else '')
        else:
            lista_capital_social.append('N/A')  # Ou outra mensagem padrão

    else:
        lista_email.append('ERRO 404')
        lista_tel.append('ERRO 404')
        lista_socio1.append('ERRO 404')
        lista_socio2.append('ERRO 404')
        lista_socio3.append('ERRO 404')
        lista_socio4.append('ERRO 404')
        lista_socio5.append('ERRO 404')

    print(f'- OK')

# Verificar e ajustar as listas de dados antes de criar o DataFrame
(
    lista_tel,
    lista_email,
    lista_socio1,
    lista_socio2,
    lista_socio3,
    lista_socio4,
    lista_socio5,
    lista_capital_social
) = ajustar_comprimento_listas(
    lista_tel,
    lista_email,
    lista_socio1,
    lista_socio2,
    lista_socio3,
    lista_socio4,
    lista_socio5,
    lista_capital_social
)

# Criar o DataFrame com os dados extraídos
df_dados_extraidos = pd.DataFrame({
    'TELEFONE': lista_tel,
    'EMAIL': lista_email,
    'SÓCIO 1': lista_socio1,
    'SÓCIO 2': lista_socio2,
    'SÓCIO 3': lista_socio3,
    'SÓCIO 4': lista_socio4,
    'SÓCIO 5': lista_socio5,
    'CAPITAL SOCIAL': lista_capital_social,
})

df_final = df_final.reset_index(drop=True)
df_dados_extraidos = df_dados_extraidos.reset_index(drop=True)
df_consolidado = pd.concat([df_final, df_dados_extraidos], axis=1)

# Salvando no arquivo XLSX com formatação monetária
nome_arquivo = 'planilha-com-novos-dados'
with pd.ExcelWriter(f'C:/Users/artur/OneDrive/Documentos/novos-dados/{nome_arquivo}.xlsx', engine='openpyxl') as writer:
    df_consolidado.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Formatando a coluna de Capital Social como moeda
    for row in range(2, len(df_consolidado) + 2):  # Começa na linha 2 (1 é o cabeçalho)
        worksheet[f'X{row}'].number_format = 'R$ #,##0.00'  # Formato de moeda

print(f'Arquivo "{nome_arquivo}.xlsx" salvo com sucesso.')
