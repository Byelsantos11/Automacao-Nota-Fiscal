import os
import json
import xmltodict  
import pandas as pd

# Função para abrir todos os arquivos e pegar informações
def pegar_infos(nome_arquivo, valores):
    print(f"Pegou as informações {nome_arquivo}")
    try:
        with open(f"nfs/{nome_arquivo}", "rb") as arquivo_xml:
            dic_arquivo = xmltodict.parse(arquivo_xml)
            
            # Identificar se o dicionário contém a chave "NFe" ou "nfeProc"
            if "NFe" in dic_arquivo: 
                infos_nf = dic_arquivo["NFe"]["infNFe"]
            else:
                infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
            
            numero_nota = infos_nf["@Id"]
            empresa_emissora = infos_nf["emit"]["xNome"]
            nome_cliente = infos_nf["dest"]["xNome"]
            endereco = infos_nf["dest"]["enderDest"]
            
            if "vol" in infos_nf["transp"]:
                peso = infos_nf["transp"]["vol"]["pesoB"]
            else:
                peso = "Não informado"
            
            valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso])
    
    except Exception as e:
        print(f"Erro ao processar o arquivo {nome_arquivo}: {e}")
        print(json.dumps(dic_arquivo, indent=4))

# Definindo as colunas do DataFrame
colunas = [
    "numero_nota",
    "empresa_emissora",
    "nome_cliente",
    "endereco",
    "peso"
]

# Lista para armazenar os valores extraídos
valores = []

# Percorre todos os arquivos da pasta nfs
lista_arquivos = os.listdir("nfs")
for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)

# Criar o DataFrame
tabela = pd.DataFrame(valores, columns=colunas)

# Salvar o DataFrame em um arquivo Excel
tabela.to_excel("NotasFiscaisExcel.xlsx", index=False, engine='openpyxl')

print("Informações salvas em 'NotasFiscaisExcel.xlsx'")
