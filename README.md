# Automação de Nota Fiscal

## Descrição

Este projeto é uma automação para processar arquivos de Nota Fiscal eletrônica (NF-e) em formato XML. O script lê arquivos XML de NF-e, extrai informações relevantes, e as salva em um arquivo Excel. A automação é útil para centralizar e analisar dados de notas fiscais eletrônicas.

## Funcionalidades

- **Leitura de Arquivos XML**: Processa arquivos de Nota Fiscal eletrônica (NF-e) em formato XML.
- **Extração de Dados**: Extrai informações como número da nota, nome da empresa emissora, nome do cliente, endereço e peso dos produtos.
- **Exportação para Excel**: Salva os dados extraídos em um arquivo Excel (`.xlsx`) para fácil análise e arquivamento.

## Tecnologias Utilizadas

- **Python**: Linguagem de programação usada para o script de automação.
- **xmltodict**: Biblioteca para conversão de XML para dicionário Python.
- **pandas**: Biblioteca para manipulação e análise de dados, usada para salvar os dados em formato Excel.
- **openpyxl**: Biblioteca usada pelo pandas para exportar dados para Excel.

## Pré-requisitos

Antes de rodar o script, você precisa instalar as seguintes bibliotecas Python:

```bash
pip install xmltodict pandas openpyxl
