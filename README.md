# Data Transfer Tool

Este projeto é uma ferramenta desenvolvida em Python que utiliza a biblioteca OpenPyXL para transferir dados entre duas planilhas Excel. 

## Funcionalidades

- Carrega dados de uma planilha de origem.
- Mapeia e compara as LIs (identificadores) entre duas planilhas.
- Copia dados correspondentes de colunas específicas de uma planilha para outra.
- Salva os dados copiados em uma nova planilha.

## Pré-requisitos

- Python 3.x
- OpenPyXL
  
  ```bash
  pip install openpyxl

## Como Usar

1. Coloque suas planilhas `dadosOrigem.xlsx` e `dadosDestino.xlsx` na pasta `planilhas`.
2. Execute o script. Os dados copiados serão salvos em `DadosFinal.xlsx` na mesma pasta.

