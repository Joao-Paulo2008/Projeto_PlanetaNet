import os
import re
from openpyxl import Workbook
import pdfplumber
from datetime import datetime

# Caminho das pastas
caminho_pasta = 'Banco de talentos - SOLIDES'

# Cria nova planilha
wb = Workbook()
aba = wb.active
aba.title = 'Informacoes_curriculos'

# Cabeçalhos
aba['A1'] = 'Nome'
aba['B1'] = 'Cidade'

# Pega lista de arquivos PDF
arquivos = [f for f in os.listdir(caminho_pasta) if f.lower().endswith('.pdf')]

# Expressões regulares para capturar Nome e Cidade
x_nome = r'Nome:\s*([A-Za-zÀ-ÿ\s]+)'
x_cidade = r'Endere[cç]o:\s*([A-Za-zÀ-ÿ\s\-]+)'

linha = 2  # começa após o cabeçalho

for arquivo in arquivos:
    caminho_arquivo = os.path.join(caminho_pasta, arquivo)

    with pdfplumber.open(caminho_arquivo) as pdf:
        primeira_pagina = pdf.pages[0]
        texto = primeira_pagina.extract_text()

        # Procura nome e cidade
        procurar_nome = re.search(x_nome, texto, re.IGNORECASE)
        procurar_cidade = re.search(x_cidade, texto, re.IGNORECASE)

        if procurar_nome:
            nome = procurar_nome.group(1).strip()
        else:
            nome = 'Erro!'

        if procurar_cidade:
            cidade = procurar_cidade.group(1).strip()
        else:
            cidade = 'Erro!'

        aba[f'A{linha}'] = nome
        aba[f'B{linha}'] = cidade

        linha += 1

# Salva a planilha com timestamp
agora = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
wb.save(f'Informacoes_{agora}.xlsx')

print("Extração concluída com sucesso!")
