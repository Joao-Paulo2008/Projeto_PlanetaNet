# Extrator de Currículos (PDF → XLSX)

Script em Python que lê uma pasta com currículos em PDF (ex.: exportados do
banco de talentos SOLIDES) e gera uma planilha `.xlsx` com **Nome** e
**Cidade** extraídos de cada currículo.

## Como funciona

- Procura por um padrão `Nome: ...` no texto do PDF.
- Se não encontrar, tenta identificar o nome pelas primeiras linhas da
  primeira página (ignorando linhas que parecem cidade, idade ou contato).
- Como último recurso, deriva o nome a partir do nome do arquivo.
- A cidade é extraída das linhas 3 a 5 da primeira página, removendo rótulos
  como "Cidade:", UF e "Brasil".

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

```bash
python extrair_curriculos.py "Banco de talentos - SOLIDES"
```

Ou definindo a pasta via variável de ambiente:

```bash
export CURRICULOS_DIR="Banco de talentos - SOLIDES"
python extrair_curriculos.py
```

Se nenhum argumento for passado e a variável de ambiente não existir, usa a
pasta padrão `Banco de talentos - SOLIDES` no diretório atual.

O resultado é salvo como `Informacoes_AAAA-MM-DD_HH-MM-SS.xlsx` na pasta onde
o script foi executado.

## Observações

- Os PDFs de currículo e as planilhas geradas **não são versionados**
  (veja `.gitignore`), já que costumam conter dados pessoais.
- Currículos com formato muito diferente do esperado podem gerar `"Erro!"`
  em Nome e/ou Cidade — nesses casos vale checar manualmente.
