import os
import re
from openpyxl import Workbook
import pdfplumber
from datetime import datetime
import unicodedata

# ===== Config =====
caminho_pasta = 'Banco de talentos - SOLIDES'

# ===== Planilha =====
wb = Workbook()
aba = wb.active
aba.title = 'Informacoes_curriculos'
aba['A1'] = 'Nome'
aba['B1'] = 'Cidade'

# ===== Utils =====
def corrige_acentos(txt: str) -> str:
    if not txt:
        return ""
    # tenta consertar PDFs com encoding latin1->utf8 bugado
    try:
        return txt.encode('latin1').decode('utf-8')
    except Exception:
        return txt

def limpar_cidade(s: str) -> str:
    s = re.sub(r'^\s*(Cidade\s*[:\-]?\s*)', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s*-\s*(Brasil|Brazil)\s*$', '', s, flags=re.IGNORECASE)
 
    return s.strip()

def is_city_line(s: str) -> bool:
    s2 = s.strip()
    if not s2:
        return False
    if re.search(r'\d', s2):                       # linhas com número (ex: "12345", "CEP")
        return True
    if re.search(r'^\s*Cidade\s*[:\-]?', s2, flags=re.IGNORECASE):
        return True
    if re.search(r'\b-\s*[A-Z]{2}\b', s2):         # "Catarina - CE"
        return True
    return False

def is_age_line(s: str) -> bool:
    return bool(re.search(r'\b\d{1,3}\s*anos\b', s, flags=re.IGNORECASE))

def is_contact_line(s: str) -> bool:
    # telefone
    if re.search(r'\b\(?\d{2}\)?\s*\d{4,5}[-\s]?\d{4}\b', s):
        return True
    # email
    if re.search(r'[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}', s, flags=re.IGNORECASE):
        return True
    # docs
    if re.search(r'\bCPF\b|\bRG\b', s, flags=re.IGNORECASE):
        return True
    return False

def looks_like_name(s: str) -> bool:
    s2 = s.strip()
    if not s2 or s2.endswith(':'):
        return False
    if re.search(r'\d', s2):
        return False
    parts = [p for p in re.split(r'\s+', s2) if p]
    # 2–6 palavras com letras/acentos/hífens/ponto
    if not (2 <= len(parts) <= 6):
        return False
    return all(re.fullmatch(r"[A-Za-zÀ-ÿ'´`^~\-\.]+", p) for p in parts)

def extrair_nome_fallback_arquivo(nome_arquivo: str):
    base = os.path.splitext(os.path.basename(nome_arquivo))[0]
    base = re.sub(r"[_\-\.]", " ", base)
    base = re.sub(r"\d+", " ", base)
    base = re.sub(r"(?<=[a-zà-ÿ])(?=[A-ZÀ-Ý])", " ", base)
    palavras = [p for p in base.split() if len(p) > 1]
    if 2 <= len(palavras) <= 6:
        return " ".join(p.capitalize() for p in palavras)
    return None

def pick_name_from_first_page(lines):
    # escolhe a primeira linha válida entre as 8 primeiras
    for l in lines[:8]:
        cand = l.strip()
        if not cand:
            continue
        if is_city_line(cand) or is_age_line(cand) or is_contact_line(cand):
            continue
        if looks_like_name(cand):
            return cand
    return None

def extrair_cidade_pg1(lines):
    """
    Regra solicitada:
    - cidade na 3ª linha (index 2);
    - se tiver número -> pula para a próxima linha;
    - tenta até a 5ª linha (index 4);
    - limpa rótulos/UF/Brasil.
    """
    for idx in (2, 3, 4):
        if len(lines) > idx:
            cand = lines[idx].strip()
            if not cand:
                continue
            # se tiver número, pula para a próxima
            if re.search(r'\d', cand):
                continue
            c = limpar_cidade(cand)
            if c and c.lower() not in {"pais", "país", "estado", "endereco", "endereço"}:
                return c
    return "Erro!"

# ===== Processamento =====
arquivos = [f for f in os.listdir(caminho_pasta) if f.lower().endswith('.pdf')]
linha_xlsx = 2

for arquivo in arquivos:
    caminho_arquivo = os.path.join(caminho_pasta, arquivo)
    with pdfplumber.open(caminho_arquivo) as pdf:
        # 1ª página
        pg1 = pdf.pages[0]
        txt_pg1 = corrige_acentos(pg1.extract_text() or "")
        linhas_pg1 = [corrige_acentos(x) for x in txt_pg1.split("\n")]

        # --- NOME ---
        # (1) tenta "Nome:" em TODO o documento
        texto_full = ""
        for p in pdf.pages:
            texto_full += corrige_acentos(p.extract_text() or "") + "\n"
        m = re.search(r"\bNome(\s+completo)?\s*[:\-]\s*([A-Za-zÀ-ÿ'´`^~\-\.\s]{3,})", texto_full, re.IGNORECASE)
        if m:
            nome = m.group(2).strip()
        else:
            # (2) escolhe a primeira linha válida do topo da 1ª página (evita cidade/idade/contato)
            nome = pick_name_from_first_page(linhas_pg1)
            # (3) fallback final: nome do arquivo
            if not nome:
                nome = extrair_nome_fallback_arquivo(arquivo) or "Erro!"

        # --- CIDADE ---
        cidade = extrair_cidade_pg1(linhas_pg1)

        # escreve na planilha
        aba[f"A{linha_xlsx}"] = nome
        aba[f"B{linha_xlsx}"] = cidade
        linha_xlsx += 1

# ===== Salvar =====
agora = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
saida = f'Informacoes_{agora}.xlsx'
wb.save(saida)
print(f"Extração concluída com sucesso! Arquivo: {saida} ❤️🚀")
