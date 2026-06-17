"""
Extrai código + título de documentos da concessão em PDF, lidando com DOIS layouts:

  Layout A (desenho técnico, tipo DE): carimbo com rótulo "TÍTULO:" / "TITLE:".
            O título vem logo após o código, antes de "PROJETO EXECUTIVO" / "NOTAS".

  Layout B (relatório, tipo RL): carimbo "EMPREENDIMENTO ... CÓDIGO ... REFERÊNCIAS".
            O nome do documento (ex.: PLANO DE DRAGAGEM) está numa linha em negrito,
            sem rótulo "TÍTULO".

Lógica: tenta Layout A; se não vier nada, tenta Layout B; se nenhum funcionar,
retorna vazio. O código (CPSI-DD-...) é extraído por regex e funciona nos dois.

Uso:
    pip install pymupdf
    python extrair_titulo_2layouts.py arquivo1.pdf arquivo2.pdf
"""

import re
import sys
import json
import fitz  # PyMuPDF

# código padrão da concessão; o grupo captura o TIPO (DE, RL, DE, etc.)
COD = re.compile(r'CPSI-DD-\d{2}\.\d{3}-C\d{2}-([A-Z]{2})-\d{3}')

# palavras que NÃO são o nome do documento (projeto, assinaturas, cabeçalho)
RUIDO = ["BRIDGE", "PONTE", "SALVADOR", "ITAPARICA", "REFER", "EMPREEND",
         "GOVERNO", "CONCESS", "ASSINATURA", "ELABORAD", "VERIFIC",
         "APROVA", "DESCRI", "DATA"]


def _codigo(txt):
    m = COD.search(txt)
    return (m.group(0), m.group(1)) if m else (None, None)


def tenta_layout_A(txt):
    """Desenho técnico: rótulo TÍTULO/TITLE; título logo após o código."""
    if not re.search(r'\bTÍTULO\b|\bTITLE\b', txt, re.I):
        return None
    m = COD.search(txt)
    if not m:
        return None
    cauda = txt[m.end():]
    fim = re.search(r'PROJETO EXECUTIVO|\bNOTAS\b|LEGENDA', cauda, re.I)
    bruto = (cauda[:fim.start()] if fim else cauda[:200]).replace("\n", " ")
    bruto = bruto.strip(" -\t")
    bruto = re.sub(r'PROJETO DE DRAGAGEM\s*$', '', bruto, flags=re.I).strip()
    return bruto or None


def tenta_layout_B(doc):
    """Relatório: título é a linha em negrito do carimbo, sem rótulo."""
    if "EMPREENDIMENTO" not in doc[0].get_text().upper():
        return None
    for bloco in doc[0].get_text("dict")["blocks"]:
        for linha in bloco.get("lines", []):
            spans = linha.get("spans", [])
            txt = "".join(s["text"] for s in spans).strip()
            if len(txt) < 5:
                continue
            bold = all(("Bold" in s["font"] or s["flags"] & 16)
                       for s in spans if s["text"].strip())
            up = txt.upper()
            if bold and not any(x in up for x in RUIDO) and not COD.search(txt):
                return txt
    return None


def extrair(pdf_path):
    """Roteia entre os dois layouts. Retorna {} se nenhum funcionar."""
    doc = fitz.open(pdf_path)
    txt = doc[0].get_text()
    codigo, tipo = _codigo(txt)

    # tenta A, se não vier nada tenta B
    titulo = tenta_layout_A(txt)
    layout = "A"
    if not titulo:
        titulo = tenta_layout_B(doc)
        layout = "B"
    doc.close()

    if not titulo and not codigo:
        return {}            # nenhum dos dois -> vazio

    return {
        "arquivo": pdf_path.split("/")[-1],
        "codigo": codigo,
        "tipo": tipo,
        "titulo": titulo,
        "_layout": layout if titulo else None,
    }


if __name__ == "__main__":
    for caminho in sys.argv[1:]:
        print(json.dumps(extrair(caminho), ensure_ascii=False, indent=2))