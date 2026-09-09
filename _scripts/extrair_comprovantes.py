#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extrai dados de todos os comprovantes do diretório Daisy e gera JSON estruturado.

Saídas:
  _extraidos/textos.json  -> path_relativo -> texto extraído (OCR/PDF)
  _extraidos/dados.json   -> estrutura completa com campos parseados
  _extraidos/resumo.txt   -> relatório de organização por competência
"""
import os
# limita threads OpenMP do tesseract para evitar saturação de CPU
os.environ.setdefault("OMP_THREAD_LIMIT", "1")
os.environ.setdefault("OMP_NUM_THREADS", "1")
import re
import json
import subprocess
import sys
import datetime
from concurrent.futures import ProcessPoolExecutor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import BASE, COMPROVANTES_DIR, EXTRAIDOS, DADOS  # noqa: E402

OUT_DIR = EXTRAIDOS
TEXTOS_PATH = os.path.join(OUT_DIR, "textos.json")
TEXTOS_META_PATH = os.path.join(OUT_DIR, "textos_meta.json")
DADOS_PATH = DADOS
RESUMO_PATH = os.path.join(OUT_DIR, "resumo.txt")

IMG_EXTS = {".png", ".jpg", ".jpeg"}
PDF_EXT = ".pdf"

MESES_ABREV = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}
MESES_FULL = {
    "janeiro": 1, "fevereiro": 2, "marco": 3, "março": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
    "outubro": 10, "novembro": 11, "dezembro": 12,
}
MESES_ABREV_CURTO = {
    "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10,
    "nov": 11, "dez": 12, "jan": 1, "fev": 2, "mar": 3, "abr": 4,
}


def categoria_da_pasta(pasta):
    """Remove sufixo de timestamp e espaços do nome da pasta."""
    nome = re.sub(r"-\d{8}T\d{6}Z-\d+-\d+$", "", pasta).strip()
    return nome


def listar_arquivos():
    """Retorna lista de dicts com categoria, caminho rel e caminho abs."""
    arquivos = []
    for root, dirs, files in os.walk(COMPROVANTES_DIR):
        rel_root = os.path.relpath(root, COMPROVANTES_DIR)
        if rel_root == ".":
            continue
        pasta_top = rel_root.split(os.sep)[0]
        cat = categoria_da_pasta(pasta_top)
        for f in sorted(files):
            if f.startswith(".~lock"):
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext not in IMG_EXTS and ext != PDF_EXT:
                continue
            rel = os.path.join("comprovantes", rel_root, f)
            partes = rel_root.split(os.sep)
            subpasta = os.sep.join(partes[1:]) if len(partes) > 1 else ""
            arquivos.append({
                "categoria": cat,
                "pasta_top": pasta_top,
                "subpasta": subpasta,
                "caminho_rel": rel,
                "caminho_abs": os.path.join(root, f),
                "ext": ext.lstrip("."),
                "nome": f,
            })
    return arquivos


def ocr_image(path):
    """OCR de uma imagem. Tenta psm 6, fallback psm 3."""
    for psm in ("6", "3"):
        r = subprocess.run(
            ["tesseract", path, "stdout", "-l", "por+eng", "--psm", psm],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        txt = (r.stdout or "").strip()
        if len(txt) >= 20:
            return txt
    return txt


def extract_pdf_text(path):
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        parts = [pg.extract_text() or "" for pg in pdf.pages]
    return "\n".join(parts)


def ocr_pdf(path):
    """Renderiza páginas do PDF em PNG e faz OCR em cada uma."""
    import tempfile
    out = ""
    with tempfile.TemporaryDirectory() as td:
        prefix = os.path.join(td, "pg")
        subprocess.run(
            ["pdftoppm", "-r", "200", "-png", path, prefix],
            capture_output=True,
        )
        for f in sorted(os.listdir(td)):
            if f.endswith(".png"):
                out += ocr_image(os.path.join(td, f)) + "\n"
    return out.strip()


def extrair_texto(arq):
    """Extrai texto de um arquivo (imagem ou PDF)."""
    path = arq["caminho_abs"]
    ext = arq["ext"]
    if ext in ("png", "jpg", "jpeg"):
        return ocr_image(path)
    else:  # pdf
        txt = extract_pdf_text(path)
        if len(txt.strip()) < 30:
            txt = ocr_pdf(path)
        return txt


def parse_valor(texto):
    """Extrai valor monetário (float) com heurística de prioridade."""
    def to_float(s):
        s = s.replace("R$", "").replace(" ", "").strip()
        s = re.sub(r"[^\d,\.]", "", s)
        if not s:
            return None
        if "," in s:
            # assume vírgula decimal
            s = s.replace(".", "").replace(",", ".")
        try:
            return round(float(s), 2)
        except ValueError:
            return None

    patterns = [
        r"Valor original\s*R\$\s*([\d.,]+)",
        r"Valor\s*R\$\s*([\d.,]+)",
        r"Total a Pagar[:]?\s*R\$\s*([\d.,]+)",
        r"Total a pagar\s*R\$\s*([\d.,]+)",
        r"Valor do Documento\s*R\$\s*([\d.,]+)",
        r"\(=\)\s*Valor do Documento\s*R\$\s*([\d.,]+)",
        r"Valor\s+documento[\s\S]{0,60}?(\d{1,3}(?:\.\d{3})*,\d{2})",
        r"Valor Total da NFS-e\s*([\d.,]+)",
        r"Valor Líquido da NFS-e\s*([\d.,]+)",
        r"VALOR TOTAL[:]?\s*R\$\s*([\d.,]+)",
        r"[Vv]alor [Tt]otal[:]?\s*([\d.,]+)",
        r"VALOR TOTAL[:]?\s*([\d.,]+)",
        r"Total\s*R\$\s*([\d.,]+)",
        r"R\$\s*([\d.,]+)",
    ]
    for pat in patterns:
        m = re.search(pat, texto, re.IGNORECASE)
        if m:
            v = to_float(m.group(1))
            if v is not None and v > 0:
                return v
    # coleta todos os valores e devolve o mais frequente (heurística p/ faturas)
    vals = [to_float(m) for m in re.findall(r"R\$\s*([\d.,]+)", texto)]
    vals = [v for v in vals if v is not None and v > 0]
    if vals:
        return max(set(vals), key=vals.count)
    return None


def parse_datas(texto):
    """Extrai datas dd/mm/aaaa e dd MES aaaa do texto."""
    datas = []
    for m in re.finditer(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", texto):
        d, mo, y = m.group(1), m.group(2), m.group(3)
        y = int(y)
        if y < 100:
            y += 2000
        datas.append(("dd/mm", int(d), int(mo), y))
    for m in re.finditer(
        r"(\d{1,2})\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[.]?\s+(\d{4})",
        texto, re.IGNORECASE,
    ):
        d = int(m.group(1))
        mes = MESES_ABREV[m.group(2).upper()]
        y = int(m.group(3))
        datas.append(("dd_mes", d, mes, y))
    for m in re.finditer(
        r"(\d{1,2})\s+de\s+(janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de\s+(\d{4})",
        texto, re.IGNORECASE,
    ):
        d = int(m.group(1))
        mes = MESES_FULL.get(m.group(2).lower())
        y = int(m.group(3))
        if mes:
            datas.append(("dd_mes_ext", d, mes, y))
    return datas


def _ano_fallback(texto, nome):
    """Resolve o ano da competência: prioriza ano no texto, depois no nome, default 2026."""
    for src in (texto, nome):
        ym = re.search(r"(20\d{2})", src)
        if ym:
            return int(ym.group(1))
    return 2026


def parse_competencia(texto, nome):
    """Extrai competência (mês/ano de referência)."""
    # 1) "Referente a MAI/2026", "competencia ... MAI/2026"
    m = re.search(r"[Rr]eferente\s+a\s*([A-Za-z]{3})/(\d{4})", texto)
    if m:
        mes = MESES_ABREV.get(m.group(1).upper())
        if mes:
            return f"{int(m.group(2)):04d}-{mes:02d}", m.group(0)
    m = re.search(r"[Cc]ompet[eê]ncia[^\n]*?([A-Za-z]{3})/(\d{4})", texto)
    if m:
        mes = MESES_ABREV.get(m.group(1).upper())
        if mes:
            return f"{int(m.group(2)):04d}-{mes:02d}", m.group(0)
    # 2) Mês no nome do arquivo
    nome_l = nome.lower()
    m = re.search(
        r"(janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)",
        nome_l,
    )
    if m:
        mes = MESES_FULL.get(m.group(1).lower())
        ano = _ano_fallback(texto, nome)
        return f"{ano:04d}-{mes:02d}", f"nome: {m.group(0)}"
    m = re.search(r"\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\b", nome_l)
    if m:
        mes = MESES_ABREV_CURTO.get(m.group(1))
        ano = _ano_fallback(texto, nome)
        return f"{ano:04d}-{mes:02d}", f"nome: {m.group(0)}"
    # 3) Data da transação (comprovantes Pix) -> mês da data
    datas = parse_datas(texto)
    if datas:
        # pega a primeira data com dia/mes/ano
        for tipo, d, mo, y in datas:
            return f"{y:04d}-{mo:02d}", f"data_transacao: {d:02d}/{mo:02d}/{y}"
    return None, None


def parse_estabelecimento(texto):
    """Extrai nome do estabelecimento/destino de um comprovante Pix ou boleto."""
    # Comprovante Nubank: "Destino\nNome MEU PRATA LOJA 1"
    m = re.search(r"Destino\s*\n\s*Nome\s+(.+)", texto, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # Boleto: "Beneficiário" / "Sacado" / "Nome do Beneficiário"
    for pat in [
        r"[Bb]enefici[áa]rio\s*[:\n]\s*(.+)",
        r"[Ee]mitente\s*[:\n]\s*(.+)",
        r"[Cc]edente\s*[:\n]\s*(.+)",
        r"[Ff]avorecido\s*[:\n]\s*(.+)",
    ]:
        m = re.search(pat, texto)
        if m:
            return m.group(1).split("\n")[0].strip()
    return None


def parse_data_pagamento(texto):
    """Extrai a data mais provável de pagamento (dd/mm/aaaa)."""
    datas = parse_datas(texto)
    if datas:
        for tipo, d, mo, y in datas:
            if 1 <= d <= 31 and 1 <= mo <= 12:
                return f"{y:04d}-{mo:02d}-{d:02d}"
    return None


def processar_arquivo(args):
    arq = args
    texto = extrair_texto(arq)
    return arq, texto


def assinatura_arquivo(arq):
    try:
        tamanho = os.path.getsize(arq["caminho_abs"])
        alterado = int(os.path.getmtime(arq["caminho_abs"]) * 1000)
        return f"{tamanho}:{alterado}"
    except OSError:
        return ""


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    arquivos = listar_arquivos()
    print(f"Total de arquivos a processar: {len(arquivos)}", flush=True)

    # --- Fase 1: extração de texto (paralelo) ---
    textos = {}
    textos_meta = {}
    if os.path.exists(TEXTOS_PATH):
        with open(TEXTOS_PATH, encoding="utf-8") as fh:
            textos = json.load(fh)
        print(f"Retomando {len(textos)} textos já extraídos.", flush=True)
    if os.path.exists(TEXTOS_META_PATH):
        with open(TEXTOS_META_PATH, encoding="utf-8") as fh:
            textos_meta = json.load(fh)

    rels_validos = {a["caminho_rel"] for a in arquivos}
    textos = {rel: txt for rel, txt in textos.items() if rel in rels_validos}
    textos_meta = {rel: sig for rel, sig in textos_meta.items() if rel in rels_validos}

    pendentes = [
        a for a in arquivos
        if a["caminho_rel"] not in textos
        or textos_meta.get(a["caminho_rel"]) != assinatura_arquivo(a)
    ]
    print(f"Pendentes: {len(pendentes)}", flush=True)

    workers = 4
    with ProcessPoolExecutor(max_workers=workers) as ex:
        for i, (arq, texto) in enumerate(ex.map(processar_arquivo, pendentes), 1):
            textos[arq["caminho_rel"]] = texto
            textos_meta[arq["caminho_rel"]] = assinatura_arquivo(arq)
            if i % 5 == 0:
                with open(TEXTOS_PATH, "w", encoding="utf-8") as fh:
                    json.dump(textos, fh, ensure_ascii=False, indent=2)
                with open(TEXTOS_META_PATH, "w", encoding="utf-8") as fh:
                    json.dump(textos_meta, fh, ensure_ascii=False, indent=2)
                print(f"  ... {i}/{len(pendentes)} processados, salvo parcial.", flush=True)

    with open(TEXTOS_PATH, "w", encoding="utf-8") as fh:
        json.dump(textos, fh, ensure_ascii=False, indent=2)
    with open(TEXTOS_META_PATH, "w", encoding="utf-8") as fh:
        json.dump(textos_meta, fh, ensure_ascii=False, indent=2)
    print(f"Textos extraídos salvos em {TEXTOS_PATH}", flush=True)

    # --- Fase 2: parsing e montagem da estrutura ---
    dados = {"meta": {}, "categorias": {}, "por_mes": {}}
    dados["meta"] = {
        "gerado_em": datetime.datetime.now().isoformat(),
        "base": BASE,
        "total_arquivos": len(arquivos),
    }
    for arq in arquivos:
        cat = arq["categoria"]
        rel = arq["caminho_rel"]
        texto = textos.get(rel, "")
        valor = parse_valor(texto)
        competencia_iso, competencia_origem = parse_competencia(texto, arq["nome"])
        data_pag = parse_data_pagamento(texto)
        estab = parse_estabelecimento(texto)
        item = {
            "categoria": cat,
            "subpasta": arq["subpasta"],
            "arquivo": arq["nome"],
            "caminho_rel": rel,
            "tipo": arq["ext"],
            "valor": valor,
            "valor_original": None,
            "competencia": competencia_iso,
            "competencia_origem": competencia_origem,
            "data_pagamento": data_pag,
            "estabelecimento": estab,
            "texto_extraido": texto,
        }
        dados["categorias"].setdefault(cat, []).append(item)
        if competencia_iso:
            dados["por_mes"].setdefault(competencia_iso, []).append({
                "categoria": cat,
                "arquivo": arq["nome"],
                "caminho_rel": rel,
                "valor": valor,
                "estabelecimento": estab,
                "data_pagamento": data_pag,
            })

    # ordena por_mes
    dados["por_mes"] = {k: dados["por_mes"][k] for k in sorted(dados["por_mes"])}

    with open(DADOS_PATH, "w", encoding="utf-8") as fh:
        json.dump(dados, fh, ensure_ascii=False, indent=2)
    print(f"Dados salvos em {DADOS_PATH}", flush=True)

    # --- Fase 3: resumo ---
    linhas = []
    linhas.append("RESUMO DE ORGANIZAÇÃO POR COMPETÊNCIA (data-mês)")
    linhas.append("=" * 60)
    total_valor = 0.0
    total_sem_valor = 0
    for cat in sorted(dados["categorias"]):
        itens = dados["categorias"][cat]
        soma = sum(i["valor"] or 0 for i in itens)
        sem = sum(1 for i in itens if i["valor"] is None)
        linhas.append(f"\n[{cat}]  {len(itens)} comprovante(s)  soma=R$ {soma:,.2f}  sem_valor={sem}")
        total_valor += soma
        total_sem_valor += sem
    linhas.append("\n" + "-" * 60)
    linhas.append(f"TOTAL: {len(arquivos)} arquivos | soma valores = R$ {total_valor:,.2f} | sem valor = {total_sem_valor}")

    linhas.append("\n\nDISTRIBUIÇÃO POR MÊS:")
    for mes in dados["por_mes"]:
        itens = dados["por_mes"][mes]
        soma = sum(i["valor"] or 0 for i in itens)
        linhas.append(f"  {mes}: {len(itens)} comprovantes  soma=R$ {soma:,.2f}")

    texto_resumo = "\n".join(linhas)
    with open(RESUMO_PATH, "w", encoding="utf-8") as fh:
        fh.write(texto_resumo)
    print("\n" + texto_resumo, flush=True)


if __name__ == "__main__":
    main()
