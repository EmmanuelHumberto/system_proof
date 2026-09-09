#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ferramenta de extração de comprovantes — reutiliza os regex/extractores por
categoria já validados (extrair_categoria.py) e gera o JSON/CSV para o VBA.

Uso:
  python extrair_ferramenta.py                                   (menu interativo)
  python extrair_ferramenta.py --arquivo <caminho> --categoria <cat>
  python extrair_ferramenta.py --lote <pasta> --categoria <cat>
  python extrair_ferramenta.py --tudo
"""
import os
import sys
import json
import csv
import hashlib
import subprocess
import tempfile
import argparse
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from extrair_categoria import EXTRACTORS, CONFIG, BASE  # noqa: E402
from classificacao import classificar_despesa  # noqa: E402
from indices_categoria import mesclar_indice_categoria  # noqa: E402

IMG_EXTS = {".png", ".jpg", ".jpeg"}

CATEGORIAS = list(EXTRACTORS.keys())


def ocr_image(path):
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
        return "\n".join(pg.extract_text() or "" for pg in pdf.pages)


def ocr_pdf(path):
    out = ""
    with tempfile.TemporaryDirectory() as td:
        prefix = os.path.join(td, "pg")
        subprocess.run(["pdftoppm", "-r", "200", "-png", path, prefix], capture_output=True)
        for f in sorted(os.listdir(td)):
            if f.endswith(".png"):
                out += ocr_image(os.path.join(td, f)) + "\n"
    return out.strip()


def extrair_texto(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in IMG_EXTS:
        return ocr_image(caminho)
    txt = extract_pdf_text(caminho)
    return txt if len(txt.strip()) >= 30 else ocr_pdf(caminho)


def extrair_um(caminho, categoria):
    """Extrai os dados de UM arquivo usando o extractor da categoria."""
    texto = extrair_texto(caminho)
    it = {
        "arquivo": os.path.basename(caminho),
        "texto_extraido": texto,
        "estabelecimento": None,
    }
    r = EXTRACTORS[categoria](it)
    r["caminho"] = caminho
    r["texto_extraido"] = texto
    return r


def extrair_lote(pasta, categoria):
    """Extrai todos os arquivos de uma pasta com o extractor da categoria."""
    resultados = []
    for f in sorted(os.listdir(pasta)):
        caminho = os.path.join(pasta, f)
        if os.path.isfile(caminho) and os.path.splitext(f)[1].lower() in (IMG_EXTS | {".pdf"}):
            print(f"  extraindo: {f}")
            resultados.append(extrair_um(caminho, categoria))
    return resultados


def mostrar_resultado(r):
    print("  " + "-" * 50)
    print(f"  Arquivo    : {r['arquivo']}")
    print(f"  Categoria  : {r['categoria']}")
    print(f"  Tipo       : {r.get('tipo')}")
    print(f"  Data       : {r.get('data')}")
    print(f"  Valor      : R$ {r.get('valor') if r.get('valor') is not None else '—'}")
    print(f"  Pagador    : {r.get('pagador')}")
    print(f"  Recebedor  : {r.get('recebedor')}")



def regenerar_csv():
    """Regera a fila global completa; usado somente no modo Extrair tudo."""
    from extrair_tudo import main as _gera
    _gera()
def salvar_fila_selecionada(resultados, categoria):
    """Gera comprovantes.json/csv somente com os resultados selecionados."""
    itens = []
    for ident, r in enumerate(resultados, 1):
        caminho = r.get("caminho") or r.get("arquivo") or ""
        data = r.get("data") or ""
        mes = data[:7] if data else ""
        try:
            arquivo_rel = os.path.relpath(caminho, BASE)
        except ValueError:
            arquivo_rel = caminho
        item = {
            "id": ident,
            "status": "pendente",
            "categoria": r.get("categoria") or categoria,
            "despesa": r.get("despesa") or classificar_despesa(categoria, r),
            "data": data,
            "mes": mes,
            "valor": r.get("valor"),
            "pagador": r.get("pagador") or "",
            "recebedor": r.get("recebedor") or "",
            "tipo": r.get("tipo") or "",
            "arquivo": arquivo_rel,
            "hash": md5_arquivo_absoluto(caminho),
            "periodicidade": r.get("periodicidade") or "Mensal",
        }
        itens.append(item)

    payload = {
        "meta": {
            "gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
            "base": BASE,
            "total_comprovantes": len(itens),
            "fonte": "selecao da ferramenta de extracao",
        },
        "comprovantes": itens,
    }
    with open(os.path.join(BASE, "comprovantes.json"), "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)

    campos = ["id", "categoria", "despesa", "data", "mes", "valor", "pagador",
              "recebedor", "tipo", "arquivo", "hash", "periodicidade"]
    with open(os.path.join(BASE, "comprovantes.csv"), "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=campos, delimiter=";", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(itens)
    print(f"Fila selecionada salva: {len(itens)} comprovante(s).")


def md5_arquivo_absoluto(caminho):
    try:
        with open(caminho, "rb") as fh:
            return hashlib.md5(fh.read()).hexdigest()
    except Exception:
        return ""


def salvar_json_lote(pasta, categoria, resultados):
    json_path, xlsx_path, normalizados = mesclar_indice_categoria(pasta, categoria, resultados)
    resultados[:] = normalizados
    print(f"\nJSON salvo: {json_path}")
    print(f"Planilha salva: {xlsx_path}")


def menu():
    print("=" * 55)
    print("  FERRAMENTA DE EXTRAÇÃO DE COMPROVANTES")
    print("=" * 55)
    print("  1. Extrair um arquivo")
    print("  2. Extrair lote (pasta de uma categoria)")
    print("  3. Extrair tudo (todas as categorias) e gerar CSV")
    print("  4. Sair")
    return input("  Escolha: ").strip()


def escolher_categoria():
    print("\n  Categorias disponíveis:")
    for i, c in enumerate(CATEGORIAS, 1):
        print(f"    {i}. {c}")
    try:
        n = int(input("  Número da categoria: ").strip())
        return CATEGORIAS[n - 1]
    except (ValueError, IndexError):
        print("  Categoria inválida.")
        return None


def modo_interativo():
    while True:
        op = menu()
        if op == "1":
            caminho = input("  Caminho do arquivo: ").strip()
            if not os.path.isfile(caminho):
                print("  Arquivo não encontrado.")
                continue
            cat = escolher_categoria()
            if cat is None:
                continue
            print()
            r = extrair_um(caminho, cat)
            mostrar_resultado(r)
            salvar_json_lote(os.path.dirname(caminho), cat, [r])
            salvar_fila_selecionada([r], cat)
        elif op == "2":
            pasta = input("  Caminho da pasta: ").strip()
            if not os.path.isdir(pasta):
                print("  Pasta não encontrada.")
                continue
            cat = escolher_categoria()
            if cat is None:
                continue
            print()
            resultados = extrair_lote(pasta, cat)
            print(f"\n  {len(resultados)} comprovantes extraídos.")
            salvar_json_lote(pasta, cat, resultados)
            salvar_fila_selecionada(resultados, cat)
        elif op == "3":
            print("\n  Extraindo tudo...")
            r = subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), "integrar_extracao_planilha.py")])
            if r.returncode != 0:
                raise SystemExit(r.returncode)
        elif op == "4":
            print("  Saindo.")
            break
        else:
            print("  Opção inválida.")
        print()


def main():
    ap = argparse.ArgumentParser(description="Ferramenta de extração de comprovantes")
    ap.add_argument("--arquivo", help="Caminho de um arquivo individual")
    ap.add_argument("--lote", help="Caminho de uma pasta (lote)")
    ap.add_argument("--categoria", help="Categoria (define o regex/extractor)")
    ap.add_argument("--tudo", action="store_true", help="Extrai todas as categorias e gera CSV")
    args = ap.parse_args()

    if args.arquivo:
        if not args.categoria:
            print("Erro: informe --categoria para escolher o regex.")
            sys.exit(1)
        r = extrair_um(args.arquivo, args.categoria)
        mostrar_resultado(r)
        salvar_json_lote(os.path.dirname(args.arquivo), args.categoria, [r])
        salvar_fila_selecionada([r], args.categoria)
        return

    if args.lote:
        if not args.categoria:
            print("Erro: informe --categoria para escolher o regex.")
            sys.exit(1)
        resultados = extrair_lote(args.lote, args.categoria)
        print(f"\n{len(resultados)} comprovantes extraídos.")
        salvar_json_lote(args.lote, args.categoria, resultados)
        salvar_fila_selecionada(resultados, args.categoria)
        return

    if args.tudo:
        r = subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), "integrar_extracao_planilha.py")])
        if r.returncode != 0:
            raise SystemExit(r.returncode)
        return

    modo_interativo()


if __name__ == "__main__":
    main()


