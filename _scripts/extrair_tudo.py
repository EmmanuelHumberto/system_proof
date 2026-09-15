#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera o comprovantes.json e o comprovantes.csv a partir dos dados_extraidos.json
de cada pasta.

Saídas:
  Daisy/comprovantes.json
  Daisy/comprovantes.csv

Colunas (separadas por ";"):
  id;categoria;despesa;data;mes;valor;pagador;recebedor;tipo;arquivo;hash;periodicidade
"""
import os
import json
import hashlib
import sys
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from classificacao import classificar_despesa  # noqa: E402
from config import BASE, COMPROVANTES_DIR  # noqa: E402
from indices_categoria import COMPROVANTE_EXTS, normalizar_resultado  # noqa: E402
from inventario_comprovantes import main as gerar_inventario  # noqa: E402


def md5_arquivo(rel):
    p = os.path.join(BASE, rel)
    try:
        with open(p, "rb") as fh:
            return hashlib.md5(fh.read()).hexdigest()
    except Exception:
        return ""


def main():
    linhas = []
    itens_json = []
    ident = 0
    caminhos_incluidos = set()
    hashes_incluidos = set()
    duplicados_ignorados = []
    indices = []
    for raiz, _pastas, arquivos in os.walk(COMPROVANTES_DIR):
        if "dados_extraidos.json" in arquivos:
            indices.append(os.path.join(raiz, "dados_extraidos.json"))

    for jp in sorted(indices):
        with open(jp, encoding="utf-8") as fh:
            indice = json.load(fh)
        cat_indice = indice.get("categoria", "")
        comps = indice.get("comprovantes", [])
        for c in comps:
            cat = c.get("categoria") or cat_indice
            if not cat:
                continue
            sub = os.path.dirname(jp)
            c = normalizar_resultado(c, cat, os.path.join(sub, c.get("arquivo", "")))
            arquivo_rel = c.get("caminho_rel") or os.path.relpath(
                os.path.join(sub, c.get("arquivo", "")), BASE
            )
            arquivo_abs = os.path.join(BASE, arquivo_rel)
            if not os.path.isfile(arquivo_abs):
                continue
            if os.path.splitext(arquivo_abs)[1].lower() not in COMPROVANTE_EXTS:
                continue
            chave_caminho = os.path.normcase(os.path.normpath(arquivo_rel))
            if chave_caminho in caminhos_incluidos:
                continue
            caminhos_incluidos.add(chave_caminho)
            ident += 1
            desp = c.get("despesa") or classificar_despesa(cat, c)
            data = c.get("data") or ""
            mes = data[:7] if data else ""
            valor = f"{c['valor']:.2f}" if c.get("valor") is not None else ""
            # heurística de periodicidade (editável na interface VBA)
            if any(k in desp.lower() for k in ("iptu", "enem", "matrícula", "matricula")):
                periodicidade = "Anual"
            else:
                periodicidade = "Mensal"
            hash_arquivo = c.get("hash") or md5_arquivo(arquivo_rel)
            if hash_arquivo and hash_arquivo in hashes_incluidos:
                duplicados_ignorados.append({
                    "arquivo": arquivo_rel,
                    "hash": hash_arquivo,
                    "motivo": "mesmo conteudo ja presente na fila consolidada",
                })
                continue
            if hash_arquivo:
                hashes_incluidos.add(hash_arquivo)
            item = {
                "id": ident,
                "status": "pendente",
                "categoria": cat,
                "despesa": desp,
                "data": data,
                "mes": mes,
                "valor": c.get("valor"),
                "pagador": c.get("pagador") or "",
                "recebedor": c.get("recebedor") or "",
                "tipo": c.get("tipo") or "",
                "arquivo": arquivo_rel,
                "hash": hash_arquivo,
                "periodicidade": periodicidade,
            }
            itens_json.append(item)
            linhas.append(";".join([
                str(ident),
                cat,
                desp,
                data,
                mes,
                valor,
                c.get("pagador") or "",
                c.get("recebedor") or "",
                c.get("tipo") or "",
                arquivo_rel,
                hash_arquivo,
                periodicidade,
            ]))

    cab = "id;categoria;despesa;data;mes;valor;pagador;recebedor;tipo;arquivo;hash;periodicidade"
    saida = os.path.join(BASE, "comprovantes.csv")
    with open(saida, "w", encoding="utf-8") as fh:
        fh.write(cab + "\n")
        fh.write("\n".join(linhas) + "\n")

    saida_json = os.path.join(BASE, "comprovantes.json")
    payload = {
        "meta": {
            "gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
            "base": BASE,
            "total_comprovantes": len(itens_json),
            "arquivos_ignorados_por_hash_repetido": len(duplicados_ignorados),
            "fonte": "todos os dados_extraidos.json validos contra arquivos fisicos",
        },
        "comprovantes": itens_json,
        "ignorados_por_hash_repetido": duplicados_ignorados,
    }
    with open(saida_json, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)

    gerar_inventario()

    print("JSON gerado:", saida_json)
    print("CSV gerado:", saida)
    print("Total de comprovantes:", len(itens_json))
    print("Copias identicas ignoradas:", len(duplicados_ignorados))
    # resumo por categoria
    from collections import Counter
    cnt = Counter()
    for ln in linhas:
        cnt[ln.split(";")[1]] += 1
    for cat, n in sorted(cnt.items()):
        print(f"  {cat:26} | {n}")


if __name__ == "__main__":
    main()
