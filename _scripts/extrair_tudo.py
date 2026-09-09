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
from config import BASE, COMPROVANTES_DIR, CONFIG  # noqa: E402
from indices_categoria import normalizar_resultado  # noqa: E402


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
    for cat, cfg in CONFIG.items():
        sub = os.path.join(COMPROVANTES_DIR, cfg["pasta_top"], cfg["sub"])
        jp = os.path.join(sub, "dados_extraidos.json")
        if not os.path.exists(jp):
            continue
        with open(jp, encoding="utf-8") as fh:
            comps = json.load(fh)["comprovantes"]
        for c in comps:
            ident += 1
            c = normalizar_resultado(c, cat, os.path.join(sub, c.get("arquivo", "")))
            desp = c.get("despesa") or classificar_despesa(cat, c)
            data = c.get("data") or ""
            mes = data[:7] if data else ""
            valor = f"{c['valor']:.2f}" if c.get("valor") is not None else ""
            # heurística de periodicidade (editável na interface VBA)
            if any(k in desp.lower() for k in ("iptu", "enem", "matrícula", "matricula")):
                periodicidade = "Anual"
            else:
                periodicidade = "Mensal"
            arquivo_rel = c.get("caminho_rel") or os.path.join("comprovantes", cfg["pasta_top"], cfg["sub"], c["arquivo"])
            hash_arquivo = c.get("hash") or md5_arquivo(arquivo_rel)
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
            "fonte": "dados_extraidos.json por categoria",
        },
        "comprovantes": itens_json,
    }
    with open(saida_json, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)

    print("JSON gerado:", saida_json)
    print("CSV gerado:", saida)
    print("Total de comprovantes:", ident)
    # resumo por categoria
    from collections import Counter
    cnt = Counter()
    for ln in linhas:
        cnt[ln.split(";")[1]] += 1
    for cat, n in sorted(cnt.items()):
        print(f"  {cat:26} | {n}")


if __name__ == "__main__":
    main()
