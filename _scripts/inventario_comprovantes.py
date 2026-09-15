#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cria o inventario completo dos comprovantes fisicos.

O inventario nao depende da tela de extracao nem da planilha. Ele percorre
recursivamente ``comprovantes/``, calcula hash, localiza a extracao existente
quando houver e mostra o que ainda nao entrou em nenhum indice. Assim ele e a
primeira conferencia antes de importar uma fila para o Excel.

Saidas em ``_extraidos/``:
  - inventario_comprovantes.json: todos os arquivos e resumo por pasta
  - inventario_comprovantes.csv: lista auditavel dos arquivos
"""
import csv
import datetime as dt
import hashlib
import json
import os
from collections import Counter, defaultdict

from config import BASE, COMPROVANTES_DIR, EXTRAIDOS
from indices_categoria import COMPROVANTE_EXTS, rel_base


def hash_arquivo(caminho):
    digest = hashlib.md5()
    with open(caminho, "rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
            digest.update(bloco)
    return digest.hexdigest()


def carregar_indices_existentes():
    """Indexa todos os dados_extraidos.json por caminho e hash, sem confiar no total salvo."""
    por_caminho = {}
    por_hash = defaultdict(list)
    if not os.path.isdir(COMPROVANTES_DIR):
        return por_caminho, por_hash

    for raiz, _dirs, arquivos in os.walk(COMPROVANTES_DIR):
        if "dados_extraidos.json" not in arquivos:
            continue
        caminho_json = os.path.join(raiz, "dados_extraidos.json")
        try:
            with open(caminho_json, encoding="utf-8") as arquivo:
                itens = json.load(arquivo).get("comprovantes", [])
        except (OSError, json.JSONDecodeError) as erro:
            print(f"Indice ignorado ({caminho_json}): {erro}")
            continue
        for item in itens:
            registro = dict(item)
            relativo = registro.get("caminho_rel") or rel_base(
                os.path.join(raiz, registro.get("arquivo", ""))
            )
            relativo = relativo.replace("/", "\\")
            registro["caminho_rel"] = relativo
            registro["indice_origem"] = rel_base(caminho_json)
            por_caminho[relativo] = registro
            if registro.get("hash"):
                por_hash[registro["hash"]].append(registro)
    return por_caminho, por_hash


def pasta_logica(relativo):
    partes = relativo.replace("/", "\\").split("\\")
    # comprovantes\\<zip-origem>\\<categoria>\\...\\arquivo
    return "\\".join(partes[:-1]) if len(partes) > 1 else "(raiz)"


def main():
    os.makedirs(EXTRAIDOS, exist_ok=True)
    por_caminho, por_hash = carregar_indices_existentes()
    registros = []

    for raiz, _dirs, arquivos in os.walk(COMPROVANTES_DIR):
        for nome in sorted(arquivos):
            extensao = os.path.splitext(nome)[1].lower()
            if extensao not in COMPROVANTE_EXTS or nome.startswith(".~lock"):
                continue
            absoluto = os.path.join(raiz, nome)
            relativo = rel_base(absoluto).replace("/", "\\")
            digest = hash_arquivo(absoluto)
            indice = por_caminho.get(relativo)
            if indice is None and len(por_hash.get(digest, [])) == 1:
                indice = por_hash[digest][0]

            valor = indice.get("valor") if indice else None
            registros.append({
                "caminho_rel": relativo,
                "pasta": pasta_logica(relativo),
                "arquivo": nome,
                "extensao": extensao.lstrip("."),
                "tamanho_bytes": os.path.getsize(absoluto),
                "hash": digest,
                "indexado": indice is not None,
                "indice_origem": indice.get("indice_origem", "") if indice else "",
                "categoria": indice.get("categoria", "") if indice else "",
                "despesa": indice.get("despesa", "") if indice else "",
                "data": indice.get("data", "") if indice else "",
                "valor": valor,
                "valor_extraido": valor is not None,
            })

    por_pasta = {}
    for pasta in sorted({r["pasta"] for r in registros}):
        itens = [r for r in registros if r["pasta"] == pasta]
        valores = [r["valor"] for r in itens if isinstance(r["valor"], (int, float))]
        por_pasta[pasta] = {
            "total_arquivos": len(itens),
            "indexados": sum(r["indexado"] for r in itens),
            "sem_indice": sum(not r["indexado"] for r in itens),
            "com_valor_extraido": len(valores),
            "sem_valor_extraido": len(itens) - len(valores),
            "soma_valores_extraidos": round(sum(valores), 2),
        }

    hashes = Counter(r["hash"] for r in registros)
    duplicados = [
        {"hash": digest, "quantidade": quantidade,
         "arquivos": [r["caminho_rel"] for r in registros if r["hash"] == digest]}
        for digest, quantidade in hashes.items() if quantidade > 1
    ]
    valores = [r["valor"] for r in registros if isinstance(r["valor"], (int, float))]
    payload = {
        "meta": {
            "gerado_em": dt.datetime.now().isoformat(timespec="seconds"),
            "fonte": "varredura recursiva de comprovantes fisicos",
            "base": BASE,
            "total_arquivos": len(registros),
            "indexados": sum(r["indexado"] for r in registros),
            "sem_indice": sum(not r["indexado"] for r in registros),
            "com_valor_extraido": len(valores),
            "sem_valor_extraido": len(registros) - len(valores),
            "soma_valores_extraidos": round(sum(valores), 2),
            "hashes_repetidos": len(duplicados),
        },
        "pastas": por_pasta,
        "duplicidades_por_hash": duplicados,
        "comprovantes": registros,
    }
    caminho_json = os.path.join(EXTRAIDOS, "inventario_comprovantes.json")
    caminho_csv = os.path.join(EXTRAIDOS, "inventario_comprovantes.csv")
    with open(caminho_json, "w", encoding="utf-8") as arquivo:
        json.dump(payload, arquivo, ensure_ascii=False, indent=2)
    with open(caminho_csv, "w", newline="", encoding="utf-8") as arquivo:
        campos = ["caminho_rel", "pasta", "arquivo", "extensao", "tamanho_bytes", "hash",
                  "indexado", "indice_origem", "categoria", "despesa", "data", "valor", "valor_extraido"]
        escritor = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
        escritor.writeheader()
        escritor.writerows(registros)
    print(f"Inventario: {caminho_json}")
    print(f"Arquivos fisicos: {len(registros)} | indexados: {payload['meta']['indexados']} | sem indice: {payload['meta']['sem_indice']}")
    print(f"Com valor: {len(valores)} | soma extraida: R$ {sum(valores):.2f}")
    print(f"Hashes repetidos: {len(duplicados)}")


if __name__ == "__main__":
    main()
