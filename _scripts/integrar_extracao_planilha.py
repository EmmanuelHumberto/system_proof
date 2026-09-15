#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pipeline de extração e enfileiramento.

Este script NAO preenche a planilha. Ele reprocessa os comprovantes por
categoria, atualiza os dados_extraidos.json e entao gera a fila
comprovantes.json para o Excel.
"""
import os
import subprocess
import sys

from config import CONFIG, COMPROVANTES_DIR
from extrair_comprovantes import main as reprocessar_arquivos, listar_arquivos
from extrair_tudo import main as gerar_json_csv
from inventario_comprovantes import main as gerar_inventario


def main():
    if not os.path.isdir(COMPROVANTES_DIR) or not listar_arquivos():
        raise SystemExit("Nenhum comprovante original encontrado. Restaure a pasta comprovantes antes de extrair; os indices existentes foram preservados.")
    base_scripts = os.path.dirname(os.path.abspath(__file__))
    extrator = os.path.join(base_scripts, "extrair_categoria.py")
    categorias = sys.argv[1:] or list(CONFIG.keys())

    print("Inventariando arquivos fisicos em comprovantes/...")
    gerar_inventario()

    print("Reprocessando arquivos fisicos em comprovantes/...")
    reprocessar_arquivos()

    print()
    print("Atualizando JSON e planilha de cada categoria...")
    for categoria in categorias:
        if categoria not in CONFIG:
            raise SystemExit(f"Categoria desconhecida: {categoria}")
        print()
        print(f"== {categoria} ==")
        r = subprocess.run([sys.executable, extrator, categoria])
        if r.returncode != 0:
            raise SystemExit(r.returncode)

    print()
    print("Gerando fila consolidada...")
    gerar_json_csv()
    print()
    print("Fila pronta. Abra o Excel e carregue/importe a fila.")


if __name__ == "__main__":
    main()
