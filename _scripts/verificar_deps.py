#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Verifica as dependências do projeto Daisy e instala as bibliotecas Python.

Uso:
  python verificar_deps.py             # apenas verifica e reporta
  python verificar_deps.py --instalar  # verifica e instala as libs Python faltantes
"""
import importlib
import shutil
import subprocess
import sys
import argparse

# bibliotecas Python necessárias (nome de import -> nome no pip)
LIBS = {
    "openpyxl": "openpyxl",
    "pdfplumber": "pdfplumber",
    "PIL": "pillow",
    "tkinter": None,  # não é instalado via pip (pacote do sistema)
}

# ferramentas externas (binário -> descrição / como instalar)
BINARIOS = {
    "tesseract": "OCR de imagens",
    "pdftoppm": "conversão de PDF escaneado (poppler-utils)",
}


def verificar():
    print("=" * 55)
    print("  VERIFICAÇÃO DE DEPENDÊNCIAS — PROJETO DAISY")
    print("=" * 55)

    print("\n[1] Bibliotecas Python")
    faltam_libs = []
    for lib, pip_nome in LIBS.items():
        try:
            importlib.import_module(lib)
            print(f"    OK     {lib}")
        except Exception:
            print(f"    FALTA  {lib}" + (f"  (pip install {pip_nome})" if pip_nome else ""))
            faltam_libs.append(lib)

    print("\n[2] Ferramentas externas")
    faltam_bin = []
    for bin_, desc in BINARIOS.items():
        p = shutil.which(bin_)
        if p:
            print(f"    OK     {bin_}  ({desc})")
        else:
            print(f"    FALTA  {bin_}  ({desc})")
            faltam_bin.append(bin_)

    # idioma de OCR (português) do tesseract
    if shutil.which("tesseract"):
        r = subprocess.run(["tesseract", "--list-langs"], capture_output=True, text=True)
        if "por" not in r.stdout:
            print("    ATENÇÃO  tesseract sem o idioma 'por' (português)")
            faltam_bin.append("tesseract-por")

    return faltam_libs, faltam_bin


def instalar_libs(faltam_libs):
    pacotes = [LIBS[l] for l in faltam_libs if LIBS.get(l)]
    if pacotes:
        print(f"\nInstalando bibliotecas: {' '.join(pacotes)} ...")
        subprocess.run([sys.executable, "-m", "pip", "install", *pacotes])
    if "tkinter" in faltam_libs:
        print("\ntkinter não é instalado via pip. Instale pelo sistema:")
        print("  Linux: sudo apt install python3-tk")
        print("  macOS: (já vem com o Python do python.org)")
        print("  Windows: (já vem com o instalador do python.org)")


def instrucoes_binarios(faltam_bin):
    print("\nFerramentas externas a instalar (conforme seu sistema):")
    print("  Linux (Debian/Ubuntu):")
    print("    sudo apt install tesseract-ocr tesseract-ocr-por poppler-utils")
    print("  macOS (Homebrew):")
    print("    brew install tesseract tesseract-lang poppler")
    print("  Windows:")
    print("    - Tesseract: https://github.com/UB-Mannheim/tesseract/wiki")
    print("      (marcar o idioma 'Portuguese' na instalação)")
    print("    - Poppler:  https://github.com/oschwartz10612/poppler-windows/releases")


def main():
    ap = argparse.ArgumentParser(description="Verifica dependências do projeto Daisy")
    ap.add_argument("--instalar", action="store_true",
                    help="Instala as bibliotecas Python faltantes")
    args = ap.parse_args()

    faltam_libs, faltam_bin = verificar()

    if args.instalar and faltam_libs:
        instalar_libs(faltam_libs)
        print()
        faltam_libs, faltam_bin = verificar()

    print("\n" + "=" * 55)
    if not faltam_libs and not faltam_bin:
        print("  RESULTADO: tudo OK — nenhuma dependência faltando.")
        return 0
    print("  RESULTADO: há dependências faltando.")
    if faltam_libs:
        print(f"    Bibliotecas Python: {', '.join(faltam_libs)}")
        print(f"    Para instalar: python verificar_deps.py --instalar")
    if faltam_bin:
        print(f"    Ferramentas externas: {', '.join(faltam_bin)}")
        instrucoes_binarios(faltam_bin)
    return 1


if __name__ == "__main__":
    sys.exit(main())
