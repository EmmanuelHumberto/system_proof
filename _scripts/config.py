#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configuração central do projeto Daisy — caminhos e mapeamento de pastas.

O caminho base (pasta Daisy) é detectado automaticamente: este arquivo fica em
Daisy/_scripts/, então a base é o diretório pai (Daisy/). Assim o projeto pode
ser copiado para qualquer máquina sem editar nada.

Para sobrescrever o caminho (caso raro), defina a variável de ambiente DAISY_BASE:
    export DAISY_BASE=/outro/caminho
"""
import os

# Pasta base do projeto (Daisy), detectada automaticamente.
_BASE_DETECTADA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.environ.get("DAISY_BASE", _BASE_DETECTADA)

# Arquivos principais
PLANILHA = os.path.join(BASE, "Planilha_Estrategica_Pensao_Alimenticia_Adolescente.xlsx")
PLANILHA_VBA = os.path.join(BASE, "Planilha_Comprovantes_VBA.xlsm")
CSV = os.path.join(BASE, "comprovantes.csv")
EXTRAIDOS = os.path.join(BASE, "_extraidos")
DADOS = os.path.join(EXTRAIDOS, "dados.json")
COMPROVANTES_DIR = os.path.join(BASE, "comprovantes")

# Mapeamento: categoria -> pasta top-level + subpasta com os comprovantes.
CONFIG = {
    "Alimentação": {"pasta_top": "Alimentação-20260907T002627Z-1-001", "sub": "Alimentação"},
    "Comunicação e tecnologia": {"pasta_top": "Comunicação e tecnologia-20260907T002634Z-1-001", "sub": "Comunicação e tecnologia"},
    "Lazer e convivência": {"pasta_top": "Lazer e convivência-20260907T002710Z-1-001", "sub": "Lazer e convivência"},
    "Moradia": {"pasta_top": "Moradia-20260907T002721Z-1-001", "sub": "Moradia"},
    "Saúde": {"pasta_top": "Saúde-20260907T002728Z-1-001", "sub": "Saúde"},
    "Transporte": {"pasta_top": "Transporte-20260907T002744Z-1-001", "sub": "Transporte"},
    "Educação": {"pasta_top": "Educação-20260907T002643Z-1-001", "sub": "Educação"},
    "Vestuário e higiene": {"pasta_top": "Vestuário e higiene -20260907T002759Z-1-001", "sub": "Vestu__rio e higiene"},
    "Esporte e desenvolvimento": {"pasta_top": "Esporte e desenvolvimento", "sub": "Esporte e desenvolvimento"},
    "Mesada": {"pasta_top": "Mesada", "sub": "Mesada"},
}

# Nomes dos meses (formatação de competência)
MES_NOME = {
    "01": "janeiro", "02": "fevereiro", "03": "março", "04": "abril",
    "05": "maio", "06": "junho", "07": "julho", "08": "agosto",
    "09": "setembro", "10": "outubro", "11": "novembro", "12": "dezembro",
}
MES_NUM = {v: k for k, v in MES_NOME.items()}

