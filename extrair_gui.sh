#!/bin/bash
# Atalho para iniciar a aplicação de extração de comprovantes
cd "$(dirname "$0")/_scripts" || exit 1
exec python3 extrair_gui.py
