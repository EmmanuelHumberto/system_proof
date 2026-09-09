#!/bin/bash
# Atalho para verificar (e opcionalmente instalar) as dependências do projeto Daisy.
cd "$(dirname "$0")/_scripts" || exit 1
python3 verificar_deps.py
echo ""
read -p "Instalar as bibliotecas Python que faltam? (s/N): " resp
if [ "$resp" = "s" ] || [ "$resp" = "S" ]; then
    python3 verificar_deps.py --instalar
fi
echo ""
read -p "Pressione Enter para fechar..."
