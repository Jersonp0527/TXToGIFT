#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
python -m pip install --upgrade pip
python -m pip install -e .
echo ""
echo "Instalado. Ahora puedes ejecutar: txt-to-gift archivo.docx [salida.txt]"
