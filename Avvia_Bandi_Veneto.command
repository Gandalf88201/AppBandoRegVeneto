#!/bin/bash
cd "$(dirname "$0")"
python3 avvia.py "$@"
echo
read -p "Premi Invio per chiudere..."
