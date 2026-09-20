#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Ponto de entrada do Software de Automação de Relatórios de Hidrômetros.
Condomínio Praça Pamplona - StruxureWare / EBO.

Modos de uso:
1. Gráfico (padrão): Duplo clique no executável ou 'python main.py'
2. Automático: 'python main.py --auto' (ideal para agendamento no Windows)
3. Linha de comando com parâmetros: 'python main.py --inicio 01/08/2026 --fim 31/08/2026 --valor 63.68'
"""

import sys
import os

# Adiciona o diretório base ao sys.path para garantir imports corretos
base_dir = os.path.dirname(os.path.abspath(__file__))
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

def main():
    # Se houver argumentos de linha de comando (além do nome do script), roda CLI
    if len(sys.argv) > 1:
        from cli.runner import run_cli
        run_cli()
    else:
        from gui.app_window import start_gui
        start_gui()

if __name__ == "__main__":
    main()
