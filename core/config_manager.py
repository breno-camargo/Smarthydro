import os
import sys
import json
import base64
import logging
from datetime import datetime

DEFAULT_CONFIG = {
    "server": "WELLCARE-PC\\SQLEXPRESS",
    "database": "StruxureWareReportsDB",
    "trusted_connection": True,
    "db_user": "",
    "db_password": "",
    "default_m3_price": 63.68,
    "output_directory": "C:\\Relatorios_Hidrometros",
    "billing_cycle_type": "ciclo_29_28",
    "open_excel_after_generation": True,
    "sort_by_consumption": True,
    "odbc_driver": "ODBC Driver 17 for SQL Server",
    "recent_reports": []
}

def get_base_dir():
    """Retorna o diretório base da aplicação (suporta modo normal e PyInstaller .exe)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def get_config_path():
    """Retorna o caminho completo para o arquivo config.json."""
    return os.path.join(get_base_dir(), "config.json")


MESES_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

def format_report_filename(reference_date):
    """
    Retorna o nome padronizado do arquivo com base no mês de referência.
    Padrão solicitado: 'Rateio de água - Junho.xlsx'
    reference_date pode ser datetime, date ou string no formato 'YYYY-MM-DD...'
    """
    try:
        if isinstance(reference_date, str):
            clean = reference_date.strip().split(" ")[0]
            parts = clean.split("-")
            month = int(parts[1])
        else:
            month = reference_date.month
        mes_nome = MESES_PT[month - 1]
    except Exception:
        mes_nome = "Medicao"
    return f"Rateio de água - {mes_nome}.xlsx"


def _encode_password(plain_text):
    """Ofusca a senha em base64 para não ficar visível em texto puro no config.json."""
    if not plain_text:
        return ""
    return "b64:" + base64.b64encode(plain_text.encode("utf-8")).decode("ascii")


def _decode_password(stored_value):
    """Decodifica a senha ofuscada. Aceita texto puro para retrocompatibilidade."""
    if not stored_value:
        return ""
    if stored_value.startswith("b64:"):
        try:
            return base64.b64decode(stored_value[4:]).decode("utf-8")
        except Exception:
            return stored_value
    # Retrocompatibilidade: senha antiga salva em texto puro
    return stored_value


def load_config():
    """Carrega as configurações a partir do arquivo config.json ou retorna os padrões."""
    cfg_path = get_config_path()
    config = DEFAULT_CONFIG.copy()

    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                saved = json.load(f)
                config.update(saved)
        except Exception as e:
            logging.error(f"Erro ao ler {cfg_path}: {e}. Utilizando configurações padrão.")
    else:
        save_config(config)

    # Decodificar senha ofuscada para uso em memória
    config["db_password"] = _decode_password(config.get("db_password", ""))

    # Garantir que a pasta de saída padrão exista
    out_dir = config.get("output_directory")
    if out_dir:
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as e:
            logging.warning(f"Não foi possível criar a pasta de saída {out_dir}: {e}")

    return config

def save_config(new_config):
    """Salva o dicionário de configurações no arquivo config.json (ofuscando a senha)."""
    cfg_path = get_config_path()
    try:
        # Criar cópia para não modificar o dict em memória
        config_to_save = new_config.copy()
        # Ofuscar senha antes de gravar no disco
        raw_pwd = config_to_save.get("db_password", "")
        if raw_pwd and not raw_pwd.startswith("b64:"):
            config_to_save["db_password"] = _encode_password(raw_pwd)

        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(config_to_save, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        logging.error(f"Falha ao salvar configurações em {cfg_path}: {e}")
        return False


def get_recent_reports():
    """
    Retorna a lista dos relatórios presentes na pasta de saída,
    enriquecidos com metadados do histórico quando disponíveis.
    Sempre escaneia a pasta para refletir mudanças (renomear, deletar, adicionar).
    """
    cfg = load_config()
    history = cfg.get("recent_reports", [])
    out_dir = cfg.get("output_directory", "")

    # Índice do histórico por caminho absoluto para consulta rápida de metadados
    hist_by_path = {}
    for item in history:
        if isinstance(item, dict):
            p = item.get("path", "")
            if p:
                hist_by_path[os.path.normcase(os.path.abspath(p))] = item

    result = []

    # Escaneia a pasta de saída para pegar TODOS os arquivos .xlsx
    if out_dir and os.path.exists(out_dir):
        try:
            files = [
                os.path.join(out_dir, f) for f in os.listdir(out_dir)
                if f.lower().endswith(".xlsx")
            ]
            # Ordenar pelos mais recentes (data de modificação)
            files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

            for f_path in files[:5]:
                abs_path = os.path.abspath(f_path)
                key = os.path.normcase(abs_path)
                mtime = os.path.getmtime(f_path)
                dt_str = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")

                # Se tiver metadados do histórico, usa; senão cria entrada básica
                if key in hist_by_path:
                    entry = hist_by_path[key].copy()
                    # Atualiza nome caso o arquivo tenha sido renomeado
                    entry["filename"] = os.path.basename(f_path)
                    entry["path"] = abs_path
                else:
                    entry = {
                        "filename": os.path.basename(f_path),
                        "path": abs_path,
                        "periodo": "",
                        "gerado_em": dt_str,
                        "total_m3": None,
                        "total_rs": None
                    }
                result.append(entry)
        except Exception:
            pass

    # Fallback: se a pasta não existe ou está vazia, usa histórico do config
    if not result:
        for item in history:
            if isinstance(item, dict):
                p = item.get("path", "")
                if p and os.path.exists(p):
                    result.append(item)

    return result[:5]


def add_recent_report(path, periodo=None, total_m3=None, total_rs=None):
    """
    Registra um novo relatório no topo do histórico (mantém até 5).
    """
    cfg = load_config()
    history = cfg.get("recent_reports", [])

    abs_path = os.path.abspath(path)
    filename = os.path.basename(abs_path)
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Remover ocorrência prévia se houver
    history = [h for h in history if isinstance(h, dict) and os.path.abspath(h.get("path", "")) != abs_path]

    new_item = {
        "filename": filename,
        "path": abs_path,
        "periodo": periodo or "",
        "gerado_em": now_str,
        "total_m3": round(float(total_m3), 2) if total_m3 is not None else None,
        "total_rs": round(float(total_rs), 2) if total_rs is not None else None,
    }
    history.insert(0, new_item)
    cfg["recent_reports"] = history[:5]
    save_config(cfg)
    return cfg["recent_reports"]
