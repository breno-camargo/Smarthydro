import os
import sys
import json
import base64
import logging
import re
import shutil
from datetime import datetime

DEFAULT_OPERATORS = [
    {
        "id": "breno",
        "name": "Breno Camargo",
        "role": "Técnico de Sistemas Prediais",
        "email": "breno.camargo@compasss.com.br",
        "phone": "+55 11 99012 7316",
        "whatsapp_phone": "+55 11 99012 7316",
        "whatsapp_apikey": "1275019",
        "smtp_user": "breno.camargo@compasss.com.br",
        "smtp_password": "",
        "is_default": True
    }
]

DEFAULT_CONDOMINIO_EMAIL_TO = "gerente.pamplona@zangari.com.br"
DEFAULT_CONDOMINIO_EMAIL_CC = (
    "assistente.pamplona2@zangari.com.br; "
    "almir@zangari.com.br; "
    "gabriel.domingos@compasss.com.br; "
    "victor.carvalho@compasss.com.br; "
    "breno.camargo@compasss.com.br"
)


def get_default_condominio_emails():
    """Retorna os e-mails oficiais de destinatários do Condomínio Praça Pamplona (Para, Cc)."""
    return DEFAULT_CONDOMINIO_EMAIL_TO, DEFAULT_CONDOMINIO_EMAIL_CC


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
    "email_recipients": "",
    "email_cc": "",
    "email_send_mode": "smtp",
    "smtp_server": "smtps.uhserver.com",
    "smtp_port": 465,
    "smtp_use_tls": False,
    "smtp_use_ssl": True,
    "smtp_user": "",
    "smtp_password": "",
    "recent_reports": [],
    "active_operator_id": "breno",
    "operators": DEFAULT_OPERATORS,
    "webhook_enabled": False,
    "webhook_platform": "whatsapp",
    "webhook_whatsapp_phone": "",
    "webhook_whatsapp_apikey": "",
    "webhook_url": "",
    "webhook_telegram_token": "",
    "webhook_telegram_chat_id": "",
    "webhook_notify_scheduled": True,
    "webhook_notify_anomalies": True
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


def get_report_year_month(reference_date):
    """
    Retorna uma tupla (ano_str, pasta_mes) com base na data de referência.
    Padrão solicitado: '10.26' (Mês.Ano com 2 dígitos cada) para ordenação cronológica correta.
    Ex: ('2026', '10.26')
    """
    try:
        if isinstance(reference_date, str):
            clean = reference_date.strip().split(" ")[0]
            parts = clean.split("-")
            year = int(parts[0])
            month = int(parts[1])
        else:
            year = reference_date.year
            month = reference_date.month
        year_short = str(year)[-2:]
        folder_name = f"{month:02d}.{year_short}"
        return str(year), folder_name
    except Exception:
        now = datetime.now()
        return str(now.year), f"{now.month:02d}.{str(now.year)[-2:]}"


def get_report_output_folder(base_out_dir, reference_date):
    """
    Retorna o caminho da subpasta organizada por Ano/Mês dentro da pasta de relatórios.
    Cria a pasta automaticamente se não existir.
    Ex: C:\...\Relatorios\2026\10.26
    """
    year, folder_name = get_report_year_month(reference_date)
    target_dir = os.path.join(base_out_dir, year, folder_name)
    os.makedirs(target_dir, exist_ok=True)
    return target_dir


def organize_loose_reports(base_out_dir):
    """
    Organiza arquivos soltos na raiz e migra pastas antigas com nome de mês por extenso
    (ex: 2026\Outubro -> 2026\10.26) para manter a ordem cronológica correta no Windows Explorer.
    """
    if not base_out_dir or not os.path.exists(base_out_dir):
        return

    try:
        # Fast-check: se não houver arquivos soltos .xlsx/.pdf na raiz, retorna imediatamente sem I/O pesado
        entries = os.listdir(base_out_dir)
        has_loose = any(
            (f.lower().endswith((".xlsx", ".pdf")) and not f.startswith("~$"))
            for f in entries
            if os.path.isfile(os.path.join(base_out_dir, f))
        )
        if not has_loose:
            return

        # 1. Migrar pastas antigas que usavam nome de mês em texto (ex: 2026\Outubro -> 2026\10.26)
        for year_item in entries:
            year_path = os.path.join(base_out_dir, year_item)
            if os.path.isdir(year_path) and re.match(r'^\d{4}$', year_item):
                year_short = year_item[-2:]
                for sub in os.listdir(year_path):
                    sub_path = os.path.join(year_path, sub)
                    if os.path.isdir(sub_path):
                        for idx, m_nome in enumerate(MESES_PT, 1):
                            if sub.lower() == m_nome.lower():
                                new_sub_name = f"{idx:02d}.{year_short}"
                                new_sub_path = os.path.join(year_path, new_sub_name)
                                os.makedirs(new_sub_path, exist_ok=True)
                                for f in os.listdir(sub_path):
                                    src_f = os.path.join(sub_path, f)
                                    dst_f = os.path.join(new_sub_path, f)
                                    if not os.path.exists(dst_f):
                                        shutil.move(src_f, dst_f)
                                    elif os.path.abspath(src_f) != os.path.abspath(dst_f):
                                        try:
                                            os.remove(src_f)
                                        except Exception:
                                            pass
                                try:
                                    os.rmdir(sub_path)
                                except Exception:
                                    pass
                                break

        # 2. Organizar arquivos soltos na raiz da pasta base
        for item in os.listdir(base_out_dir):
            item_path = os.path.join(base_out_dir, item)
            if os.path.isfile(item_path) and not item.startswith("~$"):
                item_lower = item.lower()
                if item_lower.endswith(".xlsx") or item_lower.endswith(".pdf"):
                    found_month_idx = None
                    for idx, m in enumerate(MESES_PT, 1):
                        if m.lower() in item_lower:
                            found_month_idx = idx
                            break

                    year_match = re.search(r'\b(202\d)\b', item)
                    if year_match:
                        found_year = year_match.group(1)
                    else:
                        mtime = os.path.getmtime(item_path)
                        found_year = str(datetime.fromtimestamp(mtime).year)

                    if found_month_idx:
                        year_short = found_year[-2:]
                        folder_name = f"{found_month_idx:02d}.{year_short}"
                        target_folder = os.path.join(base_out_dir, found_year, folder_name)
                        os.makedirs(target_folder, exist_ok=True)
                        dest_path = os.path.join(target_folder, item)
                        if os.path.abspath(item_path) != os.path.abspath(dest_path):
                            if not os.path.exists(dest_path):
                                shutil.move(item_path, dest_path)
                            else:
                                try:
                                    os.remove(item_path)
                                except Exception:
                                    pass
    except Exception as e:
        logging.warning(f"Erro ao organizar relatórios em {base_out_dir}: {e}")



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

decode_password = _decode_password
encode_password = _encode_password


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

    # Decodificar senhas ofuscadas para uso em memória
    config["db_password"] = _decode_password(config.get("db_password", ""))
    config["smtp_password"] = _decode_password(config.get("smtp_password", ""))

    # Garantir que operadores existam e senhas estejam decodificadas
    if "operators" in config and isinstance(config["operators"], list) and len(config["operators"]) > 0:
        for op in config["operators"]:
            if isinstance(op, dict):
                op["smtp_password"] = _decode_password(op.get("smtp_password", ""))
    else:
        # Migração automática a partir do usuário atual do SMTP
        smtp_u = config.get("smtp_user", "breno.camargo@compasss.com.br")
        config["operators"] = [
            {
                "id": "breno",
                "name": "Breno Camargo",
                "role": "Técnico de Sistemas Prediais",
                "email": smtp_u or "breno.camargo@compasss.com.br",
                "phone": "+55 11 99012 7316",
                "smtp_user": smtp_u or "breno.camargo@compasss.com.br",
                "smtp_password": "",
                "is_default": True
            }
        ]
        config["active_operator_id"] = "breno"

    return config

def save_config(new_config):
    """Salva o dicionário de configurações no arquivo config.json (ofuscando senhas e mantendo backup)."""
    cfg_path = get_config_path()
    try:
        # Backup de segurança antes de sobrescrever
        if os.path.exists(cfg_path):
            bak_path = os.path.join(os.path.dirname(cfg_path), "config.backup.json")
            try:
                shutil.copy2(cfg_path, bak_path)
            except Exception:
                pass

        # Criar cópia para não modificar o dict em memória
        config_to_save = new_config.copy()
        # Ofuscar senhas antes de gravar no disco
        raw_pwd = config_to_save.get("db_password", "")
        if raw_pwd and not raw_pwd.startswith("b64:"):
            config_to_save["db_password"] = _encode_password(raw_pwd)

        raw_smtp_pwd = config_to_save.get("smtp_password", "")
        if raw_smtp_pwd and not raw_smtp_pwd.startswith("b64:"):
            config_to_save["smtp_password"] = _encode_password(raw_smtp_pwd)

        # Ofuscar senhas individuais dos operadores se houverem
        if "operators" in config_to_save and isinstance(config_to_save["operators"], list):
            encoded_ops = []
            for op in config_to_save["operators"]:
                if isinstance(op, dict):
                    op_copy = op.copy()
                    raw_op_pwd = op_copy.get("smtp_password", "")
                    if raw_op_pwd and not raw_op_pwd.startswith("b64:"):
                        op_copy["smtp_password"] = _encode_password(raw_op_pwd)
                    encoded_ops.append(op_copy)
            config_to_save["operators"] = encoded_ops

        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(config_to_save, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        logging.error(f"Falha ao salvar configurações em {cfg_path}: {e}")
        return False


def get_operators(config=None):
    """Retorna a lista de operadores cadastrados."""
    if config is None:
        config = load_config()
    ops = config.get("operators")
    if not ops or not isinstance(ops, list) or len(ops) == 0:
        ops = DEFAULT_OPERATORS.copy()
        config["operators"] = ops
        config["active_operator_id"] = "breno"
    return ops


def get_active_operator(config=None):
    """Retorna o operador ativo atual (ou o padrão)."""
    if config is None:
        config = load_config()
    ops = get_operators(config)
    active_id = config.get("active_operator_id", "")
    for op in ops:
        if op.get("id") == active_id:
            return op
    for op in ops:
        if op.get("is_default"):
            return op
    return ops[0] if ops else DEFAULT_OPERATORS[0]


def set_active_operator(operator_id, config=None):
    """Define o operador ativo atual no sistema e persiste no config.json."""
    if config is None:
        config = load_config()
    config["active_operator_id"] = operator_id
    save_config(config)
    return config


def save_operator(operator_data, config=None):
    """Adiciona ou atualiza os dados de um operador."""
    if config is None:
        config = load_config()
    ops = [dict(o) for o in get_operators(config)]
    op_id = operator_data.get("id")
    if not op_id:
        import uuid
        op_id = str(uuid.uuid4())[:8]
        operator_data["id"] = op_id

    # Se for marcado como default, remove o default dos demais
    if operator_data.get("is_default"):
        for op in ops:
            op["is_default"] = False

    found = False
    for i, op in enumerate(ops):
        if op.get("id") == op_id:
            ops[i] = operator_data
            found = True
            break
    if not found:
        ops.append(operator_data)

    config["operators"] = ops
    if operator_data.get("is_default") or not config.get("active_operator_id"):
        config["active_operator_id"] = op_id

    save_config(config)
    return op_id


def delete_operator(operator_id, config=None):
    """Remove um operador (se houver mais de um)."""
    if config is None:
        config = load_config()
    ops = [dict(o) for o in get_operators(config)]
    if len(ops) <= 1:
        return False, "Não é possível remover o único operador cadastrado."

    new_ops = [op for op in ops if op.get("id") != operator_id]
    if len(new_ops) == len(ops):
        return False, "Operador não encontrado."

    config["operators"] = new_ops
    if config.get("active_operator_id") == operator_id:
        config["active_operator_id"] = new_ops[0].get("id")
        new_ops[0]["is_default"] = True

    save_config(config)
    return True, "Operador removido com sucesso."


def get_recent_reports():
    """
    Retorna a lista dos relatórios presentes na pasta de saída (inclusive em subpastas Ano/Mês),
    enriquecidos com metadados do histórico quando disponíveis.
    Organiza arquivos soltos e reflete mudanças (renomear, deletar, adicionar) em tempo real.
    """
    cfg = load_config()
    history = cfg.get("recent_reports", [])
    out_dir = cfg.get("output_directory", "")

    # Organizar relatórios soltos na raiz da pasta em subpastas Ano/Mês
    if out_dir and os.path.exists(out_dir):
        organize_loose_reports(out_dir)

    # Índice do histórico por caminho absoluto para consulta rápida de metadados
    hist_by_path = {}
    for item in history:
        if isinstance(item, dict):
            p = item.get("path", "")
            if p:
                hist_by_path[os.path.normcase(os.path.abspath(p))] = item

    result = []

    # Escaneia a pasta de saída recursivamente (suporta subpastas Ano/Mês)
    if out_dir and os.path.exists(out_dir):
        try:
            files_with_mtime = []
            for root, _, filenames in os.walk(out_dir):
                for f in filenames:
                    if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                        full_p = os.path.join(root, f)
                        try:
                            mt = os.path.getmtime(full_p)
                            files_with_mtime.append((mt, full_p))
                        except OSError:
                            pass

            # Ordenar pelos mais recentes de uma só vez (sem re-chamar stat)
            files_with_mtime.sort(key=lambda x: x[0], reverse=True)

            for mtime, f_path in files_with_mtime[:5]:
                abs_path = os.path.abspath(f_path)
                key = os.path.normcase(abs_path)
                dt_str = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")

                # Checar se existe arquivo PDF correspondente na mesma pasta
                pdf_candidate = os.path.splitext(abs_path)[0] + ".pdf"
                has_pdf = os.path.exists(pdf_candidate)

                # Se tiver metadados do histórico, usa; senão cria entrada básica
                if key in hist_by_path:
                    entry = hist_by_path[key].copy()
                    entry["filename"] = os.path.basename(f_path)
                    entry["path"] = abs_path
                    entry["has_pdf"] = has_pdf
                    entry["pdf_path"] = pdf_candidate if has_pdf else ""
                else:
                    entry = {
                        "filename": os.path.basename(f_path),
                        "path": abs_path,
                        "periodo": "",
                        "gerado_em": dt_str,
                        "total_m3": None,
                        "total_rs": None,
                        "has_pdf": has_pdf,
                        "pdf_path": pdf_candidate if has_pdf else ""
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
                    pdf_candidate = os.path.splitext(p)[0] + ".pdf"
                    item["has_pdf"] = os.path.exists(pdf_candidate)
                    item["pdf_path"] = pdf_candidate if item["has_pdf"] else ""
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
