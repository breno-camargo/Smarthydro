"""
Módulo de Backup & Restauração para SmartHydro.
Permite exportar e importar com segurança todas as configurações, perfis de operadores,
cadastros de e-mail e modelos HTML para transferência entre computadores e prevenção de falhas.
"""

import os
import json
import zipfile
import logging
import shutil
from datetime import datetime

from core.config_manager import get_base_dir, get_config_path, load_config, save_config, get_operators


def create_backup(target_zip_path: str = None) -> str:
    """
    Cria um arquivo .zip contendo todas as configurações, operadores e modelo de e-mail.
    Retorna o caminho absoluto do arquivo .zip gerado.
    """
    base_dir = get_base_dir()
    cfg_path = get_config_path()
    tmpl_path = os.path.join(base_dir, "modelo_email.html")

    now = datetime.now()
    now_str = now.strftime("%Y-%m-%d_%H%M%S")

    if not target_zip_path:
        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        target_dir = desktop_dir if os.path.exists(desktop_dir) else base_dir
        target_zip_path = os.path.join(target_dir, f"SmartHydro_Backup_{now_str}.zip")

    # Garante que a pasta destino exista
    os.makedirs(os.path.dirname(os.path.abspath(target_zip_path)), exist_ok=True)

    # Carrega dados atuais para criar manifesto
    current_cfg = load_config()
    ops = get_operators(current_cfg)

    manifest = {
        "app": "SmartHydro",
        "version": "2.3",
        "created_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "operators_count": len(ops),
        "operators_names": [op.get("name") for op in ops],
        "has_custom_template": os.path.exists(tmpl_path),
        "server": current_cfg.get("server", ""),
        "database": current_cfg.get("database", "")
    }

    manifest_json = json.dumps(manifest, indent=2, ensure_ascii=False)

    with zipfile.ZipFile(target_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # Grava manifesto
        zf.writestr("manifest.json", manifest_json)

        # Grava config.json
        if os.path.exists(cfg_path):
            zf.write(cfg_path, arcname="config.json")

        # Grava modelo_email.html se existir
        if os.path.exists(tmpl_path):
            zf.write(tmpl_path, arcname="modelo_email.html")

        # Grava logos customizadas se existirem
        for logo_name in ("logo_final.png", "gui_logo.png", "app_icon.ico"):
            logo_path = os.path.join(base_dir, logo_name)
            if os.path.exists(logo_path):
                zf.write(logo_path, arcname=logo_name)

    logging.info(f"Backup SmartHydro gerado com sucesso em: {target_zip_path}")
    return os.path.abspath(target_zip_path)


def read_backup_manifest(zip_path: str) -> tuple[bool, dict, str]:
    """Lê o manifesto interno do arquivo .zip sem extrair."""
    if not os.path.exists(zip_path):
        return False, {}, "Arquivo não encontrado."

    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            namelist = zf.namelist()
            if "config.json" not in namelist:
                return False, {}, "O arquivo selecionado não contém um config.json válido do SmartHydro."

            manifest = {}
            if "manifest.json" in namelist:
                manifest_data = zf.read("manifest.json").decode("utf-8", errors="ignore")
                manifest = json.loads(manifest_data)
            else:
                # Tenta inferir pelo config.json
                cfg_data = json.loads(zf.read("config.json").decode("utf-8", errors="ignore"))
                manifest = {
                    "app": "SmartHydro",
                    "version": "Legado",
                    "created_at": "Data não identificada",
                    "operators_count": len(cfg_data.get("operators", [])),
                    "operators_names": [op.get("name") for op in cfg_data.get("operators", [])],
                    "server": cfg_data.get("server", ""),
                    "database": cfg_data.get("database", "")
                }

            return True, manifest, "Manifesto lido com sucesso."
    except Exception as e:
        return False, {}, f"Arquivo de backup corrompido ou inválido: {e}"


def restore_backup(zip_path: str) -> tuple[bool, str, dict]:
    """
    Restaura um backup .zip para o diretório atual do sistema.
    Cria automaticamente uma cópia de segurança antes de aplicar para segurança total.
    """
    ok, manifest, msg = read_backup_manifest(zip_path)
    if not ok:
        return False, msg, None

    base_dir = get_base_dir()
    cfg_path = get_config_path()
    tmpl_path = os.path.join(base_dir, "modelo_email.html")

    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 1. Cria cópia de segurança do config.json atual
    if os.path.exists(cfg_path):
        safety_path = os.path.join(base_dir, f"config.safety_before_restore_{now_str}.json")
        try:
            shutil.copy2(cfg_path, safety_path)
        except Exception as e:
            logging.warning(f"Não foi possível criar snapshot de segurança: {e}")

    # 2. Extrai arquivos necessários
    restored_cfg = None
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            # Extrai e valida config.json
            cfg_bytes = zf.read("config.json")
            restored_cfg = json.loads(cfg_bytes.decode("utf-8"))

            # Salva no disco
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump(restored_cfg, f, indent=2, ensure_ascii=False)

            # Extrai modelo_email.html se presente
            if "modelo_email.html" in zf.namelist():
                html_bytes = zf.read("modelo_email.html")
                with open(tmpl_path, "wb") as f:
                    f.write(html_bytes)

            # Extrai imagens se presentes
            for img_name in ("logo_final.png", "gui_logo.png", "app_icon.ico"):
                if img_name in zf.namelist():
                    img_bytes = zf.read(img_name)
                    with open(os.path.join(base_dir, img_name), "wb") as f:
                        f.write(img_bytes)

        ops_count = len(restored_cfg.get("operators", []))
        return True, f"Backup restaurado com sucesso!\n• Operadores carregados: {ops_count}\n• Servidor: {restored_cfg.get('server', '-')}", restored_cfg
    except Exception as e:
        return False, f"Falha ao extrair e aplicar o backup: {e}", None
