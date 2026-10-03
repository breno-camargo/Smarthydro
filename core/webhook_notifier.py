"""
Módulo de Notificações via Webhook para SmartHydro.
Suporta WhatsApp (CallMeBot), Microsoft Teams, Discord, Slack, Telegram e Webhook Genérico (JSON POST).
Utiliza apenas a biblioteca padrão (urllib) para total portabilidade sem dependências externas.
"""

import json
import logging
import re
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime

from core.config_manager import load_config

PLATFORMS = [
    ("whatsapp", "WhatsApp (CallMeBot Grátis / Notificação Direta)"),
    ("teams", "Microsoft Teams (Incoming Webhook)"),
    ("discord", "Discord (Canal de Alertas)"),
    ("slack", "Slack (Incoming Webhook)"),
    ("telegram", "Telegram (Bot API)"),
    ("generic", "Webhook Genérico / WhatsApp Gateway (JSON POST)")
]


def _make_http_get(url: str, timeout: int = 15) -> tuple[bool, str]:
    """Realiza uma requisição HTTP GET segura usando urllib."""
    headers = {
        "User-Agent": "SmartHydro-Notifier/2.3 (CompaSSS Praça Pamplona)"
    }
    try:
        req = urllib.request.Request(url, headers=headers, method="GET")
        with urllib.request.urlopen(req, timeout=timeout) as response:
            status = response.status
            body = response.read().decode("utf-8", errors="ignore")
            if 200 <= status < 300:
                if "error" in body.lower():
                    return False, f"CallMeBot retornou aviso: {body[:150]}"
                return True, "Mensagem enviada para o WhatsApp com sucesso!"
            return False, f"Resposta HTTP {status}: {body[:150]}"
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")[:150]
        return False, f"Erro HTTP {e.code}: {e.reason} ({err_body})"
    except urllib.error.URLError as e:
        return False, f"Falha de conexão com serviço WhatsApp: {e.reason}"
    except Exception as e:
        return False, f"Erro inesperado: {e}"


def _make_http_post(url: str, payload_dict: dict = None, custom_json_str: str = None, headers: dict = None, timeout: int = 12) -> tuple[bool, str]:
    """Realiza uma requisição HTTP POST segura usando urllib."""
    if headers is None:
        headers = {}
    if "User-Agent" not in headers:
        headers["User-Agent"] = "SmartHydro-Notifier/2.3 (CompaSSS Praça Pamplona)"
    if "Content-Type" not in headers:
        headers["Content-Type"] = "application/json; charset=utf-8"

    if custom_json_str is not None:
        data_bytes = custom_json_str.encode("utf-8")
    elif payload_dict is not None:
        data_bytes = json.dumps(payload_dict, ensure_ascii=False).encode("utf-8")
    else:
        data_bytes = b""

    try:
        req = urllib.request.Request(url, data=data_bytes, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=timeout) as response:
            status = response.status
            body = response.read().decode("utf-8", errors="ignore")
            if 200 <= status < 300:
                return True, f"Sucesso (HTTP {status})"
            return False, f"Resposta HTTP {status}: {body[:150]}"
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")[:150]
        return False, f"Erro HTTP {e.code}: {e.reason} ({err_body})"
    except urllib.error.URLError as e:
        return False, f"Falha de conexão: {e.reason}"
    except Exception as e:
        return False, f"Erro inesperado: {e}"


def _format_whatsapp_message(summary: dict) -> str:
    """Gera texto limpo e formatado com negritos e emojis para o WhatsApp."""
    periodo = summary.get("periodo", "Período")
    total_m3 = summary.get("total_m3", 0.0)
    total_rs = summary.get("total_rs", 0.0)
    operador = summary.get("operador", "Sistema Automático")
    anomalias = summary.get("anomalias", [])
    excel_file = summary.get("excel_file", "")

    m3_fmt = f"{total_m3:,.1f} m³".replace(",", "X").replace(".", ",").replace("X", ".")
    rs_fmt = f"R$ {total_rs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    anom_txt = f"⚠️ {len(anomalias)} sala(s) sob suspeita" if anomalias else "✅ Nenhuma anomalia (Normal)"

    sabesp = summary.get("sabesp")
    sabesp_line = ""
    balanco_line = ""
    if sabesp:
        m3_sab = float(sabesp.get("consumo_sabesp_m3") or 0.0)
        val_sab = float(sabesp.get("valor_total_fatura") or 0.0)
        val_sab_fmt = f"R$ {val_sab:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        sabesp_line = f"🏷️ *Fatura Sabesp:* {val_sab_fmt} (Sabesp: {m3_sab:,.0f} m³)\n"
        if m3_sab > 0 and total_m3 > 0:
            diff_m3 = m3_sab - total_m3
            diff_fmt = f"{abs(diff_m3):,.1f} m³".replace(",", "X").replace(".", ",").replace("X", ".")
            pct = (diff_m3 / m3_sab) * 100.0
            pct_fmt = f"{pct:.1f}%".replace(".", ",")
            if diff_m3 >= 0:
                st = "Normal ✅" if pct <= 10.0 else ("Atenção ⚠️" if pct <= 15.0 else "Alerta Vazamento 🚨")
                balanco_line = f"🌿 *Balanço Hídrico:* {diff_fmt} ({pct_fmt}) • Área Comum ({st})\n"
            else:
                balanco_line = f"🌿 *Balanço Hídrico:* +{diff_fmt} privativo excedente\n"

    return (
        f"💧 *SmartHydro — Relatório de Água Concluído*\n"
        f"🏢 *Condomínio Praça Pamplona*\n\n"
        f"📅 *Período:* {periodo}\n"
        f"💧 *Consumo Total:* {m3_fmt}\n"
        f"💰 *Faturamento Estimado:* {rs_fmt}\n"
        f"{sabesp_line}"
        f"{balanco_line}"
        f"🔍 *Auditoria:* {anom_txt}\n"
        f"👤 *Operador:* {operador}\n"
        f"📄 *Planilha:* {excel_file}\n\n"
        f"_Dados consolidados direto do Schneider StruxureWare EBO._"
    )


def _format_teams_card(summary: dict) -> dict:
    """Gera um MessageCard para Microsoft Teams com paleta verde CompaSSS."""
    periodo = summary.get("periodo", "Período não informado")
    total_m3 = summary.get("total_m3", 0.0)
    total_rs = summary.get("total_rs", 0.0)
    anomalias = summary.get("anomalias", [])
    operador = summary.get("operador", "Sistema Automático")
    salas_medidas = summary.get("salas_medidas", 0)

    anomalias_txt = f"{len(anomalias)} sala(s) suspeita(s)" if anomalias else "Nenhuma suspeita detectada (OK)"

    facts = [
        {"name": "📅 Período de Medição:", "value": periodo},
        {"name": "💧 Consumo Consolidado:", "value": f"**{total_m3:,.1f} m³**".replace(",", "X").replace(".", ",").replace("X", ".")},
        {"name": "💰 Faturamento Estimado:", "value": f"**R$ {total_rs:,.2f}**".replace(",", "X").replace(".", ",").replace("X", ".")},
        {"name": "🏢 Salas / Hidrômetros:", "value": f"{salas_medidas} unidades ativas" if salas_medidas else "54 unidades"},
        {"name": "⚠️ Auditoria de Consumo:", "value": anomalias_txt},
        {"name": "👤 Operador Responsável:", "value": operador}
    ]

    sabesp = summary.get("sabesp")
    if sabesp and total_m3 > 0:
        m3_sab = float(sabesp.get("consumo_sabesp_m3") or 0.0)
        if m3_sab > 0:
            diff_m3 = m3_sab - total_m3
            pct = (diff_m3 / m3_sab) * 100.0
            st = "Normal ✅" if pct <= 10.0 else ("Atenção ⚠️" if pct <= 15.0 else "Alerta Vazamento 🚨")
            facts.insert(3, {"name": "🌿 Balanço Hídrico (Área Comum):", "value": f"**{abs(diff_m3):,.1f} m³ ({pct:.1f}%)** • {st}"})

    return {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": "3D6B24",
        "summary": f"Fechamento de Hidrômetros: {periodo}",
        "sections": [{
            "activityTitle": "💧 Relatório de Hidrômetros Concluído",
            "activitySubtitle": "Condomínio Praça Pamplona • Telemetria Schneider Electric EBO",
            "facts": facts,
            "markdown": True
        }]
    }


def _format_discord_embed(summary: dict) -> dict:
    """Gera um Embed estilizado para Discord."""
    periodo = summary.get("periodo", "Período")
    total_m3 = summary.get("total_m3", 0.0)
    total_rs = summary.get("total_rs", 0.0)
    anomalias = summary.get("anomalias", [])
    operador = summary.get("operador", "Sistema Automático")

    m3_fmt = f"{total_m3:,.1f} m³".replace(",", "X").replace(".", ",").replace("X", ".")
    rs_fmt = f"R$ {total_rs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    anom_txt = f"⚠️ {len(anomalias)} sala(s)" if anomalias else "✅ Normal (0)"

    fields = [
        {"name": "📅 Período", "value": periodo, "inline": True},
        {"name": "💧 Consumo Total", "value": m3_fmt, "inline": True},
        {"name": "💰 Faturamento", "value": rs_fmt, "inline": True},
        {"name": "🔍 Auditoria / Vazamento", "value": anom_txt, "inline": True},
        {"name": "👤 Operador", "value": operador, "inline": True}
    ]

    sabesp = summary.get("sabesp")
    if sabesp and total_m3 > 0:
        m3_sab = float(sabesp.get("consumo_sabesp_m3") or 0.0)
        if m3_sab > 0:
            diff_m3 = m3_sab - total_m3
            pct = (diff_m3 / m3_sab) * 100.0
            fields.insert(3, {"name": "🌿 Balanço Hídrico", "value": f"{abs(diff_m3):,.1f} m³ ({pct:.1f}%)", "inline": True})

    return {
        "username": "SmartHydro CompaSSS",
        "embeds": [{
            "title": "💧 Fechamento Mensal de Hidrômetros — Praça Pamplona",
            "description": "O relatório de medição foi extraído com sucesso do StruxureWare EBO e compilado.",
            "color": 4025124,  # #3D6B24 em decimal
            "fields": fields,
            "footer": {"text": "SmartHydro v2.8 • CompaSSS Engenharia Predial"}
        }]
    }


def _format_slack_blocks(summary: dict) -> dict:
    """Gera mensagem formatada para Slack."""
    periodo = summary.get("periodo", "Período")
    total_m3 = summary.get("total_m3", 0.0)
    total_rs = summary.get("total_rs", 0.0)
    operador = summary.get("operador", "Sistema")
    anomalias = summary.get("anomalias", [])

    m3_fmt = f"{total_m3:,.1f} m³".replace(",", "X").replace(".", ",").replace("X", ".")
    rs_fmt = f"R$ {total_rs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    anom_txt = f"⚠️ {len(anomalias)} suspeita(s)" if anomalias else "✅ Tudo normal"

    sabesp = summary.get("sabesp")
    balanco_txt = ""
    if sabesp and total_m3 > 0:
        m3_sab = float(sabesp.get("consumo_sabesp_m3") or 0.0)
        if m3_sab > 0:
            diff_m3 = m3_sab - total_m3
            pct = (diff_m3 / m3_sab) * 100.0
            balanco_txt = f"\n• *Balanço Hídrico:* {abs(diff_m3):,.1f} m³ ({pct:.1f}% área comum)"

    msg = (
        f"💧 *SmartHydro — Relatório de Água Concluído*\n"
        f"🏢 *Condomínio Praça Pamplona*\n"
        f"• *Período:* {periodo}\n"
        f"• *Consumo:* {m3_fmt} | *Faturamento:* {rs_fmt}{balanco_txt}\n"
        f"• *Auditoria:* {anom_txt}\n"
        f"• *Operador:* {operador}"
    )
    return {"text": msg}


def _format_telegram_message(summary: dict) -> str:
    """Gera texto com formatação Markdown para Telegram."""
    periodo = summary.get("periodo", "Período")
    total_m3 = summary.get("total_m3", 0.0)
    total_rs = summary.get("total_rs", 0.0)
    operador = summary.get("operador", "Sistema Automático")
    anomalias = summary.get("anomalias", [])

    m3_fmt = f"{total_m3:,.1f} m³".replace(",", "X").replace(".", ",").replace("X", ".")
    rs_fmt = f"R$ {total_rs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    anom_txt = f"⚠️ {len(anomalias)} sala(s) sob suspeita" if anomalias else "✅ Nenhuma anomalia"

    sabesp = summary.get("sabesp")
    balanco_line = ""
    if sabesp and total_m3 > 0:
        m3_sab = float(sabesp.get("consumo_sabesp_m3") or 0.0)
        if m3_sab > 0:
            diff_m3 = m3_sab - total_m3
            pct = (diff_m3 / m3_sab) * 100.0
            balanco_line = f"🌿 *Balanço Hídrico:* `{abs(diff_m3):,.1f} m³ ({pct:.1f}%)` • Área Comum\n"

    return (
        f"💧 *SmartHydro — Fechamento Mensal de Água*\n"
        f"🏢 *Condomínio Praça Pamplona*\n\n"
        f"📅 *Período:* {periodo}\n"
        f"💧 *Consumo Total:* `{m3_fmt}`\n"
        f"💰 *Faturamento Estimado:* `{rs_fmt}`\n"
        f"{balanco_line}"
        f"🔍 *Auditoria:* {anom_txt}\n"
        f"👤 *Operador:* {operador}\n\n"
        f"_Relatório gerado com sucesso via telemetria StruxureWare EBO._"
    )


def _format_generic_json(summary: dict) -> dict:
    """Payload JSON limpo e estruturado para integrações customizadas (WhatsApp API / Node-RED / n8n)."""
    sabesp = summary.get("sabesp") or {}
    sab_m3 = sabesp.get("consumo_sabesp_m3")
    tot_m3 = summary.get("total_m3")
    diff_m3 = None
    if sab_m3 is not None and tot_m3 is not None:
        diff_m3 = round(float(sab_m3) - float(tot_m3), 2)

    return {
        "event": "hidrometro_report_generated",
        "app": "SmartHydro",
        "condominio": "Condominio Praca Pamplona",
        "periodo": summary.get("periodo"),
        "total_consumo_m3": tot_m3,
        "total_faturamento_rs": summary.get("total_rs"),
        "consumo_sabesp_m3": sab_m3,
        "balanco_hidrico_area_comum_m3": diff_m3,
        "salas_medidas": summary.get("salas_medidas", 54),
        "anomalias_detectadas": len(summary.get("anomalias", [])),
        "operador": summary.get("operador"),
        "operador_telefone": summary.get("operador_telefone", ""),
        "operador_email": summary.get("operador_email", ""),
        "arquivo_excel": summary.get("excel_file"),
        "gerado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }


def send_report_webhook(summary: dict, config: dict = None) -> tuple[bool, str]:
    """
    Envia notificação via Webhook para a plataforma configurada.
    Roteia automaticamente para o WhatsApp e credenciais do operador ativo atual.
    """
    if config is None:
        config = load_config()

    if not config.get("webhook_enabled", False):
        return False, "Webhooks desativados nas configurações."

    from core.config_manager import get_active_operator
    active_op = get_active_operator(config)

    platform = config.get("webhook_platform", "whatsapp").lower()
    url = config.get("webhook_url", "").strip()

    if platform == "whatsapp":
        # Roteamento automático: busca no perfil do operador ativo primeiro
        op_phone = ""
        op_apikey = ""
        if active_op:
            op_phone = (active_op.get("whatsapp_phone") or active_op.get("phone") or "").strip()
            op_apikey = (active_op.get("whatsapp_apikey") or "").strip()

        phone_raw = op_phone if op_phone else config.get("webhook_whatsapp_phone", "")
        phone = re.sub(r"[^\d]", "", phone_raw)
        apikey = op_apikey if op_apikey else config.get("webhook_whatsapp_apikey", "").strip()

        if not phone or not apikey:
            op_name = active_op.get("name", "Operador") if active_op else "Geral"
            return False, f"WhatsApp ou Chave API CallMeBot não configurados para '{op_name}' (ou nas configurações gerais)."

        if len(phone) in (10, 11) and not phone.startswith("55"):
            phone = "55" + phone

        text = _format_whatsapp_message(summary)
        encoded_text = urllib.parse.quote_plus(text)
        callme_url = f"https://api.callmebot.com/whatsapp.php?phone={phone}&text={encoded_text}&apikey={apikey}"
        return _make_http_get(callme_url)

    if platform == "telegram":
        token = config.get("webhook_telegram_token", "").strip()
        chat_id = config.get("webhook_telegram_chat_id", "").strip()
        if not token or not chat_id:
            return False, "Token ou Chat ID do Telegram não configurados."
        
        tele_url = f"https://api.telegram.org/bot{token}/sendMessage"
        text = _format_telegram_message(summary)
        payload = {
            "chat_id": chat_id,
            "text": text,
            "parse_mode": "Markdown"
        }
        return _make_http_post(tele_url, payload)

    if not url:
        return False, "URL do Webhook não informada."

    if platform == "teams":
        payload = _format_teams_card(summary)
    elif platform == "discord":
        payload = _format_discord_embed(summary)
    elif platform == "slack":
        payload = _format_slack_blocks(summary)
    else:  # generic
        payload = _format_generic_json(summary)

    return _make_http_post(url, payload)


def send_test_webhook(platform: str, url: str, token: str = "", chat_id: str = "", whatsapp_phone: str = "", whatsapp_apikey: str = "", operator_name: str = "") -> tuple[bool, str]:
    """Envia uma mensagem de teste para validar a conexão com o webhook ou WhatsApp."""
    platform = platform.lower()
    now_str = datetime.now().strftime("%d/%m/%Y às %H:%M:%S")

    sample_summary = {
        "periodo": "29/09/2026 a 28/10/2026 (Exemplo de Teste)",
        "total_m3": 1450.2,
        "total_rs": 92348.73,
        "salas_medidas": 54,
        "anomalias": ["Sala 1402 (Teste)"],
        "operador": operator_name or "Teste de Notificação CompaSSS",
        "excel_file": "Rateio_Agua_Teste.xlsx"
    }

    if platform == "whatsapp":
        phone = re.sub(r"[^\d]", "", whatsapp_phone)
        apikey = whatsapp_apikey.strip()
        if not phone or not apikey:
            return False, "Informe o número de telefone (com DDD) e a Apikey do CallMeBot."

        if len(phone) in (10, 11) and not phone.startswith("55"):
            phone = "55" + phone

        op_info = f" para *{operator_name}*" if operator_name else ""
        test_msg = (
            f"🔔 *SmartHydro Praça Pamplona*\n\n"
            f"Teste de notificação no WhatsApp realizado com sucesso{op_info} em {now_str}!\n\n"
            f"💧 Sistema conectado e pronto para enviar os fechamentos mensais."
        )
        callme_url = f"https://api.callmebot.com/whatsapp.php?phone={phone}&text={urllib.parse.quote_plus(test_msg)}&apikey={apikey}"
        return _make_http_get(callme_url)

    if platform == "telegram":
        if not token or not chat_id:
            return False, "Preencha o Token do Bot e o Chat ID para testar o Telegram."
        tele_url = f"https://api.telegram.org/bot{token}/sendMessage"
        text = (
            f"🔔 *Teste de Notificação SmartHydro*\n"
            f"Conexão com Telegram Bot realizada com sucesso!\n"
            f"📅 Testado em: `{now_str}`\n"
            f"🏢 Condomínio Praça Pamplona • CompaSSS"
        )
        return _make_http_post(tele_url, {"chat_id": chat_id, "text": text, "parse_mode": "Markdown"})

    if not url:
        return False, "Insira a URL do Webhook antes de testar."

    if platform == "teams":
        payload = {
            "@type": "MessageCard",
            "@context": "http://schema.org/extensions",
            "themeColor": "3D6B24",
            "summary": "Teste de Webhook SmartHydro",
            "sections": [{
                "activityTitle": "🔔 Teste de Notificação — SmartHydro",
                "activitySubtitle": f"Comunicação com Microsoft Teams OK em {now_str}",
                "facts": [
                    {"name": "Status:", "value": "Conexão Estabelecida com Sucesso"},
                    {"name": "Origem:", "value": "SmartHydro CompaSSS (Praça Pamplona)"}
                ],
                "markdown": True
            }]
        }
    elif platform == "discord":
        payload = {
            "username": "SmartHydro CompaSSS",
            "embeds": [{
                "title": "🔔 Teste de Notificação — SmartHydro",
                "description": f"Conexão via Webhook com Discord validada com sucesso em {now_str}!",
                "color": 4025124,
                "footer": {"text": "SmartHydro Praça Pamplona"}
            }]
        }
    elif platform == "slack":
        payload = {
            "text": f"🔔 *Teste de Notificação SmartHydro:* Conexão com Slack estabelecida com sucesso em {now_str}!"
        }
    else:
        payload = {
            "event": "test_ping",
            "message": "Teste de conexao SmartHydro realizado com sucesso",
            "timestamp": now_str
        }

    return _make_http_post(url, payload)
