"""
Módulo de Análise e Extração de Faturas da Sabesp para SmartHydro.
Lê faturas em PDF da Sabesp, extrai períodos de medição, consumos e tarifas marginais (Água + Esgoto).
"""

import os
import re
import imaplib
import email
from email.header import decode_header
import base64
from datetime import datetime, timedelta
import logging

import pypdf

logger = logging.getLogger("sabesp_parser")


def parse_sabesp_pdf(pdf_path: str) -> tuple[bool, str, dict | None]:
    """
    Analisa um arquivo PDF de fatura da Sabesp e extrai os metadados principais.
    Retorna (sucesso: bool, mensagem: str, dados: dict | None)
    """
    if not os.path.exists(pdf_path):
        return False, f"Arquivo não encontrado: {pdf_path}", None

    try:
        reader = pypdf.PdfReader(pdf_path)
        text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
    except Exception as e:
        return False, f"Falha ao ler o arquivo PDF: {e}", None

    if not text or ("sabesp" not in text.lower() and "saneamento" not in text.lower()):
        return False, "O arquivo fornecido não parece ser uma fatura oficial da Sabesp.", None

    # 1. Valor Total da Fatura (ex: R$ ****************38.999,00)
    m_val = re.search(r"R\$\s*[\*]*([\d\.]+,\d{2})", text)
    val_total = float(m_val.group(1).replace(".", "").replace(",", ".")) if m_val else 0.0

    # 2. Leituras e Consumo Geral Sabesp (ex: 28/08/26 28/09/26 623 31)
    m_dates = re.search(r"(\d{2}/\d{2}/\d{2,4})\s+(\d{2}/\d{2}/\d{2,4})\s+(\d+)\s+(\d+)", text)
    if not m_dates:
        m_dates = re.search(r"(\d{2}/\d{2}/\d{2,4})\s*\n?\s*(\d{2}/\d{2}/\d{2,4})\s*\n?\s*(\d+)", text)

    def _expand_year(d_str: str) -> str:
        if not d_str:
            return ""
        parts = d_str.split("/")
        if len(parts) == 3 and len(parts[2]) == 2:
            return f"{parts[0]}/{parts[1]}/20{parts[2]}"
        return d_str

    if m_dates:
        groups = m_dates.groups()
        d_ant_fmt = _expand_year(groups[0])
        d_atual_fmt = _expand_year(groups[1])
        consumo_sabesp = float(groups[2])
        dias = int(groups[3]) if len(groups) > 3 else 30
    else:
        d_ant_fmt, d_atual_fmt, consumo_sabesp, dias = "", "", 0.0, 0

    # Período de rateio das salas no condomínio:
    # O ciclo começa no dia seguinte à leitura anterior da Sabesp (00:00:00) e vai até o dia da leitura atual (23:59:59)
    periodo_ini_str = ""
    periodo_fim_str = d_atual_fmt
    if d_ant_fmt:
        try:
            dt_ant = datetime.strptime(d_ant_fmt, "%d/%m/%Y")
            dt_ini_cond = dt_ant + timedelta(days=1)
            periodo_ini_str = dt_ini_cond.strftime("%d/%m/%Y")
        except Exception:
            periodo_ini_str = d_ant_fmt

    # 3. Extração da Tarifa Marginal da Faixa (> 50 m³)
    # Na tabela: '573 31,840 18.244,32 573 31,840 18.244,32De 50,01 até 9999999'
    tarifa_agua = 0.0
    tarifa_esgoto = 0.0
    m_faixa = re.search(r"(\d+)\s+([\d\.]+,\d{3})\s+[\d\.]+,\d{2}\s+(\d+)\s+([\d\.]+,\d{3})\s+[\d\.]+,\d{2}\s*De\s+50,01", text)
    if m_faixa:
        t_agua_str = m_faixa.group(2).replace(".", "").replace(",", ".")
        t_esg_str = m_faixa.group(4).replace(".", "").replace(",", ".")
        tarifa_agua = float(t_agua_str)
        tarifa_esgoto = float(t_esg_str)
        tarifa_faixa = round(tarifa_agua + tarifa_esgoto, 2)
    else:
        # Busca genérica pelo valor de tarifa na faixa de 50 m³
        m_t = re.findall(r"(\d{2},\d{3})", text)
        if m_t:
            tarifa_agua = float(m_t[-1].replace(",", "."))
            tarifa_esgoto = tarifa_agua
            tarifa_faixa = round(tarifa_agua * 2, 2)
        else:
            tarifa_faixa = round(val_total / consumo_sabesp, 2) if consumo_sabesp > 0 else 63.68

    tarifa_media = round(val_total / consumo_sabesp, 2) if consumo_sabesp > 0 else tarifa_faixa

    # 4. Metadados complementares (Vencimento, Próxima Leitura, Hidrômetro, RGI)
    m_venc = re.search(r"VENCIMENTO:\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    vencimento = m_venc.group(1) if m_venc else ""

    m_prox = re.search(r"(\d{2}/\d{2}/\d{4})Pr[oó]xima Leitura", text, re.IGNORECASE)
    if not m_prox:
        m_prox = re.search(r"Pr[oó]xima Leitura:\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    prox_leitura = m_prox.group(1) if m_prox else ""

    m_hid = re.search(r"Hidr[oô]metro:\s*([A-Z0-9]+)", text, re.IGNORECASE)
    hidrometro = m_hid.group(1) if m_hid else ""
    if not hidrometro or len(hidrometro) > 15:
        m_hid2 = re.search(r"\d{8,10}\s+([A-Z0-9]{8,15})", text)
        if m_hid2:
            hidrometro = m_hid2.group(1)

    m_rgi = re.search(r"Pde/Rgi:\s*(\d+)", text, re.IGNORECASE)
    rgi = m_rgi.group(1) if m_rgi else ""

    data = {
        "cliente": "Condomínio Praça Pamplona",
        "leitura_anterior": d_ant_fmt,
        "leitura_atual": d_atual_fmt,
        "periodo_rateio_ini": periodo_ini_str,
        "periodo_rateio_fim": periodo_fim_str,
        "consumo_sabesp_m3": consumo_sabesp,
        "dias_faturamento": dias,
        "valor_total_fatura": val_total,
        "tarifa_agua": tarifa_agua,
        "tarifa_esgoto": tarifa_esgoto,
        "tarifa_faixa": tarifa_faixa,
        "tarifa_media": tarifa_media,
        "vencimento": vencimento,
        "proxima_leitura": prox_leitura,
        "hidrometro": hidrometro,
        "rgi": rgi,
        "arquivo_origem": os.path.basename(pdf_path),
        "caminho_completo": os.path.abspath(pdf_path)
    }

    return True, "Fatura da Sabesp identificada e processada com sucesso!", data


def search_sabesp_in_email(config: dict, destination_dir: str = None) -> tuple[bool, str, str | None]:
    """
    Conecta via IMAP SSL à caixa de entrada do operador e busca e-mails recentes
    que contenham anexo em PDF de fatura da Sabesp.
    Salva o PDF baixado em destination_dir e retorna (sucesso, mensagem, caminho_pdf).
    """
    user = config.get("smtp_user", "").strip()
    pwd_raw = config.get("smtp_password", "").strip()
    if pwd_raw.startswith("b64:"):
        try:
            pwd = base64.b64decode(pwd_raw[4:]).decode("utf-8")
        except Exception:
            pwd = pwd_raw
    else:
        pwd = pwd_raw

    if not user or not pwd:
        return False, "Usuário ou senha de e-mail não configurados nas preferências.", None

    smtp_srv = config.get("smtp_server", "smtps.uhserver.com")
    imap_srv = smtp_srv.replace("smtps.", "imap.").replace("smtp.", "imap.")
    if imap_srv == "smtps.uhserver.com":
        imap_srv = "imap.uhserver.com"

    if destination_dir is None:
        from core.config_manager import get_base_dir
        destination_dir = os.path.join(get_base_dir(), "temp_sabesp")
    os.makedirs(destination_dir, exist_ok=True)

    try:
        mail = imaplib.IMAP4_SSL(imap_srv, 993, timeout=12)
        mail.login(user, pwd)

        # Prioriza subpastas como Praça Pamplona e depois a INBOX principal
        folders_to_check = []
        try:
            typ_l, folders = mail.list()
            if typ_l == "OK" and folders:
                for f in folders:
                    f_str = f.decode("latin1", errors="ignore")
                    if "pamplona" in f_str.lower() and "enviados" not in f_str.lower() and "sent" not in f_str.lower():
                        folder_name = f_str.split(' "." ')[-1].strip()
                        if folder_name not in folders_to_check:
                            folders_to_check.append(folder_name)
        except Exception:
            pass

        if "INBOX" not in folders_to_check and '"INBOX"' not in folders_to_check:
            folders_to_check.append("INBOX")

        for folder in folders_to_check:
            try:
                typ_sel, _ = mail.select(folder)
                if typ_sel != "OK":
                    continue
            except Exception:
                continue

            typ, data = mail.search(None, "ALL")
            if typ != "OK" or not data[0]:
                continue

            msg_ids = data[0].split()
            recent_ids = msg_ids[-30:]  # Olha as últimas 30 mensagens da pasta
            recent_ids.reverse()

            for m_id in recent_ids:
                typ_f, m_data = mail.fetch(m_id, "(RFC822)")
                if typ_f != "OK" or not m_data or not m_data[0]:
                    continue

                raw_email = m_data[0][1]
                msg = email.message_from_bytes(raw_email)

                subject = ""
                for part, enc in decode_header(msg.get("Subject", "")):
                    if isinstance(part, bytes):
                        subject += part.decode(enc or "utf-8", errors="ignore")
                    else:
                        subject += str(part)

                sender = msg.get("From", "")
                is_candidate = any(k in subject.lower() for k in ["sabesp", "agua", "água", "fatura", "conta", "insumos", "pamplona", "rateio"]) or \
                               any(k in sender.lower() for k in ["sabesp", "zangari", "pamplona", "joyce", "yasmim", "gerente", "assistente"])

                # Percorre os anexos
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if part.get("Content-Disposition") is None:
                        continue

                    filename = part.get_filename()
                    if not filename:
                        continue

                    fn_clean = ""
                    for p, enc in decode_header(filename):
                        if isinstance(p, bytes):
                            fn_clean += p.decode(enc or "utf-8", errors="ignore")
                        else:
                            fn_clean += str(p)

                    if fn_clean.lower().endswith(".pdf"):
                        if is_candidate or any(k in fn_clean.lower() for k in ["sabesp", "fatura", "conta", "agua"]):
                            pdf_path = os.path.join(destination_dir, fn_clean)
                            with open(pdf_path, "wb") as f_out:
                                f_out.write(part.get_payload(decode=True))

                            # Valida se realmente é Sabesp
                            ok, _, _ = parse_sabesp_pdf(pdf_path)
                            if ok:
                                mail.logout()
                                return True, f"Fatura encontrada no e-mail de {sender} ({subject[:45]}...)", pdf_path

        mail.logout()
        return False, "Nenhuma fatura da Sabesp em PDF foi encontrada nos e-mails recentes.", None

    except Exception as e:
        return False, f"Erro ao acessar a caixa de e-mail via IMAP: {e}", None
