import os
import sys
import re
import smtplib
import imaplib
import time
import mimetypes
import logging
import base64
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.header import Header
from email.utils import formatdate, make_msgid

from core.config_manager import get_base_dir, MESES_PT

logger = logging.getLogger(__name__)

DEFAULT_SUBJECT_TEMPLATE = "Relatório de insumos - {MES}/{ANO} - Praça Pamplona - CompaSSS"

DEFAULT_HTML_BODY = """<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="pt-BR">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Relatório de insumos - {MES}/{ANO} - Praça Pamplona - CompaSSS</title>
  <!--[if mso]>
  <style type="text/css">
    body, table, td, p, h1, span, a, div {font-family: Calibri, Arial, sans-serif !important;}
  </style>
  <![endif]-->
</head>
<body bgcolor="#F6F9F2" style="margin: 0; padding: 20px 0; background-color: #F6F9F2; font-family: 'Segoe UI', Calibri, Arial, sans-serif; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;">
  <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#F6F9F2" style="background-color: #F6F9F2; width: 100%; margin: 0; padding: 0;">
    <tr>
      <td align="center" valign="top" style="padding: 10px 15px;">
        <!--[if (gte mso 9)|(IE)]>
        <table role="presentation" width="640" align="center" cellpadding="0" cellspacing="0" border="0" style="width: 640px;">
          <tr>
            <td>
        <![endif]-->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" style="max-width: 640px; width: 100%; background-color: #FFFFFF; border: 1px solid #D5E5C9; border-collapse: separate; border-radius: 6px; overflow: hidden; margin: 0 auto;">
          <!-- ─── CABEÇALHO ─── -->
          <tr>
            <td bgcolor="#3D6B24" style="background-color: #3D6B24; padding: 20px 24px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td align="left" valign="middle">
                    <h1 style="margin: 0; font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 18px; font-weight: 600; color: #FFFFFF; line-height: 1.2;">Relatório de Insumos — {MES}/{ANO}</h1>
                    <p style="margin: 4px 0 0 0; font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 12px; color: #D7ECD0; line-height: 1.2;">Condomínio Praça Pamplona • Automação Predial</p>
                  </td>
                  <td align="right" valign="middle" width="155">
                    <table role="presentation" border="0" cellpadding="0" cellspacing="0" align="right" style="border-collapse: separate;">
                      <tr>
                        <td bgcolor="#4E7C33" style="background-color: #4E7C33; border: 1px solid #7FA864; border-radius: 4px; padding: 6px 12px; text-align: center;">
                          <div style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 9px; font-weight: 700; letter-spacing: 0.8px; color: #E8F5E5; text-transform: uppercase;">Telemetria BMS</div>
                          <div style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 11px; font-weight: 600; color: #FFFFFF; white-space: nowrap; margin-top: 2px;">StruxureWare EBO</div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- ─── CONTEÚDO PRINCIPAL ─── -->
          <tr>
            <td bgcolor="#FFFFFF" style="background-color: #FFFFFF; padding: 22px 24px; font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #1B2A12;">
              <p style="margin: 0 0 12px 0; font-size: 14px; color: #1B2A12;">Prezados, boa tarde!</p>

              <p style="margin: 0 0 16px 0; font-size: 14px; color: #1B2A12;">
                Segue em anexo o relatório de insumos referente ao mês de <strong>{MES}</strong>, período de <strong>{PERIODO}</strong>.
              </p>

              <!-- ─── QUADRO DE RESUMO DA MEDIÇÃO ─── -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#F6F9F2" style="background-color: #F6F9F2; border: 1px solid #D5E5C9; border-radius: 5px; margin: 16px 0;">
                <tr>
                  <td style="padding: 14px 18px;">
                    <div style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; font-weight: 700; color: #3D6B24; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 0.5px;">Resumo da Medição</div>
                    <table role="presentation" width="100%" border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse;">
                      <tr>
                        <td align="left" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; color: #55664C; padding: 3px 0;">Período:</td>
                        <td align="right" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; font-weight: 600; color: #1B2A12; padding: 3px 0;">{PERIODO}</td>
                      </tr>
                      <tr>
                        <td align="left" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; color: #55664C; padding: 3px 0;">Consumo Total:</td>
                        <td align="right" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; font-weight: 700; color: #3D6B24; padding: 3px 0;">{TOTAL_M3} m³</td>
                      </tr>
                      <tr>
                        <td align="left" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; color: #55664C; padding: 3px 0;">Valor do m³:</td>
                        <td align="right" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; font-weight: 600; color: #1B2A12; padding: 3px 0;">R$ {VALOR_M3}</td>
                      </tr>
                      <tr>
                        <td align="left" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; color: #55664C; padding: 3px 0;">Total Rateado:</td>
                        <td align="right" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; font-weight: 700; color: #3D6B24; padding: 3px 0;">R$ {TOTAL_VALOR}</td>
                      </tr>
                      <tr>
                        <td align="left" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; color: #55664C; padding: 3px 0;">Medidores:</td>
                        <td align="right" style="font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; font-weight: 600; color: #1B2A12; padding: 3px 0;">{QTD_SALAS}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- ─── QUADRO DE ANEXOS ─── -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FAFCF8" style="background-color: #FAFCF8; border-left: 3px solid #3D6B24; border-top: 1px solid #EBF3E6; border-right: 1px solid #EBF3E6; border-bottom: 1px solid #EBF3E6; margin: 16px 0;">
                <tr>
                  <td style="padding: 10px 14px; font-family: 'Segoe UI', Calibri, Arial, sans-serif; font-size: 13px; color: #44553B;">
                    <strong style="color: #1B2A12;">Arquivos em anexo:</strong>
                    <ul style="margin: 6px 0 0 0; padding-left: 18px;">
                      <li style="margin-bottom: 4px;"><strong>Planilha Excel (.xlsx)</strong> — Rateio detalhado sala a sala com histórico.</li>
                      <li style="margin-bottom: 2px;"><strong>Relatório PDF (.pdf)</strong> — Versão formatada para impressão.</li>
                    </ul>
                  </td>
                </tr>
              </table>

              <p style="margin: 16px 0 0 0; font-size: 14px; color: #1B2A12;">Quaisquer dúvidas estamos à disposição!</p>

              <!-- ─── ASSINATURA OFICIAL COMPASSS ─── -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-top: 24px; border-top: 1px solid #D5E5C9;">
                <tr>
                  <td style="padding-top: 14px;">
                    <div style="font-family: Calibri, 'Segoe UI', Arial, sans-serif; line-height: 1.35;">
                      <span style="color: #006600; font-weight: bold; font-size: 11pt;">Breno Camargo</span><br />
                      <span style="color: #006600; font-weight: bold; font-size: 10pt;">+55 11 99012 7316</span><br />
                      <a href="mailto:breno.camargo@compasss.com.br" style="color: #0563C1; text-decoration: underline; font-size: 10pt;">breno.camargo@compasss.com.br</a><br />
                      <a href="http://www.compasss.com.br" style="color: #0563C1; text-decoration: underline; font-size: 10pt;">www.compasss.com.br</a><br />
                      <span style="color: #006600; font-size: 9.5pt;">RJ – Praia de Botafogo, 300 – Mezanino – Botafogo – CEP: 20031-040</span><br />
                      <span style="color: #006600; font-size: 9.5pt;">SP – Alameda Santos, 2477 – 11º Andar – Jardim Paulista – CEP: 01419-101</span>
                    </div>
                    <div style="margin-top: 8px;">
                      <img src="cid:logo_compasss" alt="CompaSSS" width="220" style="width: 220px; height: auto; max-width: 100%; display: block; border: 0;" />
                    </div>
                  </td>
                </tr>
              </table>

            </td>
          </tr>
        </table>
        <!--[if (gte mso 9)|(IE)]>
            </td>
          </tr>
        </table>
        <![endif]-->
      </td>
    </tr>
  </table>
</body>
</html>
"""


def get_logo_image_path():
    """Localiza o ficheiro da logo CompaSSS (signature_logo.png, logo_final.png, etc.)."""
    base = get_base_dir()
    candidates = [
        os.path.join(base, "signature_logo.png"),
        os.path.join(base, "logo_final.png"),
        os.path.join(base, "excel_logo_0.png"),
        os.path.join(base, "gui_logo.png"),
        os.path.join(base, "logo_cropped.png"),
        os.path.join(base, "logo.png"),
    ]
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        candidates.insert(0, os.path.join(exe_dir, "signature_logo.png"))
        candidates.insert(1, os.path.join(exe_dir, "logo_final.png"))
        candidates.insert(2, os.path.join(exe_dir, "excel_logo_0.png"))
        candidates.insert(3, os.path.join(exe_dir, "gui_logo.png"))
    for cand in candidates:
        if os.path.exists(cand):
            return cand
    return None


def get_logo_base64():
    """Retorna a string base64 da logo CompaSSS para prévia em navegadores."""
    p = get_logo_image_path()
    if p and os.path.exists(p):
        try:
            with open(p, "rb") as f:
                return base64.b64encode(f.read()).decode("ascii")
        except Exception as e:
            logger.warning(f"Erro ao converter logo para base64: {e}")
    return ""


def prepare_html_for_preview(html_content):
    """
    Substitui referências 'cid:logo_compasss' (e URLs externas quebradas)
    por Data URI Base64 da logo real para renderização perfeita em navegadores/prévias.
    """
    if not html_content:
        return html_content

    b64 = get_logo_base64()
    if b64:
        # Substitui cid:logo_compasss
        res = re.sub(
            r'src=["\']cid:logo_compasss["\']',
            f'src="data:image/png;base64,{b64}"',
            html_content,
            flags=re.IGNORECASE
        )
        # Substitui também URL antiga ou externa se houver
        res = re.sub(
            r'src=["\']https?://[^"\']*logo[^"\']*["\']',
            f'src="data:image/png;base64,{b64}"',
            res,
            flags=re.IGNORECASE
        )
        return res
    return html_content


def get_email_template_path():
    """Retorna o caminho completo para o arquivo modelo_email.html."""
    base_dir = get_base_dir()
    tmpl_path = os.path.join(base_dir, "modelo_email.html")
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        cand = os.path.join(exe_dir, "modelo_email.html")
        if os.path.exists(cand):
            return cand
    return tmpl_path


def ensure_default_template():
    """Garante que o arquivo modelo_email.html exista no disco."""
    path = get_email_template_path()
    if not os.path.exists(path):
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(DEFAULT_HTML_BODY.strip())
        except Exception as e:
            logger.error(f"Erro ao criar modelo de e-mail padrão em {path}: {e}")
    return path


def load_email_template():
    """
    Carrega o modelo de e-mail do disco.
    Retorna uma tupla (assunto, corpo_html).
    Extrai o assunto da tag <title> ou usa o padrão.
    """
    path = get_email_template_path()
    ensure_default_template()

    try:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()

        # Extrair assunto da tag <title>...</title>
        title_match = re.search(r"<title>(.*?)</title>", content, re.IGNORECASE | re.DOTALL)
        if title_match:
            subject = title_match.group(1).strip()
        else:
            subject = DEFAULT_SUBJECT_TEMPLATE

        return subject, content
    except Exception as e:
        logger.error(f"Erro ao carregar modelo de e-mail {path}: {e}")
        return DEFAULT_SUBJECT_TEMPLATE, DEFAULT_HTML_BODY


def save_email_template(subject, html_content):
    """
    Salva o modelo de e-mail no disco, sincronizando o assunto dentro da tag <title>.
    """
    path = get_email_template_path()
    try:
        # Atualizar ou injetar <title>
        if "<title>" in html_content.lower():
            html_content = re.sub(
                r"<title>(.*?)</title>",
                f"<title>{subject}</title>",
                html_content,
                flags=re.IGNORECASE | re.DOTALL
            )
        else:
            if "<head>" in html_content.lower():
                html_content = re.sub(
                    r"(<head[^>]*>)",
                    r"\1\n  <title>" + subject + r"</title>",
                    html_content,
                    flags=re.IGNORECASE
                )
            else:
                html_content = f"<title>{subject}</title>\n" + html_content

        with open(path, "w", encoding="utf-8") as f:
            f.write(html_content)
        return True
    except Exception as e:
        logger.error(f"Falha ao salvar modelo de e-mail em {path}: {e}")
        return False


def parse_recipients(recipients_val):
    """
    Recebe string (ex: 'a@uol.com.br; b@empresa.com') ou lista de e-mails
    e retorna lista limpa de e-mails válidos.
    """
    if isinstance(recipients_val, list):
        items = recipients_val
    elif isinstance(recipients_val, str):
        items = re.split(r"[,;\s]+", recipients_val.strip())
    else:
        items = []

    valid = []
    for item in items:
        item = item.strip()
        if item and "@" in item and "." in item:
            valid.append(item)
    return valid


def format_br_currency(val):
    """Formata valor float para moeda brasileira (ex: 2.954,75)."""
    try:
        val_f = float(val)
        return f"{val_f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val or "0,00")


def format_br_number(val):
    """Formata valor float para número brasileiro com 2 casas (ex: 46,40)."""
    try:
        val_f = float(val)
        return f"{val_f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val or "0,00")


def extract_report_summary(xlsx_path):
    """
    Lê uma planilha gerada pelo sistema e extrai metadados completos:
    Período, Consumo Total, Valor Total, Valor do m³, Quantidade de Salas, Mês e Ano.
    """
    summary = {
        "periodo": "",
        "mes": "",
        "ano": "",
        "total_m3": "0,00",
        "total_valor": "0,00",
        "valor_m3": "63,68",
        "qtd_salas": "289",
        "data_emissao": datetime.now().strftime("%d/%m/%Y %H:%M")
    }

    if not xlsx_path or not os.path.exists(xlsx_path):
        return summary

    # 1. Tentar extrair mês e ano do nome do arquivo ou da pasta
    fname = os.path.basename(xlsx_path)
    for idx, m_nome in enumerate(MESES_PT, 1):
        if m_nome.lower() in fname.lower():
            summary["mes"] = m_nome
            break

    # Tentar extrair ano
    year_match = re.search(r"\b(202\d)\b", xlsx_path)
    if year_match:
        summary["ano"] = year_match.group(1)
    else:
        summary["ano"] = str(datetime.now().year)

    # 2. Ler cabeçalho e totais com openpyxl
    try:
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
        ws = wb.active

        # Linhas de cabeçalho (1 a 6)
        for r in range(1, 7):
            try:
                row_vals = [ws.cell(r, c).value for c in range(1, 10)]
            except Exception:
                continue

            for idx, val in enumerate(row_vals):
                if not val:
                    continue
                s_val = str(val).strip()

                if "período:" in s_val.lower() or "periodo:" in s_val.lower():
                    if idx + 1 < len(row_vals) and row_vals[idx + 1]:
                        summary["periodo"] = str(row_vals[idx + 1]).strip()

                if "valor do m" in s_val.lower():
                    if idx + 1 < len(row_vals) and row_vals[idx + 1]:
                        summary["valor_m3"] = format_br_currency(row_vals[idx + 1])

                if "data emiss" in s_val.lower():
                    if idx + 1 < len(row_vals) and row_vals[idx + 1]:
                        summary["data_emissao"] = str(row_vals[idx + 1]).strip()

                if "registros:" in s_val.lower():
                    if idx + 1 < len(row_vals) and row_vals[idx + 1]:
                        raw_reg = str(row_vals[idx + 1]).strip()
                        m_rooms = re.search(r"(\d+)", raw_reg)
                        if m_rooms:
                            summary["qtd_salas"] = m_rooms.group(1)

        # Totais gerais no final da planilha
        max_r = ws.max_row or 350
        for r in range(max_r, max(7, max_r - 20), -1):
            try:
                cell_a = ws.cell(r, 1).value
            except Exception:
                continue
            if cell_a and "TOTAL GERAL" in str(cell_a).upper():
                t_m3 = ws.cell(r, 2).value
                t_rs = ws.cell(r, 3).value
                if t_m3 is not None:
                    summary["total_m3"] = format_br_number(t_m3)
                if t_rs is not None:
                    summary["total_valor"] = format_br_currency(t_rs)
                break

        wb.close()
    except Exception as e:
        logger.warning(f"Erro ao extrair metadados de {xlsx_path}: {e}")

    if not summary["mes"]:
        # Fallback de mês
        curr_m = datetime.now().month
        summary["mes"] = MESES_PT[curr_m - 1]

    return summary


def render_email(subject_template, body_template, context):
    """
    Substitui as tags/placeholders do modelo pelos valores reais do contexto.
    Tags suportadas:
    {MES}, {ANO}, {PERIODO}, {TOTAL_M3}, {TOTAL_VALOR}, {VALOR_M3}, {QTD_SALAS}, {DATA_EMISSAO}
    """
    rendered_subject = subject_template
    rendered_body = body_template

    mapping = {
        "{MES}": str(context.get("mes", "")),
        "{ANO}": str(context.get("ano", "")),
        "{PERIODO}": str(context.get("periodo", "")),
        "{TOTAL_M3}": str(context.get("total_m3", "0,00")),
        "{TOTAL_VALOR}": str(context.get("total_valor", "0,00")),
        "{VALOR_M3}": str(context.get("valor_m3", "63,68")),
        "{QTD_SALAS}": str(context.get("qtd_salas", "289")),
        "{DATA_EMISSAO}": str(context.get("data_emissao", "")),
    }

    # Substituição case-insensitive
    for key, val in mapping.items():
        pattern = re.compile(re.escape(key), re.IGNORECASE)
        rendered_subject = pattern.sub(val, rendered_subject)
        rendered_body = pattern.sub(val, rendered_body)

    return rendered_subject, rendered_body


def build_mime_message(to_addrs, subject, html_body, attachment_paths, cc_addrs=None, from_addr=None, is_unsent_draft=False):
    """
    Constrói um objeto MIMEMultipart completo compatível com RFC 2822,
    com suporte a múltiplos anexos codificados em UTF-8, cópia (Cc) e cabeçalho de rascunho.
    """
    msg = MIMEMultipart("mixed")
    if is_unsent_draft:
        msg["X-Unsent"] = "1"

    msg["Subject"] = Header(subject, "utf-8")
    if from_addr:
        msg["From"] = from_addr
    msg["To"] = ", ".join(to_addrs)
    if cc_addrs:
        msg["Cc"] = ", ".join(cc_addrs)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()

    # Corpo em HTML com fallback texto simples
    # Extrai texto simples básico retirando tags HTML
    plain_text = re.sub(r"<[^>]+>", " ", html_body)
    plain_text = re.sub(r"\s+", " ", plain_text).strip()

    # Estrutura multipart/related para suportar imagem inline da logo (CID)
    related_part = MIMEMultipart("related")
    alt_part = MIMEMultipart("alternative")
    alt_part.attach(MIMEText(plain_text, "plain", "utf-8"))
    alt_part.attach(MIMEText(html_body, "html", "utf-8"))
    related_part.attach(alt_part)

    # Anexar a logo CompaSSS como imagem inline se referenciada no HTML
    if "cid:logo_compasss" in html_body.lower():
        logo_path = get_logo_image_path()
        if logo_path and os.path.exists(logo_path):
            try:
                with open(logo_path, "rb") as f_img:
                    img_part = MIMEImage(f_img.read(), _subtype="png")
                img_part.add_header("Content-ID", "<logo_compasss>")
                img_part.add_header("Content-Disposition", "inline", filename="logo_compasss.png")
                related_part.attach(img_part)
            except Exception as e:
                logger.warning(f"Erro ao anexar logo inline: {e}")

    msg.attach(related_part)

    # Anexar arquivos
    for file_path in attachment_paths:
        if not file_path or not os.path.exists(file_path):
            continue

        filename = os.path.basename(file_path)
        mime_type, _ = mimetypes.guess_type(file_path)
        if not mime_type:
            mime_type = "application/octet-stream"

        main_type, sub_type = mime_type.split("/", 1)
        try:
            with open(file_path, "rb") as f:
                part = MIMEApplication(f.read(), _subtype=sub_type)

            # Cabeçalho RFC 2231 para preservar acentos no nome do anexo
            part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", filename))
            msg.attach(part)
        except Exception as e:
            logger.error(f"Erro ao anexar arquivo {file_path}: {e}")

    return msg


def open_in_outlook(to_addrs, subject, html_body, attachment_paths, cc_addrs=None):
    """
    Abre o e-mail diretamente no Microsoft Outlook com rascunho preenchido,
    destinatário principal (gerente), pessoas em cópia (Cc) e os anexos (Excel e PDF).
    Tenta primeiro via Outlook COM; se não for possível, utiliza arquivo EML nativo.
    """
    # 1. Tentar criar arquivo .eml com X-Unsent: 1 e abrir nativamente no Outlook
    import tempfile
    try:
        temp_dir = tempfile.gettempdir()
        safe_subject = re.sub(r'[\\/*?:"<>|]', "", subject)[:40].strip() or "Relatorio_Agua"
        eml_filename = f"{safe_subject}_{datetime.now().strftime('%H%M%S')}.eml"
        eml_path = os.path.join(temp_dir, eml_filename)

        msg = build_mime_message(
            to_addrs=to_addrs,
            subject=subject,
            html_body=html_body,
            attachment_paths=attachment_paths,
            cc_addrs=cc_addrs,
            is_unsent_draft=True
        )

        with open(eml_path, "wb") as f:
            f.write(msg.as_bytes())

        os.startfile(eml_path)
        return True, "E-mail aberto no Microsoft Outlook com sucesso!"
    except Exception as e:
        logger.warning(f"Falha ao abrir via EML ({e}), tentando COM...")

    # 2. Fallback COM se startfile falhar
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = "; ".join(to_addrs)
        if cc_addrs:
            mail.CC = "; ".join(cc_addrs)
        mail.Subject = subject
        mail.HTMLBody = html_body

        for f_path in attachment_paths:
            if f_path and os.path.exists(f_path):
                mail.Attachments.Add(os.path.abspath(f_path))

        if "cid:logo_compasss" in html_body.lower():
            logo_path = get_logo_image_path()
            if logo_path and os.path.exists(logo_path):
                try:
                    att = mail.Attachments.Add(os.path.abspath(logo_path))
                    att.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_compasss"
                    )
                except Exception as ex_cid:
                    logger.warning(f"Erro ao configurar CID no Outlook COM: {ex_cid}")

        mail.Display()
        return True, "E-mail exibido no Microsoft Outlook!"
    except Exception as com_err:
        return False, f"Não foi possível abrir o e-mail no Outlook: {com_err}"


def get_effective_smtp_host(host, user):
    """Retorna o host correto do UOL (smtps.uol.com.br para @uol.com.br ou smtps.uhserver.com para domínios próprios)."""
    h = (host or "").strip()
    u = (user or "").strip()
    if u.lower().endswith("@uol.com.br") and h in ("", "smtps.uhserver.com"):
        return "smtps.uol.com.br"
    return h or "smtps.uhserver.com"


def get_effective_imap_host(user):
    """Retorna o servidor IMAP correto com base no domínio do usuário."""
    u = (user or "").strip().lower()
    if u.endswith("@uol.com.br"):
        return "imap.uol.com.br"
    return "imap.uhserver.com"


def save_email_to_sent_imap(smtp_cfg, msg_bytes):
    """
    Salva uma cópia exata do e-mail enviado na pasta 'Itens Enviados' da conta via IMAP.
    Dessa forma, o e-mail aparece no Outlook e Webmail com a data e hora em que foi enviado.
    """
    user = smtp_cfg.get("smtp_user", "").strip()
    pwd = smtp_cfg.get("smtp_password", "").strip()
    if not user or not pwd:
        return False, "Usuário ou senha não configurados."

    imap_host = get_effective_imap_host(user)

    try:
        imap = imaplib.IMAP4_SSL(imap_host, 993, timeout=15)
        imap.login(user, pwd)

        # Detectar a pasta de Enviados do servidor
        target_folder = None
        status, folder_list = imap.list()
        if status == "OK" and folder_list:
            candidates = ["itens enviados", "sent", "sent items", "inbox.sent", "inbox/sent", "enviados"]
            for f_bytes in folder_list:
                f_str = f_bytes.decode("utf-8", errors="ignore")
                for cand in candidates:
                    if f'"{cand}"' in f_str.lower() or f' {cand}' in f_str.lower() or f'.{cand}' in f_str.lower():
                        for sep in [' "/" ', ' "." ', ' ']:
                            parts = f_str.split(sep)
                            if len(parts) > 1:
                                cand_name = parts[-1].strip().strip('"')
                                if cand in cand_name.lower():
                                    target_folder = cand_name
                                    break
                        if target_folder:
                            break
                if target_folder:
                    break

        if not target_folder:
            target_folder = "INBOX.Sent"

        now_time = imaplib.Time2Internaldate(time.time())
        res, _ = imap.append(target_folder, r'(\Seen)', now_time, msg_bytes)

        if res != "OK":
            for fallback in ["Sent", "Itens Enviados", "INBOX.Sent", "INBOX/Sent", "Enviados"]:
                if fallback != target_folder:
                    res, _ = imap.append(fallback, r'(\Seen)', now_time, msg_bytes)
                    if res == "OK":
                        target_folder = fallback
                        break

        imap.logout()
        if res == "OK":
            logger.info(f"Cópia arquivada com sucesso em '{target_folder}' no servidor via IMAP.")
            return True, target_folder
        return False, f"Resposta IMAP: {res}"
    except Exception as e:
        logger.warning(f"Erro ao salvar cópia em Itens Enviados via IMAP: {e}")
        return False, str(e)


def send_email_smtp(smtp_cfg, to_addrs, subject, html_body, attachment_paths, cc_addrs=None):
    """
    Envia o e-mail diretamente via servidor SMTP (ex: UOL Pro - smtps.uhserver.com / smtps.uol.com.br)
    para o destinatário principal e cópias (Cc), e salva uma cópia na pasta 'Itens Enviados' via IMAP.
    """
    raw_host = smtp_cfg.get("smtp_server", "").strip()
    user = smtp_cfg.get("smtp_user", "").strip()
    server_host = get_effective_smtp_host(raw_host, user)
    port = int(smtp_cfg.get("smtp_port", 587))
    use_ssl = bool(smtp_cfg.get("smtp_use_ssl", False))
    use_tls = bool(smtp_cfg.get("smtp_use_tls", True))
    pwd = smtp_cfg.get("smtp_password", "").strip()

    if not server_host:
        return False, "Endereço do servidor SMTP não informado."
    if not to_addrs:
        return False, "Nenhum destinatário informado."
    if not user:
        return False, "Por favor, preencha o seu e-mail do UOL nas configurações para enviar."

    from_addr = user
    msg = build_mime_message(
        to_addrs=to_addrs,
        subject=subject,
        html_body=html_body,
        attachment_paths=attachment_paths,
        cc_addrs=cc_addrs,
        from_addr=from_addr,
        is_unsent_draft=False
    )

    # Lista total de envelopes para entrega
    all_recipients = list(to_addrs) + list(cc_addrs or [])

    try:
        if use_ssl or port == 465:
            server = smtplib.SMTP_SSL(server_host, port, timeout=25)
        else:
            server = smtplib.SMTP(server_host, port, timeout=25)
            server.ehlo()
            if use_tls:
                server.starttls()
                server.ehlo()

        if user and pwd:
            server.login(user, pwd)

        server.sendmail(from_addr, all_recipients, msg.as_string())
        server.quit()

        # Salvar cópia na pasta 'Itens Enviados' do UOL Pro via IMAP
        saved_sent = False
        try:
            ok_imap, _ = save_email_to_sent_imap(smtp_cfg, msg.as_bytes())
            saved_sent = ok_imap
        except Exception as ex_imap:
            logger.warning(f"Não foi possível arquivar em Itens Enviados: {ex_imap}")

        cc_count_txt = f" e {len(cc_addrs)} em cópia" if cc_addrs else ""
        msg_result = f"E-mail enviado com sucesso para {len(to_addrs)} destinatário(s){cc_count_txt}!"
        if saved_sent:
            msg_result += "\n\n✓ Cópia arquivada na sua pasta 'Itens Enviados' da conta UOL Pro."

        return True, msg_result
    except smtplib.SMTPAuthenticationError:
        return False, "Erro de autenticação SMTP: Usuário ou senha do e-mail incorretos."
    except smtplib.SMTPConnectError as ce:
        return False, f"Falha ao conectar no servidor SMTP ({server_host}:{port}): {ce}"
    except Exception as e:
        return False, f"Erro ao enviar e-mail via SMTP: {e}"


def test_smtp_connection(smtp_cfg):
    """
    Testa a conexão e credenciais com o servidor SMTP sem enviar mensagem.
    """
    raw_host = smtp_cfg.get("smtp_server", "").strip()
    user = smtp_cfg.get("smtp_user", "").strip()
    server_host = get_effective_smtp_host(raw_host, user)
    port = int(smtp_cfg.get("smtp_port", 587))
    use_ssl = bool(smtp_cfg.get("smtp_use_ssl", False))
    use_tls = bool(smtp_cfg.get("smtp_use_tls", True))
    pwd = smtp_cfg.get("smtp_password", "").strip()

    if not server_host:
        return False, "Servidor SMTP não configurado."

    try:
        if use_ssl or port == 465:
            server = smtplib.SMTP_SSL(server_host, port, timeout=12)
        else:
            server = smtplib.SMTP(server_host, port, timeout=12)
            server.ehlo()
            if use_tls:
                server.starttls()
                server.ehlo()

        if user and pwd:
            server.login(user, pwd)

        server.quit()
        return True, f"Conexão com {server_host}:{port} e autenticação realizadas com sucesso!"
    except smtplib.SMTPAuthenticationError:
        return False, "Conectado ao servidor, mas usuário ou senha foram rejeitados."
    except Exception as e:
        return False, f"Falha ao conectar com {server_host}:{port}:\n{e}"
