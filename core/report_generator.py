import os
import sys
import math
import shutil
import logging
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.worksheet.page import PageMargins
from openpyxl.drawing.image import Image as XlImage

from core.config_manager import get_base_dir

# ─────────────────────────────────────────────────────────────
# PALETA CORPORATIVA EXECUTIVA — CompaSSS (Elegante & Minimalista)
# ─────────────────────────────────────────────────────────────
CLR_FOREST_DEEP  = "1E3A16"   # Verde floresta profundo (cabeçalhos nobres, totais)
CLR_BRAND_GREEN  = "3D6B24"   # Verde institucional CompaSSS
CLR_ACCENT_GREEN = "8CC63F"   # Verde sutil da logo (#90C671 / #8CC63F)
CLR_CARD_BG      = "F8FAF6"   # Fundo suave de cards (off-white com nuance sálvia)
CLR_ZEBRA_LIGHT  = "FAFCF8"   # Alternância ultra-sutil (quase imperceptível)
CLR_LINE_SUBTLE  = "E2E8F0"   # Linha divisória horizontal fina e limpa
CLR_TEXT_DARK    = "1E293B"   # Carvão profundo de alto contraste (elegante)
CLR_TEXT_MUTED   = "64748B"   # Cinza ardósia para subtítulos e metadados
CLR_TEXT_ZERO    = "94A3B8"   # Cinza suave para salas zeradas (destaca quem consumiu)
CLR_TEXT_WHITE   = "FFFFFF"

# Compatibilidade
CLR_GREEN_DARK   = CLR_FOREST_DEEP
CLR_GREEN_MED    = CLR_BRAND_GREEN
CLR_GREEN_BRAND  = CLR_ACCENT_GREEN
CLR_GREEN_LIGHT  = "B8D98A"
CLR_GREEN_PALE   = CLR_CARD_BG

CLR_BG_TITLE     = CLR_FOREST_DEEP
CLR_BG_HEADER    = CLR_FOREST_DEEP
CLR_BG_ZEBRA_A   = "FFFFFF"
CLR_BG_ZEBRA_B   = CLR_ZEBRA_LIGHT
CLR_BG_TOTAL     = CLR_CARD_BG
CLR_BG_META      = CLR_CARD_BG
CLR_BG_KPI       = CLR_CARD_BG

CLR_ACCENT_GOLD  = "C2A33A"
CLR_ACCENT_LINE  = CLR_ACCENT_GREEN
CLR_BORDER_LIGHT  = CLR_LINE_SUBTLE
CLR_BORDER_HEADER = CLR_FOREST_DEEP


def get_template_path():
    """Retorna o caminho do arquivo de modelo modelo_relatorio.xlsx se ele existir."""
    candidates = [
        os.path.join(get_base_dir(), "modelo_relatorio.xlsx"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "modelo_relatorio.xlsx"),
    ]
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        candidates.insert(0, os.path.join(exe_dir, "modelo_relatorio.xlsx"))

    for p in candidates:
        if os.path.exists(p):
            return os.path.abspath(p)
    return None


def open_template_in_excel():
    """Abre o arquivo modelo no Excel padrão do Windows para edição do usuário."""
    tmpl_path = get_template_path()
    if not tmpl_path:
        # Se ainda não existir, cria o modelo inicial primeiro
        base_dir = get_base_dir()
        tmpl_path = os.path.join(base_dir, "modelo_relatorio.xlsx")
        create_default_template_file(tmpl_path)
    os.startfile(tmpl_path)
    return tmpl_path


def _get_logo_path():
    """Localiza o ficheiro da logo CompaSSS."""
    candidates = [
        os.path.join(get_base_dir(), "logo_final.png"),
        os.path.join(get_base_dir(), "logo_cropped.png"),
        os.path.join(get_base_dir(), "logo.png"),
        os.path.join(get_base_dir(), "gui_logo.png"),
    ]
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
        candidates.insert(0, os.path.join(base, "logo_final.png"))
        candidates.insert(1, os.path.join(base, "logo_cropped.png"))
        candidates.insert(2, os.path.join(base, "logo.png"))
        candidates.insert(3, os.path.join(base, "gui_logo.png"))

    for p in candidates:
        if os.path.exists(p):
            return os.path.abspath(p)
    return None


def _safe_float(val):
    """Converte valores com segurança para float, retornando 0.0 para None, NaN ou inválidos."""
    if val is None:
        return 0.0
    try:
        f = float(val)
        return 0.0 if math.isnan(f) else f
    except (ValueError, TypeError):
        return 0.0


def sort_dataframe_by_consumption(df):
    """
    Ordena o DataFrame pelo consumo em m³ em ordem decrescente (do maior consumidor para o menor).
    Em caso de empate no consumo (ex: salas zeradas), ordena em ordem alfabética pelo nome da sala.
    """
    if df is None or df.empty:
        return df
    df_sorted = df.copy()
    df_sorted["_sort_consumo"] = df_sorted["Consumo_m3"].apply(_safe_float)
    df_sorted = df_sorted.sort_values(by=["_sort_consumo", "Usuario"], ascending=[False, True]).reset_index(drop=True)
    return df_sorted.drop(columns=["_sort_consumo"])


def _apply_cell(cell, font=None, fill=None, alignment=None, border=None, number_format=None):
    """Helper para aplicar estilos múltiplos a uma célula."""
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if number_format:
        cell.number_format = number_format


def _configure_print_settings(ws, last_data_row=None):
    """Configura a planilha para impressão otimizada: paisagem, ajustada à largura, margens reduzidas."""
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 0 = quantas páginas de altura forem necessárias
    # Ativar o modo "Ajustar à página" (fitToPage) no Excel
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(
        left=0.4, right=0.4,
        top=0.5, bottom=0.5,
        header=0.3, footer=0.3
    )
    # Repetir cabeçalho em todas as páginas impressas
    ws.print_title_rows = '1:7'
    # Área de impressão (se última linha conhecida)
    if last_data_row:
        ws.print_area = f'A1:G{last_data_row + 5}'


def format_date_display(d_str):
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y'):
        try:
            return datetime.strptime(d_str, fmt).strftime('%d/%m/%Y')
        except ValueError:
            pass
    return str(d_str)


def _build_graphics_sheet(wb, df, dt_inicio_str, dt_fim_str, valor_m3):
    """
    Cria a aba 'Gráficos' com padrão executivo / diretoria:
    - Tipografia Segoe UI refinada
    - Cartões de KPI modernos (estilo Metric Cards)
    - Tabela de ranking Top 10 estilo leaderboard (sem grades verticais)
    - Gráfico de barras flat institucional CompaSSS
    - Gráfico de proporção executivo
    """
    if "Gráficos" in wb.sheetnames:
        del wb["Gráficos"]

    ws_g = wb.create_sheet(title="Gráficos")
    ws_g.views.sheetView[0].showGridLines = True

    # Tipografia Executiva
    font_brand_title = Font(name="Segoe UI", size=13, bold=True, color=CLR_FOREST_DEEP)
    font_sub         = Font(name="Segoe UI", size=9, color=CLR_TEXT_MUTED)
    font_tbl_hdr     = Font(name="Segoe UI", size=9, bold=True, color=CLR_TEXT_WHITE)
    font_rank        = Font(name="Segoe UI", size=9, bold=True, color=CLR_TEXT_MUTED)
    font_data        = Font(name="Segoe UI", size=9, color=CLR_TEXT_DARK)
    font_total_bold  = Font(name="Segoe UI", size=9, bold=True, color=CLR_FOREST_DEEP)
    font_kpi_lbl     = Font(name="Segoe UI", size=8, bold=True, color=CLR_TEXT_MUTED)
    font_kpi_val     = Font(name="Segoe UI", size=13, bold=True, color=CLR_FOREST_DEEP)
    font_kpi_pct     = Font(name="Segoe UI", size=13, bold=True, color=CLR_BRAND_GREEN)

    # Preenchimentos
    fill_hdr   = PatternFill(start_color=CLR_FOREST_DEEP, end_color=CLR_FOREST_DEEP, fill_type="solid")
    fill_subhdr= PatternFill(start_color="2D4F1E", end_color="2D4F1E", fill_type="solid")
    fill_za    = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_zb    = PatternFill(start_color=CLR_ZEBRA_LIGHT, end_color=CLR_ZEBRA_LIGHT, fill_type="solid")
    fill_card  = PatternFill(start_color=CLR_CARD_BG, end_color=CLR_CARD_BG, fill_type="solid")

    # Bordas
    side_subtle   = Side(style="thin", color=CLR_LINE_SUBTLE)
    side_top_tot  = Side(style="thin", color=CLR_FOREST_DEEP)
    side_bot_tot  = Side(style="double", color=CLR_FOREST_DEEP)

    border_card   = Border(left=side_subtle, right=side_subtle, top=side_subtle, bottom=side_subtle)
    border_row    = Border(top=side_subtle, bottom=side_subtle)
    border_total  = Border(top=side_top_tot, bottom=side_bot_tot)

    # Larguras das colunas para respiração ideal
    ws_g.column_dimensions['A'].width = 7
    ws_g.column_dimensions['B'].width = 25
    ws_g.column_dimensions['C'].width = 15
    ws_g.column_dimensions['D'].width = 15
    ws_g.column_dimensions['E'].width = 15
    ws_g.column_dimensions['F'].width = 16
    ws_g.column_dimensions['G'].width = 13
    ws_g.column_dimensions['H'].width = 3

    # Cabeçalho Limpo e Elegante
    ws_g.merge_cells('A1:G1')
    ws_g['A1'] = "PAINEL EXECUTIVO DE CONSUMO E INDICADORES"
    ws_g['A1'].font = font_brand_title
    ws_g['A1'].alignment = Alignment(horizontal="left", vertical="center")
    ws_g.row_dimensions[1].height = 24

    d_ini_disp = format_date_display(dt_inicio_str)
    d_fim_disp = format_date_display(dt_fim_str)
    ws_g.merge_cells('A2:G2')
    ws_g['A2'] = f"Condomínio Praça Pamplona  •  Ciclo de Medição: {d_ini_disp} a {d_fim_disp}"
    ws_g['A2'].font = font_sub
    ws_g['A2'].alignment = Alignment(horizontal="left", vertical="center")
    ws_g.row_dimensions[2].height = 18

    # Linha divisória verde suave
    for c in range(1, 8):
        ws_g.cell(row=3, column=c).border = Border(bottom=Side(style="thin", color=CLR_ACCENT_GREEN))
    ws_g.row_dimensions[3].height = 6

    # Ordenar dados e calcular métricas
    df_sorted = sort_dataframe_by_consumption(df)
    total_consumo = df_sorted["Consumo_m3"].apply(_safe_float).sum()
    total_valor = df_sorted["Valor_RS"].apply(_safe_float).sum()
    total_cm11 = df_sorted["Consumo_Medio_11m"].apply(_safe_float).sum() if "Consumo_Medio_11m" in df_sorted.columns else 0.0
    var_global = ((total_consumo - total_cm11) / total_cm11) if total_cm11 > 0 else 0.0

    df_top10 = df_sorted.head(10).copy()
    top10_consumo = df_top10["Consumo_m3"].apply(_safe_float).sum()
    top10_cm11 = df_top10["Consumo_Medio_11m"].apply(_safe_float).sum() if "Consumo_Medio_11m" in df_top10.columns else 0.0
    top10_valor = df_top10["Valor_RS"].apply(_safe_float).sum()
    top10_var = ((top10_consumo - top10_cm11) / top10_cm11) if top10_cm11 > 0 else 0.0
    resto_consumo = max(total_consumo - top10_consumo, 0.0)
    num_salas = len(df_sorted)

    # ── Cartões de Indicadores (Metric Cards nas Linhas 5 e 6) ──
    # Card 1: Consumo Total (A5:B6)
    ws_g.merge_cells('A5:B5')
    ws_g['A5'] = "CONSUMO TOTAL"
    ws_g['A5'].font = font_kpi_lbl
    ws_g['A5'].fill = fill_card
    ws_g['A5'].alignment = Alignment(horizontal="center", vertical="center")

    ws_g.merge_cells('A6:B6')
    ws_g['A6'] = total_consumo
    ws_g['A6'].font = font_kpi_val
    ws_g['A6'].fill = fill_card
    ws_g['A6'].alignment = Alignment(horizontal="center", vertical="center")
    ws_g['A6'].number_format = '#,##0.00 "m³"'

    # Card 2: Valor Total (C5:D6)
    ws_g.merge_cells('C5:D5')
    ws_g['C5'] = "VALOR TOTAL RATEADO"
    ws_g['C5'].font = font_kpi_lbl
    ws_g['C5'].fill = fill_card
    ws_g['C5'].alignment = Alignment(horizontal="center", vertical="center")

    ws_g.merge_cells('C6:D6')
    ws_g['C6'] = total_valor
    ws_g['C6'].font = font_kpi_val
    ws_g['C6'].fill = fill_card
    ws_g['C6'].alignment = Alignment(horizontal="center", vertical="center")
    ws_g['C6'].number_format = 'R$ #,##0.00'

    # Card 3: % Top 10 (E5:E6)
    ws_g['E5'] = "% TOP 10"
    ws_g['E5'].font = font_kpi_lbl
    ws_g['E5'].fill = fill_card
    ws_g['E5'].alignment = Alignment(horizontal="center", vertical="center")

    ws_g['E6'] = (top10_consumo / total_consumo) if total_consumo > 0 else 0.0
    ws_g['E6'].font = font_kpi_pct
    ws_g['E6'].fill = fill_card
    ws_g['E6'].alignment = Alignment(horizontal="center", vertical="center")
    ws_g['E6'].number_format = '0.0%'

    # Card 4: Variação Geral vs Média Histórica (F5:G6)
    ws_g.merge_cells('F5:G5')
    ws_g['F5'] = "VARIAÇÃO GERAL VS MÉDIA"
    ws_g['F5'].font = font_kpi_lbl
    ws_g['F5'].fill = fill_card
    ws_g['F5'].alignment = Alignment(horizontal="center", vertical="center")

    ws_g.merge_cells('F6:G6')
    ws_g['F6'] = var_global
    ws_g['F6'].font = font_kpi_val
    ws_g['F6'].fill = fill_card
    ws_g['F6'].alignment = Alignment(horizontal="center", vertical="center")
    ws_g['F6'].number_format = '+0.0%;-0.0%;"0.0%"'

    for r in (5, 6):
        ws_g.row_dimensions[r].height = 20
        for c in range(1, 8):
            ws_g.cell(row=r, column=c).border = border_card

    ws_g.row_dimensions[7].height = 10

    # ── Tabela Leaderboard Top 10 com Comparativo (Linhas 8 a 19) ──
    headers = ["Rank", "Sala / Unidade", "Consumo (m³)", "Média 11m (m³)", "Variação %", "Valor Rateado", "% do Total"]
    for col_i, h in enumerate(headers, 1):
        c = ws_g.cell(row=8, column=col_i, value=h)
        c.font = font_tbl_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center" if col_i in (1, 5, 7) else ("left" if col_i == 2 else "right"), vertical="center")
    ws_g.row_dimensions[8].height = 26

    for idx, (_, r_top) in enumerate(df_top10.iterrows(), 1):
        rn = 8 + idx
        ws_g.row_dimensions[rn].height = 21
        sala_name = str(r_top["Usuario"])
        c_val = _safe_float(r_top.get("Consumo_m3"))
        cm_val = _safe_float(r_top.get("Consumo_Medio_11m"))
        var_pct = ((c_val - cm_val) / cm_val) if cm_val > 0 else 0.0
        v_val = _safe_float(r_top.get("Valor_RS"))
        pct = (c_val / total_consumo) if total_consumo > 0 else 0.0

        vals = [f"{idx:02d}", sala_name, c_val, cm_val, var_pct, v_val, pct]
        fill = fill_za if idx % 2 == 1 else fill_zb

        for c_i, val in enumerate(vals, 1):
            cell = ws_g.cell(row=rn, column=c_i, value=val)
            cell.font = font_rank if c_i == 1 else font_data
            cell.fill = fill
            cell.border = border_row
            if c_i == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_i == 3:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0.00'
            elif c_i == 4:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0.00'
            elif c_i == 5:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '+0.0%;-0.0%;"0.0%"'
            elif c_i == 6:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = 'R$ #,##0.00'
            elif c_i == 7:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '0.0%'

    # Linha Total Top 10 com Accounting Double Underline
    row_sub = 8 + len(df_top10) + 1
    ws_g.row_dimensions[row_sub].height = 23

    ws_g.cell(row=row_sub, column=1, value="TOTAL")
    ws_g.cell(row=row_sub, column=2, value="Top 10 Maiores Consumidores")
    ws_g.cell(row=row_sub, column=3, value=top10_consumo).number_format = '#,##0.00'
    ws_g.cell(row=row_sub, column=4, value=top10_cm11).number_format = '#,##0.00'
    ws_g.cell(row=row_sub, column=5, value=top10_var).number_format = '+0.0%;-0.0%;"0.0%"'
    ws_g.cell(row=row_sub, column=6, value=top10_valor).number_format = 'R$ #,##0.00'
    ws_g.cell(row=row_sub, column=7, value=(top10_consumo / total_consumo) if total_consumo > 0 else 0.0).number_format = '0.0%'

    for c in range(1, 8):
        cell = ws_g.cell(row=row_sub, column=c)
        cell.font = font_total_bold
        cell.fill = fill_card
        cell.border = border_total
        al = Alignment(horizontal="center" if c in (1, 5, 7) else ("left" if c == 2 else "right"), vertical="center")
        cell.alignment = al

    # ── Tabela Distribuição Geral (Linhas 22 a 25) ──
    ws_g.cell(row=22, column=1, value="Grupo").font = font_tbl_hdr
    ws_g.cell(row=22, column=1).fill = fill_subhdr
    ws_g.cell(row=22, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws_g.merge_cells("A22:B22")

    ws_g.cell(row=22, column=3, value="Consumo (m³)").font = font_tbl_hdr
    ws_g.cell(row=22, column=3).fill = fill_subhdr
    ws_g.cell(row=22, column=3).alignment = Alignment(horizontal="right", vertical="center")

    ws_g.cell(row=23, column=1, value="Top 10 Maiores Consumidores").font = font_data
    ws_g.cell(row=23, column=1).fill = fill_za
    ws_g.cell(row=23, column=1).border = border_row
    ws_g.merge_cells("A23:B23")

    c23 = ws_g.cell(row=23, column=3, value=top10_consumo)
    c23.font = font_data
    c23.fill = fill_za
    c23.border = border_row
    c23.alignment = Alignment(horizontal="right", vertical="center")
    c23.number_format = '#,##0.00'

    demais_count = max(num_salas - len(df_top10), 0)
    ws_g.cell(row=24, column=1, value=f"Demais Salas ({demais_count} un)").font = font_data
    ws_g.cell(row=24, column=1).fill = fill_zb
    ws_g.cell(row=24, column=1).border = border_row
    ws_g.merge_cells("A24:B24")

    c24 = ws_g.cell(row=24, column=3, value=resto_consumo)
    c24.font = font_data
    c24.fill = fill_zb
    c24.border = border_row
    c24.alignment = Alignment(horizontal="right", vertical="center")
    c24.number_format = '#,##0.00'

    for r in (22, 23, 24):
        ws_g.row_dimensions[r].height = 20

    # ── Gráfico 1: Barras Comparativo Top 10 (Consumo Atual vs Média 11m) a partir de I5 ──
    if len(df_top10) > 0 and top10_consumo > 0:
        bar = BarChart()
        bar.type = "col"
        bar.style = 10
        bar.title = "Top 10 — Consumo Atual vs. Média Histórica (m³)"
        bar.y_axis.title = "Consumo (m³)"
        bar.x_axis.title = None
        bar.width = 23
        bar.height = 13

        # Dados: colunas 3 (Consumo) e 4 (Média 11m)
        data_bar = Reference(ws_g, min_col=3, min_row=8, max_col=4, max_row=8 + len(df_top10))
        cats_bar = Reference(ws_g, min_col=2, min_row=9, max_row=8 + len(df_top10))
        bar.add_data(data_bar, titles_from_data=True)
        bar.set_categories(cats_bar)

        if len(bar.series) >= 2:
            bar.series[0].graphicalProperties.solidFill = CLR_BRAND_GREEN
            bar.series[1].graphicalProperties.solidFill = CLR_ACCENT_GREEN
        elif len(bar.series) == 1:
            bar.series[0].graphicalProperties.solidFill = CLR_BRAND_GREEN

        ws_g.add_chart(bar, "A27")

    # ── Gráfico 2: Proporção do Consumo (abaixo do gráfico de barras) ──
    if total_consumo > 0:
        pie = PieChart()
        pie.title = "Distribuição do Consumo: Top 10 vs. Demais Salas"
        pie.width = 23
        pie.height = 12
        data_pie = Reference(ws_g, min_col=3, min_row=22, max_row=24)
        cats_pie = Reference(ws_g, min_col=1, min_row=23, max_row=24)
        pie.add_data(data_pie, titles_from_data=True)
        pie.set_categories(cats_pie)
        ws_g.add_chart(pie, "A44")

    # Configurar impressão da aba Gráficos
    _configure_print_settings(ws_g, last_data_row=60)


def generate_excel_from_template(template_path, df, dt_inicio_str, dt_fim_str, valor_m3, output_path, sort_by_consumption=False, progress_callback=None):
    """
    Gera o relatório preservando 100% o layout, logos, imagens e cabeçalhos do arquivo modelo_relatorio.xlsx.
    Retorna (output_path, warnings) onde warnings é uma lista de avisos não-fatais.
    """
    warnings = []

    if progress_callback:
        progress_callback(60, "3/4: Formatando dados na planilha executiva...")

    if sort_by_consumption:
        df = sort_dataframe_by_consumption(df)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    shutil.copyfile(template_path, output_path)

    wb = load_workbook(output_path)
    ws = wb.active

    # Atualizar Metadados
    ws["B4"].value = f"{format_date_display(dt_inicio_str)}  a  {format_date_display(dt_fim_str)}"
    ws["E4"].value = datetime.now().strftime('%d/%m/%Y %H:%M')
    ws["B5"].value = float(valor_m3)
    ws["B5"].number_format = 'R$ #,##0.00'
    ws["E5"].value = f"{len(df)} salas/medidores"

    # Limpar qualquer resquício de erro (#VALUE!) nas células F4:G5
    # (ocorre se o usuário selecionar 'Colocar na Célula' no Excel 365, que não é suportado pelo openpyxl)
    for r in (4, 5):
        for c in (6, 7):
            val_str = str(ws.cell(r, c).value or "").strip()
            if val_str.startswith("#"):
                ws.cell(r, c).value = None

    # Garantir que a logo esteja presente e perfeitamente centralizada nas 4 células F4:G5
    if len(ws._images) == 0:
        logo_path = _get_logo_path()
        if logo_path:
            try:
                img = XlImage(logo_path)
                img.width = 156
                img.height = 44
                ws.add_image(img, "F4")
                if hasattr(img, 'anchor') and hasattr(img.anchor, '_from'):
                    img.anchor._from.col = 5   # Coluna F
                    img.anchor._from.row = 3   # Linha 4
                    img.anchor._from.colOff = int(53 * 9525)   # Centraliza entre F e G
                    img.anchor._from.rowOff = int(7.33 * 9525) # Centraliza entre linhas 4 e 5
            except Exception as e:
                logging.warning(f"Erro ao reinserir logo no template: {e}")
    else:
        # Se já existe imagem no template, assegurar o posicionamento centralizado
        try:
            img = ws._images[0]
            if hasattr(img, 'anchor') and hasattr(img.anchor, '_from'):
                img.anchor._from.col = 5
                img.anchor._from.row = 3
                img.anchor._from.colOff = int(53 * 9525)
                img.anchor._from.rowOff = int(7.33 * 9525)
        except Exception:
            pass

    n_records = len(df)
    ROW_START = 8

    # Localizar dinamicamente a linha do TOTAL GERAL no modelo existente
    total_row_in_template = None
    for r in range(ROW_START, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip().upper()
        if "TOTAL GERAL" in v or v == "TOTAL":
            total_row_in_template = r
            break

    if total_row_in_template is None:
        total_row_in_template = ROW_START + n_records

    existing_records = total_row_in_template - ROW_START
    diff_records = n_records - existing_records

    # Se a quantidade de salas for diferente do modelo, ajusta as linhas mantendo a estrutura
    if diff_records > 0:
        new_merged_ranges = []
        for rng in list(ws.merged_cells.ranges):
            min_col, min_row, max_col, max_row = rng.bounds
            if min_row >= total_row_in_template:
                new_min_row = min_row + diff_records
                new_max_row = max_row + diff_records
                c1 = get_column_letter(min_col)
                c2 = get_column_letter(max_col)
                new_coord = f"{c1}{new_min_row}:{c2}{new_max_row}"
                ws.merged_cells.remove(rng)
                new_merged_ranges.append(new_coord)

        ws.insert_rows(total_row_in_template, amount=diff_records)

        for coord in new_merged_ranges:
            ws.merge_cells(coord)
    elif diff_records < 0:
        amount_to_delete = abs(diff_records)
        del_start = ROW_START + n_records
        ws.delete_rows(del_start, amount=amount_to_delete)

    total_row = ROW_START + n_records

    # Estilização no Padrão Executivo / Diretoria
    font_tbl_hdr= Font(name="Segoe UI", size=9, bold=True, color=CLR_TEXT_WHITE)
    fill_tbl_hdr= PatternFill(start_color=CLR_FOREST_DEEP, end_color=CLR_FOREST_DEEP, fill_type="solid")
    ws.row_dimensions[7].height = 32
    # Garantir larguras corretas das colunas (podem ser perdidas ao inserir/excluir linhas)
    col_widths_template = {"A": 28, "B": 14, "C": 16, "D": 18, "E": 18, "F": 18, "G": 18}
    for letter, width in col_widths_template.items():
        ws.column_dimensions[letter].width = width
    for c in range(1, 8):
        cell_h = ws.cell(row=7, column=c)
        _apply_cell(cell_h, font=font_tbl_hdr, fill=fill_tbl_hdr, border=Border())
        cell_h.alignment = Alignment(
            horizontal="left" if c == 1 else "right",
            vertical="center",
            wrap_text=True
        )

    # Fontes e preenchimentos dos dados executivos
    font_data   = Font(name="Segoe UI", size=9, color=CLR_TEXT_DARK)
    font_data_0 = Font(name="Segoe UI", size=9, color=CLR_TEXT_ZERO)
    font_total  = Font(name="Segoe UI", size=9, bold=True, color=CLR_FOREST_DEEP)

    fill_a      = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_b      = PatternFill(start_color=CLR_ZEBRA_LIGHT, end_color=CLR_ZEBRA_LIGHT, fill_type="solid")
    fill_total  = PatternFill(start_color=CLR_CARD_BG, end_color=CLR_CARD_BG, fill_type="solid")

    side_hairline = Side(style="thin", color=CLR_LINE_SUBTLE)
    border_row    = Border(top=side_hairline, bottom=side_hairline)
    border_total  = Border(
        top=Side(style="thin", color=CLR_FOREST_DEEP),
        bottom=Side(style="double", color=CLR_FOREST_DEEP)
    )
    al_left     = Alignment(horizontal="left", vertical="center")
    al_right    = Alignment(horizontal="right", vertical="center")

    # Inserir cada linha de medição com respiração e sem grades verticais
    for i, (_, r) in enumerate(df.iterrows()):
        curr_r = ROW_START + i
        ws.row_dimensions[curr_r].height = 21

        consumo = _safe_float(r.get("Consumo_m3"))
        valor   = _safe_float(r.get("Valor_RS"))
        ct11    = _safe_float(r.get("Consumo_Total_11m"))
        vt11    = _safe_float(r.get("Valor_Total_11m"))
        cm11    = _safe_float(r.get("Consumo_Medio_11m"))
        vm11    = _safe_float(r.get("Valor_Medio_11m"))

        is_zero = (consumo == 0 and valor == 0)
        f = font_data_0 if is_zero else font_data
        fill = fill_a if i % 2 == 0 else fill_b

        row_vals = [
            (str(r["Usuario"]), al_left, None),
            (consumo, al_right, '#,##0.00'),
            (valor, al_right, 'R$ #,##0.00'),
            (ct11, al_right, '#,##0.00'),
            (vt11, al_right, 'R$ #,##0.00'),
            (cm11, al_right, '#,##0.00'),
            (vm11, al_right, 'R$ #,##0.00'),
        ]

        for col_idx, (val, align, num_fmt) in enumerate(row_vals, 1):
            cell = ws.cell(row=curr_r, column=col_idx, value=val)
            _apply_cell(cell, font=f, fill=fill, alignment=align, border=border_row, number_format=num_fmt)

    # Linha Total Geral no Padrão Contábil Clássico (Double Underline)
    ws.row_dimensions[total_row].height = 24
    ws.cell(row=total_row, column=1, value="TOTAL GERAL")
    last_data_row = total_row - 1
    for c_idx in range(1, 8):
        cell = ws.cell(row=total_row, column=c_idx)
        if c_idx > 1:
            col_l = get_column_letter(c_idx)
            cell.value = f"=SUM({col_l}{ROW_START}:{col_l}{last_data_row})"
            cell.number_format = '#,##0.00' if c_idx in (2, 4, 6) else 'R$ #,##0.00'
        al = al_left if c_idx == 1 else al_right
        _apply_cell(cell, font=font_total, fill=fill_total, alignment=al, border=border_total)

    # Atualizar Fórmulas dos KPIs abaixo do total geral com visual limpo
    for r in range(total_row + 1, ws.max_row + 1):
        cell_a = ws.cell(row=r, column=1)
        if cell_a.value and str(cell_a.value).startswith("=B"):
            cell_a.value = f"=B{total_row}"
        cell_c = ws.cell(row=r, column=3)
        if cell_c.value and str(cell_c.value).startswith("=C"):
            cell_c.value = f"=C{total_row}"
        cell_e = ws.cell(row=r, column=5)
        if cell_e.value and "B" in str(cell_e.value) and "/" in str(cell_e.value):
            cell_e.value = f"=IF(B{total_row}=0,0,B{total_row}/{max(n_records, 1)})"

    # Atualizar filtro automático se existir
    ws.auto_filter.ref = f"A7:G{last_data_row}"

    # ── Criar ou atualizar a aba separada "Gráficos" ──
    try:
        if progress_callback:
            progress_callback(78, "3/4: Construindo gráficos comparativos e indicadores...")
        _build_graphics_sheet(wb, df, dt_inicio_str, dt_fim_str, valor_m3)
    except Exception as e:
        logging.warning(f"Não foi possível construir a aba de Gráficos: {e}")
        warnings.append(f"Os gráficos não puderam ser gerados: {e}")

    # Manter a aba 'Consumo e Rateio' como ativa ao abrir o arquivo
    wb.active = 0

    # Configurar impressão: paisagem, ajustado à largura, margens reduzidas
    _configure_print_settings(ws, last_data_row=total_row)

    if progress_callback:
        progress_callback(92, "4/4: Gravando arquivo Excel...")

    try:
        wb.save(output_path)
    except PermissionError:
        raise PermissionError(
            f"O arquivo '{os.path.basename(output_path)}' está aberto no Excel ou por outro programa.\n"
            f"Feche o arquivo e tente novamente."
        )

    # Registrar no histórico de relatórios recentes
    try:
        from core.config_manager import add_recent_report
        tot_m3 = df["Consumo_m3"].apply(_safe_float).sum() if "Consumo_m3" in df.columns else None
        tot_rs = df["Valor_RS"].apply(_safe_float).sum() if "Valor_RS" in df.columns else None
        per_str = f"{format_date_display(dt_inicio_str)} a {format_date_display(dt_fim_str)}"
        add_recent_report(output_path, periodo=per_str, total_m3=tot_m3, total_rs=tot_rs)
    except Exception as e:
        logging.warning(f"Não foi possível salvar no histórico recente: {e}")

    return output_path, warnings


def export_to_pdf(excel_path, pdf_path=None):
    """
    Converte uma planilha Excel (.xlsx) para PDF com fidelidade total usando o Microsoft Excel via COM.
    Exclui a aba 'Gráficos' da impressão pois não renderiza bem em PDF.
    Retorna o caminho do PDF gerado ou None se o Excel não estiver disponível ou ocorrer erro.
    """
    if not excel_path or not os.path.exists(excel_path):
        return None

    if not pdf_path:
        pdf_path = os.path.splitext(excel_path)[0] + ".pdf"

    abs_excel = os.path.abspath(excel_path)
    abs_pdf = os.path.abspath(pdf_path)

    try:
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(abs_excel)

            # Localizar a planilha principal de medição (ignora qualquer aba com gráfico)
            target_sheet = None
            for i in range(1, wb.Sheets.Count + 1):
                s_name = str(wb.Sheets(i).Name).lower()
                # Descarta abas de gráficos mesmo com variações de acentuação/encoding
                if "graf" not in s_name and "chart" not in s_name:
                    target_sheet = wb.Sheets(i)
                    break

            if target_sheet is None:
                target_sheet = wb.Sheets(1)

            # Exportar APENAS a planilha de dados selecionada para o PDF
            target_sheet.ExportAsFixedFormat(0, abs_pdf)

            wb.Close(False)
            logging.info(f"PDF gerado com sucesso em: {abs_pdf}")
            return abs_pdf
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()
    except Exception as e:
        logging.warning(f"Não foi possível gerar versão em PDF via Excel: {e}")
        return None


def generate_excel_report(df, dt_inicio_str, dt_fim_str, valor_m3, output_path, sort_by_consumption=True, progress_callback=None):
    """
    Função principal: se existir o modelo_relatorio.xlsx personalizado, usa ele.
    Caso contrário, gera o relatório completo do zero com o design padrão CompaSSS.
    Gera automaticamente a versão em Excel (.xlsx) e uma cópia executiva em PDF (.pdf).
    Retorna (output_path, warnings).
    """
    if sort_by_consumption:
        df = sort_dataframe_by_consumption(df)

    tmpl = get_template_path()
    res_path = None
    warnings = []

    if tmpl and os.path.exists(tmpl):
        try:
            res_path, warnings = generate_excel_from_template(tmpl, df, dt_inicio_str, dt_fim_str, valor_m3, output_path, sort_by_consumption=False, progress_callback=progress_callback)
        except PermissionError:
            raise
        except Exception as e:
            logging.warning(f"Não foi possível aplicar o modelo {tmpl}: {e}. Gerando com renderizador padrão.")

    if res_path is None:
        res_path, warnings = _generate_excel_full_code(df, dt_inicio_str, dt_fim_str, valor_m3, output_path, sort_by_consumption=False, progress_callback=progress_callback)

    # Gerar cópia executiva em PDF automaticamente na mesma pasta
    try:
        if progress_callback:
            progress_callback(95, "5/5: Gerando cópia executiva em PDF...")
        pdf_res = export_to_pdf(res_path)
        if not pdf_res:
            warnings.append("Aviso: A versão em PDF não pôde ser gerada automaticamente (Excel indisponível).")
    except Exception as e:
        logging.warning(f"Erro ao gerar cópia em PDF: {e}")

    return res_path, warnings


def create_default_template_file(template_path):
    """Cria um arquivo modelo_relatorio.xlsx oficial baseado no mês de setembro com todos os dados reais."""
    candidates = [
        os.path.join(os.path.expanduser("~"), "Desktop", "Rateio de água - Setembro.xlsx"),
        os.path.join(os.path.expanduser("~"), "Desktop", "Relatorios", "Rateio de água - Setembro.xlsx"),
        os.path.join(os.path.expanduser("~"), "Desktop", "Rateio_Agua_2026_09.xlsx"),
        os.path.join(os.path.expanduser("~"), "Desktop", "Relatorios", "Rateio_Agua_2026_09.xlsx"),
    ]
    for c in candidates:
        if os.path.exists(c):
            shutil.copy2(c, template_path)
            return template_path

    # Se não houver arquivo pronto, busca no banco o ciclo de setembro e renderiza completo
    try:
        from core.database import fetch_hidrometros_data
        df_full = fetch_hidrometros_data("2026-08-29 00:00:00", "2026-09-28 23:59:59", 63.68)
        _generate_excel_full_code(df_full, "2026-08-29 00:00:00", "2026-09-28 23:59:59", 63.68, template_path, include_chart=True, sort_by_consumption=True)
    except Exception:
        import pandas as pd
        df_sample = pd.DataFrame([{
            "Usuario": "Exemplo Sala 101",
            "Consumo_m3": 10.5,
            "Valor_RS": 668.64,
            "Consumo_Total_11m": 115.5,
            "Valor_Total_11m": 7355.04,
            "Consumo_Medio_11m": 10.5,
            "Valor_Medio_11m": 668.64
        }])
        _generate_excel_full_code(df_sample, "2026-08-29", "2026-09-28", 63.68, template_path, include_chart=True, sort_by_consumption=False)
    return template_path


def _generate_excel_full_code(df, dt_inicio_str, dt_fim_str, valor_m3, output_path, include_chart=True, sort_by_consumption=False, progress_callback=None):
    """
    Renderizador completo do zero com paleta CompaSSS e logotipo.
    Retorna (output_path, warnings).
    """
    warnings = []

    if progress_callback:
        progress_callback(60, "3/4: Formatando dados na planilha executiva...")

    if sort_by_consumption:
        df = sort_dataframe_by_consumption(df)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Consumo e Rateio"
    ws.sheet_properties.tabColor = CLR_GREEN_MED
    ws.views.sheetView[0].showGridLines = False

    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_options.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = "landscape"

    # Fontes Executivas
    font_title    = Font(name="Segoe UI", size=15, bold=True, color=CLR_FOREST_DEEP)
    font_subtitle = Font(name="Segoe UI", size=9, color=CLR_TEXT_MUTED)
    font_label    = Font(name="Segoe UI", size=8, bold=True, color=CLR_TEXT_MUTED)
    font_val      = Font(name="Segoe UI", size=10, bold=True, color=CLR_TEXT_DARK)
    font_header   = Font(name="Segoe UI", size=9, bold=True, color=CLR_TEXT_WHITE)
    font_data     = Font(name="Segoe UI", size=9, color=CLR_TEXT_DARK)
    font_data_m   = Font(name="Segoe UI", size=9, color=CLR_TEXT_ZERO)
    font_total    = Font(name="Segoe UI", size=10, bold=True, color=CLR_FOREST_DEEP)
    font_total_v  = Font(name="Segoe UI", size=10, bold=True, color=CLR_FOREST_DEEP)
    font_footer   = Font(name="Segoe UI", size=8, italic=True, color=CLR_TEXT_MUTED)
    font_stats_t  = Font(name="Segoe UI", size=8, bold=True, color=CLR_TEXT_MUTED)
    font_stats_v  = Font(name="Segoe UI", size=12, bold=True, color=CLR_FOREST_DEEP)

    # Preenchimentos
    fill_title_bar = PatternFill(start_color="FFFFFF",      end_color="FFFFFF",      fill_type="solid")
    fill_header    = PatternFill(start_color=CLR_FOREST_DEEP, end_color=CLR_FOREST_DEEP, fill_type="solid")
    fill_zebra_a   = PatternFill(start_color="FFFFFF",      end_color="FFFFFF",      fill_type="solid")
    fill_zebra_b   = PatternFill(start_color=CLR_ZEBRA_LIGHT, end_color=CLR_ZEBRA_LIGHT, fill_type="solid")
    fill_total     = PatternFill(start_color=CLR_CARD_BG,   end_color=CLR_CARD_BG,   fill_type="solid")
    fill_meta_bg   = PatternFill(start_color=CLR_CARD_BG,   end_color=CLR_CARD_BG,   fill_type="solid")
    fill_kpi_bg    = PatternFill(start_color=CLR_CARD_BG,   end_color=CLR_CARD_BG,   fill_type="solid")

    # Bordas
    side_subtle   = Side(style="thin", color=CLR_LINE_SUBTLE)
    border_data   = Border(top=side_subtle, bottom=side_subtle)
    border_header = Border(left=Side(style=None), right=Side(style=None),
                           top=side_subtle,
                           bottom=Side(style="medium", color=CLR_ACCENT_GREEN))
    border_total  = Border(top=Side(style="thin", color=CLR_FOREST_DEEP),
                           bottom=Side(style="double", color=CLR_FOREST_DEEP))
    border_meta   = Border(left=side_subtle, right=side_subtle, top=side_subtle, bottom=side_subtle)

    al_left   = Alignment(horizontal="left",   vertical="center")
    al_right  = Alignment(horizontal="right",  vertical="center")
    al_center = Alignment(horizontal="center", vertical="center")
    al_wrap   = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. Faixa de Título
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 18

    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = "HIDRÔMETROS — PRAÇA PAMPLONA"
    _apply_cell(title_cell, font=font_title, fill=fill_title_bar,
                alignment=Alignment(horizontal="left", vertical="center", indent=1))

    for col in range(1, 8):
        _apply_cell(ws.cell(row=1, column=col), fill=fill_title_bar)
        _apply_cell(ws.cell(row=2, column=col), fill=fill_title_bar)

    ws.merge_cells("A2:G2")
    sub_cell = ws["A2"]
    sub_cell.value = "Consumo de Água e Rateio por Sala  •  Sistema StruxureWare EBO"
    _apply_cell(sub_cell, font=font_subtitle, fill=fill_title_bar,
                alignment=Alignment(horizontal="left", vertical="top", indent=1))

    # 2. Metadados e Logo (F4:G5)
    ws.row_dimensions[3].height = 6
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 22

    meta_fields = [
        (4, "A", "Período:",        f"{format_date_display(dt_inicio_str)}  a  {format_date_display(dt_fim_str)}"),
        (4, "D", "Data Emissão:",   datetime.now().strftime('%d/%m/%Y %H:%M')),
        (5, "A", "Valor do m³:",    None),
        (5, "D", "Registros:",      f"{len(df)} salas/medidores"),
    ]

    for row_n, col_l, label, value in meta_fields:
        cell_label = ws[f"{col_l}{row_n}"]
        cell_label.value = label
        _apply_cell(cell_label, font=font_label, fill=fill_meta_bg,
                    alignment=al_left, border=border_meta)

        next_col = chr(ord(col_l) + 1)
        cell_val = ws[f"{next_col}{row_n}"]
        if value is not None:
            cell_val.value = value
        _apply_cell(cell_val, font=font_val, fill=fill_meta_bg,
                    alignment=al_left, border=border_meta)

    ws["B5"].value = float(valor_m3)
    ws["B5"].number_format = 'R$ #,##0.00'

    for r in (4, 5):
        for c in range(1, 8):
            cell = ws.cell(row=r, column=c)
            if cell.fill == PatternFill():
                cell.fill = fill_meta_bg

    # Inserir Logo centralizada no bloco de 4 células F4:G5
    logo_path = _get_logo_path()
    if logo_path:
        try:
            img = XlImage(logo_path)
            img.width = 156
            img.height = 44
            ws.add_image(img, "F4")
            if hasattr(img, 'anchor') and hasattr(img.anchor, '_from'):
                img.anchor._from.col = 5   # Coluna F
                img.anchor._from.row = 3   # Linha 4
                img.anchor._from.colOff = int(53 * 9525)   # Centraliza entre colunas F e G
                img.anchor._from.rowOff = int(7.33 * 9525) # Centraliza entre linhas 4 e 5
        except Exception:
            pass

    # 3. Cabeçalho da Tabela
    ws.row_dimensions[6].height = 6
    headers = [
        "Sala / Usuário",
        "Consumo\n(m³)",
        "Valor\n(R$)",
        "Consumo Total\n11 meses (m³)",
        "Valor Total\n11 meses (R$)",
        "Consumo Médio\n11 meses (m³)",
        "Valor Médio\n11 meses (R$)"
    ]

    ROW_HEADER = 7
    ws.row_dimensions[ROW_HEADER].height = 30
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=ROW_HEADER, column=col_idx, value=h)
        _apply_cell(c, font=font_header, fill=fill_header, alignment=al_wrap, border=border_header)

    # 4. Dados
    ROW_DATA_START = 8
    curr_row = ROW_DATA_START

    for idx, r in df.iterrows():
        ws.row_dimensions[curr_row].height = 21
        consumo = _safe_float(r.get("Consumo_m3"))
        valor   = _safe_float(r.get("Valor_RS"))
        ct11    = _safe_float(r.get("Consumo_Total_11m"))
        vt11    = _safe_float(r.get("Valor_Total_11m"))
        cm11    = _safe_float(r.get("Consumo_Medio_11m"))
        vm11    = _safe_float(r.get("Valor_Medio_11m"))

        is_zero = (consumo == 0 and valor == 0)
        row_font = font_data_m if is_zero else font_data
        row_fill = fill_zebra_a if (curr_row - ROW_DATA_START) % 2 == 0 else fill_zebra_b

        ws.cell(row=curr_row, column=1, value=str(r["Usuario"]))
        ws.cell(row=curr_row, column=2, value=consumo)
        ws.cell(row=curr_row, column=3, value=valor)
        ws.cell(row=curr_row, column=4, value=ct11)
        ws.cell(row=curr_row, column=5, value=vt11)
        ws.cell(row=curr_row, column=6, value=cm11)
        ws.cell(row=curr_row, column=7, value=vm11)

        for c_idx in range(1, 8):
            cell = ws.cell(row=curr_row, column=c_idx)
            al = al_left if c_idx == 1 else al_right
            _apply_cell(cell, font=row_font, fill=row_fill, alignment=al, border=border_data)
            if c_idx in (2, 4, 6):
                cell.number_format = '#,##0.00'
            elif c_idx in (3, 5, 7):
                cell.number_format = 'R$ #,##0.00'

        curr_row += 1

    last_data_row = curr_row - 1

    # 5. Total Geral
    ws.row_dimensions[curr_row].height = 26
    ws.cell(row=curr_row, column=1, value="TOTAL GERAL")

    if last_data_row >= ROW_DATA_START:
        for c_idx in range(2, 8):
            col_let = get_column_letter(c_idx)
            ws.cell(row=curr_row, column=c_idx,
                    value=f"=SUM({col_let}{ROW_DATA_START}:{col_let}{last_data_row})")
    else:
        for c_idx in range(2, 8):
            ws.cell(row=curr_row, column=c_idx, value=0.0)

    for c_idx in range(1, 8):
        cell = ws.cell(row=curr_row, column=c_idx)
        al = al_left if c_idx == 1 else al_right
        f = font_total if c_idx == 1 else font_total_v
        _apply_cell(cell, font=f, fill=fill_total, alignment=al, border=border_total)
        if c_idx in (2, 4, 6):
            cell.number_format = '#,##0.00'
        elif c_idx in (3, 5, 7):
            cell.number_format = 'R$ #,##0.00'

    total_row = curr_row
    curr_row += 1

    # 6. KPIs
    curr_row += 1
    ws.row_dimensions[curr_row].height = 4
    for c in range(1, 8):
        fill_accent = PatternFill(start_color=CLR_GREEN_BRAND, end_color=CLR_GREEN_BRAND, fill_type="solid")
        _apply_cell(ws.cell(row=curr_row, column=c), fill=fill_accent)
    curr_row += 1

    ws.row_dimensions[curr_row].height = 34
    kpi_row = curr_row

    kpi_labels = [
        ("A", "Consumo Total"),
        ("C", "Valor Total"),
        ("E", "Média por Sala"),
    ]
    num_salas = max(len(df), 1)

    for col_l, label in kpi_labels:
        cell_l = ws[f"{col_l}{kpi_row}"]
        cell_l.value = label
        _apply_cell(cell_l, font=font_stats_t, fill=fill_kpi_bg, alignment=al_left)

    next_row = kpi_row + 1
    ws.row_dimensions[next_row].height = 28

    ws[f"A{next_row}"].value = f"=B{total_row}"
    _apply_cell(ws[f"A{next_row}"], font=font_stats_v, fill=fill_kpi_bg,
                alignment=al_left, number_format='#,##0.00 "m³"')

    ws[f"C{next_row}"].value = f"=C{total_row}"
    _apply_cell(ws[f"C{next_row}"], font=font_stats_v, fill=fill_kpi_bg,
                alignment=al_left, number_format='R$ #,##0.00')

    ws[f"E{next_row}"].value = f"=IF(B{total_row}=0,0,B{total_row}/{num_salas})"
    _apply_cell(ws[f"E{next_row}"], font=font_stats_v, fill=fill_kpi_bg,
                alignment=al_left, number_format='#,##0.00 "m³/sala"')

    for r in (kpi_row, next_row):
        for c in range(1, 8):
            cell = ws.cell(row=r, column=c)
            if cell.fill == PatternFill():
                cell.fill = fill_kpi_bg

    curr_row = next_row + 1

    # 7. Aba Separada de Gráficos e Indicadores
    if include_chart:
        try:
            if progress_callback:
                progress_callback(78, "3/4: Construindo gráficos comparativos e indicadores...")
            _build_graphics_sheet(wb, df, dt_inicio_str, dt_fim_str, valor_m3)
        except Exception as e:
            logging.warning(f"Não foi possível construir a aba de Gráficos: {e}")
            warnings.append(f"Os gráficos não puderam ser gerados: {e}")
        wb.active = 0

    # 8. Rodapé
    curr_row += 1
    ws.merge_cells(f"A{curr_row}:G{curr_row}")
    footer = ws.cell(row=curr_row, column=1,
                     value=f"Relatório gerado automaticamente em {datetime.now().strftime('%d/%m/%Y às %H:%M')}  •  "
                           f"Fonte: StruxureWare EBO  •  Condomínio Praça Pamplona  •  CompaSSS")
    _apply_cell(footer, font=font_footer, alignment=al_center)

    col_widths = {
        "A": 28, "B": 14, "C": 16,
        "D": 18, "E": 18, "F": 18, "G": 18,
    }
    for letter, width in col_widths.items():
        ws.column_dimensions[letter].width = width

    ws.freeze_panes = f"A{ROW_DATA_START}"
    ws.auto_filter.ref = f"A{ROW_HEADER}:G{last_data_row}"

    # Configurar impressão: paisagem, ajustado à largura, margens reduzidas
    _configure_print_settings(ws, last_data_row=last_data_row)

    if progress_callback:
        progress_callback(92, "4/4: Gravando arquivo Excel...")

    try:
        wb.save(output_path)
    except PermissionError:
        raise PermissionError(
            f"O arquivo '{os.path.basename(output_path)}' está aberto no Excel ou por outro programa.\n"
            f"Feche o arquivo e tente novamente."
        )

    # Registrar no histórico de relatórios recentes
    try:
        from core.config_manager import add_recent_report
        tot_m3 = df["Consumo_m3"].apply(_safe_float).sum() if "Consumo_m3" in df.columns else None
        tot_rs = df["Valor_RS"].apply(_safe_float).sum() if "Valor_RS" in df.columns else None
        per_str = f"{format_date_display(dt_inicio_str)} a {format_date_display(dt_fim_str)}"
        add_recent_report(output_path, periodo=per_str, total_m3=tot_m3, total_rs=tot_rs)
    except Exception as e:
        logging.warning(f"Não foi possível salvar no histórico recente: {e}")

    return output_path, warnings
