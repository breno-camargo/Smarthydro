import os
import sys
import argparse
import logging
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta

from core.config_manager import load_config, get_base_dir, format_report_filename, get_report_output_folder
from core.database import fetch_hidrometros_data, test_db_connection
from core.report_generator import generate_excel_report

def setup_cli_logging():
    log_file = os.path.join(get_base_dir(), "execucao.log")
    handlers = [logging.FileHandler(log_file, encoding="utf-8")]
    if sys.stdout is not None:
        handlers.append(logging.StreamHandler(sys.stdout))
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=handlers,
        force=True
    )

def get_billing_cycle_dates(cycle_type="ciclo_29_28", ref_date=None):
    """
    Calcula as datas inicial e final do ciclo.
    - ciclo_29_28: dia 29 do mês anterior até o dia 28 do mês de referência.
    - mes_civil: dia 1º ao último dia do mês anterior.
    """
    if ref_date is None:
        ref_date = date.today()

    if cycle_type == "ciclo_29_28":
        # Se hoje for antes do dia 28, o ciclo fechado mais recente terminou no dia 28 do mês passado.
        if ref_date.day < 28:
            end_month = ref_date - relativedelta(months=1)
        else:
            end_month = ref_date

        d_fim = date(end_month.year, end_month.month, 28)
        start_month = end_month - relativedelta(months=1)
        d_ini = date(start_month.year, start_month.month, 29)
    else:
        # Mês civil anterior fechado
        first_day_current_month = date(ref_date.year, ref_date.month, 1)
        d_fim = first_day_current_month - timedelta(days=1)
        d_ini = date(d_fim.year, d_fim.month, 1)

    dt_inicio_str = f"{d_ini.strftime('%Y-%m-%d')} 00:00:00"
    dt_fim_str = f"{d_fim.strftime('%Y-%m-%d')} 23:59:59"
    return dt_inicio_str, dt_fim_str, d_ini, d_fim

def execute_extraction(dt_inicio, dt_fim, valor_m3=None, output_path=None, config=None, sort_by_consumption=None, progress_callback=None):
    """
    Executa a extração dos dados e salva a planilha Excel.
    Retorna (output_path, warnings) onde warnings é uma lista de avisos não-fatais.
    """
    if config is None:
        config = load_config()

    if valor_m3 is None:
        valor_m3 = float(config.get("default_m3_price", 63.68))

    if valor_m3 <= 0:
        raise ValueError(f"O valor do m³ deve ser positivo. Valor recebido: R$ {valor_m3:.2f}")

    if sort_by_consumption is None:
        sort_by_consumption = bool(config.get("sort_by_consumption", True))

    logging.info(f"Iniciando extração: Período {dt_inicio} até {dt_fim} | Valor/m³: R$ {valor_m3:.2f} | Ordenação por consumo: {sort_by_consumption}")

    if progress_callback:
        progress_callback(10, "1/4: Conectando ao SQL Server (StruxureWare)...")

    df = fetch_hidrometros_data(dt_inicio, dt_fim, valor_m3, config, progress_callback=progress_callback)
    logging.info(f"Dados obtidos com sucesso do SQL Server: {len(df)} registros encontrados.")

    if output_path is None:
        base_out = config.get("output_directory", os.path.join(get_base_dir(), "relatorios"))
        target_dir = get_report_output_folder(base_out, dt_fim)
        # Identificador padronizado: "Rateio de água - Junho.xlsx"
        filename = format_report_filename(dt_fim)
        output_path = os.path.join(target_dir, filename)

    if progress_callback:
        progress_callback(55, "3/4: Formatando planilha e gráficos de consumo...")

    result_path, warnings = generate_excel_report(df, dt_inicio, dt_fim, valor_m3, output_path, sort_by_consumption=sort_by_consumption, progress_callback=progress_callback)
    logging.info(f"Relatório Excel gravado com sucesso em: {result_path}")
    if warnings:
        for w in warnings:
            logging.warning(f"Aviso: {w}")

    if progress_callback:
        progress_callback(100, f"Relatório gerado com sucesso: {os.path.basename(result_path)}")

    return result_path, warnings

def run_cli():
    setup_cli_logging()
    config = load_config()

    parser = argparse.ArgumentParser(description="Automação de Relatórios de Hidrômetros - Praça Pamplona")
    parser.add_argument("--auto", action="store_true", help="Execução automática do último ciclo de faturamento fechado")
    parser.add_argument("--test-connection", action="store_true", help="Testa a conexão com o banco de dados e sai")
    parser.add_argument("--inicio", type=str, help="Data inicial (formato DD/MM/AAAA ou YYYY-MM-DD)")
    parser.add_argument("--fim", type=str, help="Data final (formato DD/MM/AAAA ou YYYY-MM-DD)")
    parser.add_argument("--valor", type=float, help="Valor do m³ em Reais (ex: 63.68)")
    parser.add_argument("--saida", type=str, help="Caminho do arquivo Excel de saída")
    parser.add_argument("--alfabetico", action="store_true", help="Ordena salas em ordem alfabética em vez de maior consumo")

    args = parser.parse_args()
    sort_by_consumption = False if args.alfabetico else None

    if args.test_connection:
        ok, msg, drv = test_db_connection(config)
        if ok:
            logging.info(f"Sucesso: {msg}")
            sys.exit(0)
        else:
            logging.error(f"Erro: {msg}")
            sys.exit(1)

    if args.auto:
        cycle_type = config.get("billing_cycle_type", "ciclo_29_28")
        dt_ini, dt_fim, _, _ = get_billing_cycle_dates(cycle_type)
        val = args.valor or float(config.get("default_m3_price", 63.68))
        try:
            out_file, warnings = execute_extraction(dt_ini, dt_fim, val, args.saida, config, sort_by_consumption=sort_by_consumption)
            logging.info(f"[SUCESSO] Relatório gerado: {out_file}")
            if warnings:
                for w in warnings:
                    logging.warning(f"[AVISO] {w}")
            if sys.stdout:
                print(f"[SUCESSO] Relatório gerado: {out_file}")
            sys.exit(0)
        except Exception as e:
            logging.exception(f"Falha na extração automática: {e}")
            sys.exit(1)

    if args.inicio and args.fim:
        def parse_date(d_str, is_end=False):
            for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    dt = datetime.strptime(d_str, fmt)
                    time_part = "23:59:59" if is_end else "00:00:00"
                    return f"{dt.strftime('%Y-%m-%d')} {time_part}"
                except ValueError:
                    pass
            raise ValueError(f"Formato de data inválido: {d_str}")

        dt_ini = parse_date(args.inicio, is_end=False)
        dt_fim = parse_date(args.fim, is_end=True)
        val = args.valor or float(config.get("default_m3_price", 63.68))
        try:
            out_file, warnings = execute_extraction(dt_ini, dt_fim, val, args.saida, config, sort_by_consumption=sort_by_consumption)
            logging.info(f"[SUCESSO] Relatório gerado: {out_file}")
            if warnings:
                for w in warnings:
                    logging.warning(f"[AVISO] {w}")
            if sys.stdout:
                print(f"[SUCESSO] Relatório gerado: {out_file}")
            sys.exit(0)
        except Exception as e:
            logging.exception(f"Falha na extração por linha de comando: {e}")
            sys.exit(1)

    if sys.stdout:
        parser.print_help()
