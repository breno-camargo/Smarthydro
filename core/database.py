import pyodbc
import pandas as pd
import logging
from core.config_manager import load_config

def get_available_odbc_drivers():
    """Retorna lista dos drivers ODBC disponíveis no sistema operacional."""
    try:
        return pyodbc.drivers()
    except Exception as e:
        logging.warning(f"Não foi possível listar drivers ODBC via pyodbc: {e}")
        return []

def select_best_odbc_driver(preferred_driver=None):
    """Seleciona o driver ODBC mais adequado disponível no Windows."""
    installed = get_available_odbc_drivers()
    
    if preferred_driver and preferred_driver in installed:
        return preferred_driver

    candidates = [
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "ODBC Driver 11 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server"
    ]

    for cand in candidates:
        if cand in installed:
            return cand

    return preferred_driver or "SQL Server"

def build_connection_string(config=None):
    """Monta a string de conexão ODBC com base nas configurações."""
    if config is None:
        config = load_config()

    server = config.get("server", "localhost\\SQLEXPRESS")
    database = config.get("database", "StruxureWareReportsDB")
    preferred = config.get("odbc_driver", "ODBC Driver 17 for SQL Server")
    driver = select_best_odbc_driver(preferred)

    conn_parts = [
        f"Driver={{{driver}}}",
        f"Server={server}",
        f"Database={database}",
    ]

    if config.get("trusted_connection", True):
        conn_parts.append("Trusted_Connection=yes")
    else:
        user = config.get("db_user", "")
        pwd = config.get("db_password", "")
        conn_parts.append(f"UID={user}")
        conn_parts.append(f"PWD={pwd}")

    # ODBC 18 exige TrustServerCertificate por padrão
    if "ODBC Driver 18" in driver:
        conn_parts.append("TrustServerCertificate=yes")

    return ";".join(conn_parts) + ";", driver

def test_db_connection(config=None):
    """Testa a conexão com o banco de dados SQL Server."""
    conn_str, driver = build_connection_string(config)
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        cursor = conn.cursor()
        cursor.execute("SELECT @@VERSION")
        version_row = cursor.fetchone()
        version = version_row[0].split("\n")[0] if version_row else "SQL Server OK"
        conn.close()
        return True, f"Conexão estabelecida com sucesso via [{driver}].\n{version}", driver
    except Exception as e:
        return False, f"Falha na conexão usando [{driver}]:\n{str(e)}", driver

def fetch_hidrometros_data(dt_inicio, dt_fim, valor_m3, config=None, progress_callback=None):
    """
    Executa a consulta SQL no banco StruxureWareReportsDB e retorna um DataFrame com os consumos.
    dt_inicio e dt_fim devem estar no formato 'YYYY-MM-DD HH:MM:SS'
    """
    # Normalização defensiva: se vier apenas a data YYYY-MM-DD, adiciona o horário completo
    dt_ini_str = str(dt_inicio).strip()
    dt_fim_str = str(dt_fim).strip()
    if len(dt_ini_str) == 10:
        dt_ini_str = f"{dt_ini_str} 00:00:00"
    if len(dt_fim_str) == 10:
        dt_fim_str = f"{dt_fim_str} 23:59:59"

    conn_str, driver = build_connection_string(config)
    
    sql_query = """
    SET NOCOUNT ON;
    SET DATEFORMAT ymd;

    DECLARE @INICIO DATETIME = CONVERT(DATETIME, ?, 120);
    DECLARE @FIM    DATETIME = CONVERT(DATETIME, ?, 120);
    DECLARE @valor_m3 DECIMAL(10,2) = ?;
    DECLARE @Name   VARCHAR(150) = '%';

    IF OBJECT_ID('tempdb..#TempHidroMedia') IS NOT NULL DROP TABLE #TempHidroMedia;

    SELECT 
        Name,
        REPLACE(Name, 'Sala 1007_', 'Sala 1007') as Name2,
        CAST((SELECT TOP(1) FloatVALUE FROM tbLogTimeValues WHERE tbLogTimeValues.ParentID = tbTrendLogRelation.EntityID AND DateTimeStamp BETWEEN @INICIO AND @FIM ORDER BY DateTimeStamp) AS DECIMAL(18,2)) AS LEITURA_INICIAL,
        (SELECT TOP(1) DateTimeStamp FROM tbLogTimeValues WHERE tbLogTimeValues.ParentID = tbTrendLogRelation.EntityID AND DateTimeStamp BETWEEN @INICIO AND @FIM ORDER BY DateTimeStamp) AS HORA_LEITURA_INICIAL,
        CAST((SELECT TOP(1) FloatVALUE FROM tbLogTimeValues WHERE tbLogTimeValues.ParentID = tbTrendLogRelation.EntityID AND DateTimeStamp BETWEEN @INICIO AND @FIM ORDER BY DateTimeStamp DESC) AS DECIMAL(18,2)) AS LEITURA_FINAL,
        (SELECT TOP(1) DateTimeStamp FROM tbLogTimeValues WHERE tbLogTimeValues.ParentID = tbTrendLogRelation.EntityID AND DateTimeStamp BETWEEN @INICIO AND @FIM ORDER BY DateTimeStamp DESC) AS HORA_LEITURA_FINAL,

        CAST((SELECT TOP(1) FloatVALUE FROM tbLogTimeValues WHERE tbLogTimeValues.ParentID = tbTrendLogRelation.EntityID AND DateTimeStamp BETWEEN DATEADD(day,-365,@INICIO) AND @INICIO ORDER BY DateTimeStamp) AS DECIMAL(18,2)) AS LEITURA_INICIAL_ANO,
        (SELECT TOP(1) DateTimeStamp FROM tbLogTimeValues WHERE tbLogTimeValues.ParentID = tbTrendLogRelation.EntityID AND DateTimeStamp BETWEEN DATEADD(day,-365,@INICIO) AND @INICIO ORDER BY DateTimeStamp) AS HORA_LEITURA_INICIAL_ANO,
        CAST((SELECT TOP(1) FloatVALUE FROM tbLogTimeValues WHERE tbLogTimeValues.ParentID = tbTrendLogRelation.EntityID AND DateTimeStamp BETWEEN DATEADD(day,-365,@INICIO) AND @INICIO ORDER BY DateTimeStamp DESC) AS DECIMAL(18,2)) AS LEITURA_FINAL_ANO,
        (SELECT TOP(1) DateTimeStamp FROM tbLogTimeValues WHERE tbLogTimeValues.ParentID = tbTrendLogRelation.EntityID AND DateTimeStamp BETWEEN DATEADD(day,-365,@INICIO) AND @INICIO ORDER BY DateTimeStamp DESC) AS HORA_LEITURA_FINAL_ANO
    INTO #TempHidroMedia
    FROM tbTrendLogRelation
    WHERE Type = 'trend.ETLog'
      AND Name NOT LIKE '%Sala 1001-1006%'
      AND CAST(tbtrendlogrelation.name AS VARCHAR(150)) LIKE CASE WHEN @Name = '%' THEN CAST(tbtrendlogrelation.name AS VARCHAR(150)) ELSE @Name END;

    SELECT
        Name AS [Usuario], 
        CASE 
            WHEN Name LIKE 'Sala Com Problema' THEN 1.0
            ELSE ((LEITURA_FINAL - LEITURA_INICIAL)/1000.0) 
        END AS [Consumo_m3], 
        
        CASE 
            WHEN Name LIKE 'Sala Com Problema' THEN @valor_m3
            ELSE ((LEITURA_FINAL - LEITURA_INICIAL) * @valor_m3 / 1000.0)
        END AS [Valor_RS],

        CASE 
            WHEN LEITURA_FINAL_ANO < LEITURA_INICIAL_ANO THEN 0
            ELSE ((LEITURA_FINAL_ANO - LEITURA_INICIAL_ANO)/1000.0) 
        END AS [Consumo_Total_11m],
        
        CASE 
            WHEN LEITURA_FINAL_ANO < LEITURA_INICIAL_ANO THEN 0
            ELSE (((LEITURA_FINAL_ANO - LEITURA_INICIAL_ANO) * @valor_m3 / 1000.0))
        END AS [Valor_Total_11m],

        CASE 
            WHEN LEITURA_FINAL_ANO < LEITURA_INICIAL_ANO THEN 0
            ELSE (((LEITURA_FINAL_ANO - LEITURA_INICIAL_ANO)/1000.0)/11.0)
        END AS [Consumo_Medio_11m],

        CASE 
            WHEN LEITURA_FINAL_ANO < LEITURA_INICIAL_ANO THEN 0
            ELSE ((((LEITURA_FINAL_ANO - LEITURA_INICIAL_ANO) * @valor_m3 / 1000.0))/11.0)
        END AS [Valor_Medio_11m]

    FROM #TempHidroMedia
    WHERE Name NOT IN ('Sala 1007')
      AND Name NOT LIKE '%Sala 1001 - 1006%'
      AND Name2 NOT LIKE '%Sala 1001 - 1006%'
      AND [Name] NOT LIKE 'Ano_Final_pos_bug'
      AND [Name] NOT LIKE 'Ano_Inicial'
      AND [Name] NOT LIKE 'Dia_Final'
      AND [Name] NOT LIKE 'Dia_Inicial'
      AND [Name] NOT LIKE 'Energia Extended Trend Log'
      AND [Name] NOT LIKE 'Extended Trend Log'
      AND [Name] NOT LIKE 'KWh Extended Trend Log'
      AND [Name] NOT LIKE 'Mes_Final%'
      AND [Name] NOT LIKE 'Mes_Inicial'
      AND [Name] NOT LIKE 'valor_m3'
      AND [Name2] NOT LIKE 'Ano_Final_pos_bug'
      AND [Name2] NOT LIKE 'Ano_Inicial'
      AND [Name2] NOT LIKE 'Dia_Final'
      AND [Name2] NOT LIKE 'Dia_Inicial'
      AND [Name2] NOT LIKE 'Energia Extended Trend Log'
      AND [Name2] NOT LIKE 'Extended Trend Log'
      AND [Name2] NOT LIKE 'KWh Extended Trend Log'
      AND [Name2] NOT LIKE 'Mes_Final%'
      AND [Name2] NOT LIKE 'Mes_Inicial'
      AND [Name2] NOT LIKE 'valor_m3'
    ORDER BY [Name] ASC;

    DROP TABLE #TempHidroMedia;
    """

    if progress_callback:
        progress_callback(15, "1/4: Conectando ao banco de dados StruxureWare...")

    conn = pyodbc.connect(conn_str, timeout=30)
    try:
        if progress_callback:
            progress_callback(35, "2/4: Consultando histórico e medições de hidrômetros...")
        cursor = conn.cursor()
        cursor.execute(sql_query, (dt_ini_str, dt_fim_str, float(valor_m3)))
        rows = cursor.fetchall()
        col_names = [column[0] for column in cursor.description]
        df = pd.DataFrame.from_records(rows, columns=col_names)
        return df
    finally:
        conn.close()
