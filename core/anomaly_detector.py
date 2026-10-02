"""
Módulo de Detecção de Anomalias e Auditoria de Consumo
CompaSSS — Automação de Hidrômetros (Praça Pamplona)

Analisa os dados de consumo do mês contra o histórico de 11 meses do StruxureWare
para identificar possíveis vazamentos e medidores parados ou travados.
"""

import logging
import pandas as pd

logger = logging.getLogger(__name__)


def detect_anomalies(df, threshold_mult=3.0, min_diff_m3=5.0, min_zero_hist_m3=5.0):
    """
    Analisa o DataFrame de medições e retorna uma lista de anomalias detectadas.

    Parâmetros:
    - df: DataFrame com colunas ['Usuario', 'Consumo_m3', 'Consumo_Medio_11m']
          ou planilha carregada com cabeçalhos correspondentes.
    - threshold_mult: Multiplicador sobre a média para considerar anomalia (padrão: 3x).
    - min_diff_m3: Diferença mínima em m³ para evitar falsos positivos em volumes pequenos (padrão: 5.0 m³).
    - min_zero_hist_m3: Média histórica mínima para alertar quando a leitura vier zerada (padrão: 5.0 m³).

    Retorna:
    - Lista de dicionários com as salas anômalas ordenadas por gravidade.
    """
    if df is None or len(df) == 0:
        return []

    # Identificar nomes de colunas flexíveis (seja do banco ou do Excel)
    col_user = None
    col_consumo = None
    col_media = None

    for c in df.columns:
        c_str = str(c).strip().lower()
        if "usuario" in c_str or "usuário" in c_str or "sala" in c_str:
            col_user = c
        elif "consumo_m3" in c_str or ("consumo" in c_str and "m³" in c_str and "médio" not in c_str and "medio" not in c_str and "total" not in c_str):
            col_consumo = c
        elif "consumo_medio" in c_str or "consumo_médio" in c_str or ("médio" in c_str and "11" in c_str) or ("medio" in c_str and "11" in c_str):
            col_media = c

    if not col_user or not col_consumo:
        logger.warning(f"Colunas de medição não identificadas no DataFrame: {df.columns.tolist()}")
        return []

    col_user_idx = df.columns.get_loc(col_user)
    col_consumo_idx = df.columns.get_loc(col_consumo)
    col_media_idx = df.columns.get_loc(col_media) if col_media else None

    for row in df.itertuples(index=False):
        sala_raw = row[col_user_idx]
        if pd.isna(sala_raw):
            continue

        sala = str(sala_raw).strip()
        # Ignorar linhas de totalização ou cabeçalhos
        if "total geral" in sala.lower() or ("total" in sala.lower() and len(sala) <= 12):
            continue

        try:
            val_c = row[col_consumo_idx]
            val_consumo = float(val_c or 0.0) if not pd.isna(val_c) else 0.0
        except (ValueError, TypeError):
            continue

        val_media = 0.0
        if col_media_idx is not None:
            try:
                val_m = row[col_media_idx]
                val_media = float(val_m or 0.0) if not pd.isna(val_m) else 0.0
            except (ValueError, TypeError):
                val_media = 0.0

        diff = val_consumo - val_media

        # 1. SUSPEITA DE VAZAMENTO / ALTO CONSUMO REPENTINO
        if val_media > 0 and val_consumo >= (threshold_mult * val_media) and diff >= min_diff_m3:
            pct = ((diff) / val_media) * 100
            anomalies.append({
                "sala": sala,
                "consumo": val_consumo,
                "media": val_media,
                "diff": diff,
                "variacao": f"+{pct:.0f}%",
                "tipo": "Suspeita de Vazamento",
                "detalhe": f"Consumo {val_consumo:.1f} m³ está {val_consumo/val_media:.1f}x acima da média ({val_media:.1f} m³)",
                "severidade": "alta" if val_consumo >= 20.0 else "media"
            })
        elif val_media == 0 and val_consumo >= 10.0:
            anomalies.append({
                "sala": sala,
                "consumo": val_consumo,
                "media": 0.0,
                "diff": val_consumo,
                "variacao": "Novo Consumo",
                "tipo": "Consumo Repentino (Sem Histórico)",
                "detalhe": f"Sala sem consumo anterior registrou {val_consumo:.1f} m³",
                "severidade": "media"
            })

        # 2. MEDIDOR TRAVADO / CONSUMO ZERADO EM SALA HISTORICAMENTE ATIVA
        elif val_media >= min_zero_hist_m3 and val_consumo == 0.0:
            anomalies.append({
                "sala": sala,
                "consumo": 0.0,
                "media": val_media,
                "diff": -val_media,
                "variacao": "-100%",
                "tipo": "Possível Medidor Travado",
                "detalhe": f"Média de {val_media:.1f} m³, mas medição deste mês veio 0,0 m³ (medidor parado ou sala desocupada)",
                "severidade": "media"
            })

    # Ordenar: vazamentos primeiro (pelo maior consumo absoluto), depois medidores zerados
    anomalies.sort(key=lambda x: (x["tipo"] != "Suspeita de Vazamento", -x["consumo"], -x["media"]))

    return anomalies
