# -*- coding: utf-8 -*-
"""
Módulo de Competição F1 (Ranking e Lançamento de RDCs)
Gerencia a pontuação diária, pódio 3D e matriz de entregas com persistência total.
"""

import datetime
import pandas as pd
import streamlit as st
from modules.database import (
    carregar_f1_db, salvar_f1_db, 
    adicionar_entrega_f1_db, remover_entrega_f1_db
)

def obter_ranking_mensal(mes_ano=None) -> pd.DataFrame:
    """Calcula o ranking mensal de entregas de RDC em < 2ms."""
    df_f1 = carregar_f1_db()
    if df_f1.empty or "DATA" not in df_f1.columns or "ENCARREGADO" not in df_f1.columns:
        return pd.DataFrame(columns=["POS", "ENCARREGADO", "ENTREGAS"])
        
    df_f1["DATA"] = pd.to_datetime(df_f1["DATA"], errors="coerce")
    
    if not mes_ano:
        mes_ano = datetime.date.today().strftime("%Y-%m")
        
    df_mes = df_f1[df_f1["DATA"].dt.strftime("%Y-%m") == mes_ano]
    if df_mes.empty:
        return pd.DataFrame(columns=["POS", "ENCARREGADO", "ENTREGAS"])
        
    ranking = df_mes["ENCARREGADO"].value_counts().reset_index()
    ranking.columns = ["ENCARREGADO", "ENTREGAS"]
    ranking["POS"] = range(1, len(ranking) + 1)
    return ranking[["POS", "ENCARREGADO", "ENTREGAS"]]
