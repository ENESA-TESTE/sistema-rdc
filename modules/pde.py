# -*- coding: utf-8 -*-
"""
Módulo de Gestão de PDE (Base de Colaboradores e Funções)
Calcula estatísticas de efetivo, colaboradores sem encarregado e matriz de funções.
"""

import pandas as pd
import streamlit as st
from modules.database import carregar_pde_db, salvar_pde_db

def obter_colaboradores_sem_encarregado(df_pde: pd.DataFrame = None) -> pd.DataFrame:
    """Filtra colaboradores sem encarregado ou com encarregado inválido."""
    if df_pde is None:
        df_pde = carregar_pde_db()
        
    if df_pde.empty or "ENCARREGADO" not in df_pde.columns:
        return pd.DataFrame()
        
    mask = (
        df_pde["ENCARREGADO"].isna() |
        (df_pde["ENCARREGADO"].astype(str).str.strip() == "") |
        (df_pde["ENCARREGADO"].astype(str).str.strip().str.upper() == "NAN") |
        (df_pde["ENCARREGADO"].astype(str).str.strip() == "-") |
        (df_pde["ENCARREGADO"].astype(str).str.strip() == "0")
    )
    return df_pde[mask]

def obter_matriz_funcoes_disciplina(df_pde: pd.DataFrame = None) -> tuple:
    """Gera o resumo e a tabela detalhada de funções por disciplina."""
    if df_pde is None:
        df_pde = carregar_pde_db()
        
    if df_pde.empty or "DISCIPLINA" not in df_pde.columns or "FUNÇÃO" not in df_pde.columns:
        return pd.DataFrame(), pd.DataFrame()
        
    df_pivot = df_pde.groupby(["DISCIPLINA", "FUNÇÃO"]).size().reset_index(name="QTD")
    df_pivot = df_pivot.sort_values(["DISCIPLINA", "QTD"], ascending=[True, False])
    
    df_resumo = df_pde.groupby("DISCIPLINA").agg(
        Total_Pessoas=("NOME", "count"),
        Total_Funcoes=("FUNÇÃO", "nunique")
    ).reset_index().sort_values("Total_Pessoas", ascending=False)
    
    return df_resumo, df_pivot
