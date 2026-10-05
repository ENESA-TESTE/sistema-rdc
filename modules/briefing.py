# -*- coding: utf-8 -*-
"""
Módulo de Briefing Matinal & IA (Resumo Executivo Diário)
Garante carregamento instantâneo, edição interativa e exportação executiva.
"""

import datetime
import pandas as pd
import streamlit as st
from modules.database import (
    carregar_briefings_db, obter_briefing_dia_db, 
    salvar_briefing_dia_db, excluir_briefing_dia_db
)

def renderizar_painel_briefing(df_rdcs, data_selecionada, nome_site="SGO RDC"):
    """Renderiza a interface rápida de Briefing Matinal."""
    st.markdown("### 📋 Briefing Matinal & Síntese da Reunião")
    st.caption("Resumo executivo de avanços, pontos de atenção e bloqueios com persistência na nuvem.")
    
    # 1. Carregar todos os briefings salvos do banco SQLite (< 1ms)
    todos_briefings = carregar_briefings_db()
    
    # Obter dados para a data atual
    briefing_atual = obter_briefing_dia_db(data_selecionada)
    
    if briefing_atual:
        st.success(f"💾 Briefing carregado para **{data_selecionada}** (Origem: {briefing_atual.get('origem', 'Salvo')})")
    else:
        st.info(f"💡 Nenhum briefing salvo ainda para **{data_selecionada}**.")
        
    return briefing_atual
