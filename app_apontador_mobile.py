# -*- coding: utf-8 -*-
"""
====================================================================
ENESA ENGENHARIA — SISTEMA MOBILE DO APONTADOR DE CAMPO (PWA)
Controle e Remanejamento de Efetivo em Tempo Real via Celular
Integração 100% Automática: SQLite + PostgreSQL + Google Sheets
====================================================================
"""

import streamlit as st
import pandas as pd
import datetime
import time
import os
import io
import importlib
import pde_engine
importlib.reload(pde_engine)
import rdc_impressao

# --- CONFIGURAÇÃO MOBILE-FIRST ---
st.set_page_config(
    page_title="ENESA · Apontador de Campo",
    page_icon="📱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- ESTILOS VISUAIS PARA CELULAR (TOUCH-FRIENDLY & DARK MODERN) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700&family=Inter:wght@400;500;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .mobile-header {
        background: linear-gradient(135deg, #091e36, #030f1d);
        border: 1px solid rgba(56, 189, 248, 0.2);
        border-radius: 16px;
        padding: 16px;
        margin-bottom: 16px;
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.4);
    }
    
    .worker-card {
        background: rgba(15, 23, 42, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 14px;
        padding: 14px;
        margin-bottom: 12px;
        backdrop-filter: blur(10px);
    }
    
    .leader-box {
        background: rgba(14, 165, 233, 0.08);
        border: 1px solid rgba(56, 189, 248, 0.3);
        border-radius: 14px;
        padding: 14px;
        margin-bottom: 14px;
    }
    
    .badge-active {
        background: rgba(16, 185, 129, 0.15);
        border: 1px solid rgba(16, 185, 129, 0.35);
        color: #10b981;
        border-radius: 12px;
        padding: 2px 10px;
        font-size: 11px;
        font-weight: 600;
        display: inline-block;
    }
    
    .badge-cc {
        background: rgba(14, 165, 233, 0.15);
        border: 1px solid rgba(14, 165, 233, 0.35);
        color: #38bdf8;
        border-radius: 12px;
        padding: 2px 10px;
        font-size: 11px;
        font-weight: 600;
        display: inline-block;
    }
    
    .badge-none {
        background: rgba(239, 68, 68, 0.15);
        border: 1px solid rgba(239, 68, 68, 0.35);
        color: #f87171;
        border-radius: 12px;
        padding: 2px 10px;
        font-size: 11px;
        font-weight: 600;
        display: inline-block;
    }
    
    .stButton > button {
        border-radius: 12px !important;
        height: 48px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        letter-spacing: 0.3px !important;
    }
    
    .pwa-banner {
        background: linear-gradient(135deg, rgba(14, 165, 233, 0.15), rgba(99, 102, 241, 0.15));
        border: 1px solid rgba(56, 189, 248, 0.4);
        border-radius: 14px;
        padding: 12px 16px;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
</style>

<!-- PWA META TAGS & WEB MANIFEST INJECTION -->
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="ENESA Campo">
<meta name="application-name" content="ENESA Campo">
<meta name="theme-color" content="#091e36">
<meta name="msapplication-navbutton-color" content="#091e36">
<link rel="manifest" href="data:application/manifest+json,%7B%22name%22%3A%22ENESA%20%C2%B7%20Apontador%20de%20Campo%22%2C%22short_name%22%3A%22ENESA%20Campo%22%2C%22start_url%22%3A%22%2F%22%2C%22display%22%3A%22standalone%22%2C%22background_color%22%3A%22%23030f1d%22%2C%22theme_color%22%3A%22%23091e36%22%7D">
""", unsafe_allow_html=True)

# --- CABEÇALHO DO APLICATIVO MOBILE ---
hora_agora = datetime.datetime.now().strftime("%H:%M")
data_agora = datetime.datetime.now().strftime("%d/%m/%Y")

st.markdown(f"""
<div class="mobile-header">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <div style="font-size: 11px; color: #38bdf8; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;">ENESA ENGENHARIA</div>
            <h2 style="margin: 2px 0 0 0; font-size: 1.3rem; color: #f8fafc; font-family: 'Outfit', sans-serif;">📱 Apontador de Campo</h2>
        </div>
        <div style="text-align: right;">
            <div class="badge-active">● ONLINE</div>
            <div style="font-size: 11px; color: #94a3b8; margin-top: 4px;">{data_agora} · {hora_agora}</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- BANNER DE INSTALAÇÃO PWA / APP NO CELULAR ---
with st.expander("📲 Como Instalar este App no Celular (Tela Cheia)", expanded=False):
    st.markdown("""
    <div style="font-size: 13px; color: #cbd5e1; line-height: 1.6;">
        <b>Para usar como aplicativo oficial na tela inicial sem barra de navegador:</b>
        <br><br>
        🤖 <b>No Android (Google Chrome):</b>
        <ol style="margin: 6px 0 10px 18px; padding: 0;">
            <li>Toque nos <b>3 pontinhos (⋮)</b> no canto superior direito do Chrome.</li>
            <li>Selecione <b>"Instalar aplicativo"</b> ou <b>"Adicionar à tela inicial"</b>.</li>
            <li>Pronto! O ícone da <b>ENESA</b> será criado junto com seus outros apps.</li>
        </ol>
        🍎 <b>No iPhone (Safari):</b>
        <ol style="margin: 6px 0 6px 18px; padding: 0;">
            <li>Toque no botão de <b>Compartilhar (quadrado com seta ⎋)</b> no rodapé do Safari.</li>
            <li>Role para baixo e toque em <b>"Adicionar à Tela de Início" (➕)</b>.</li>
            <li>Toque em <b>"Adicionar"</b> no canto superior direito.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)



# --- CARREGAR BASE PDE DO MOTOR ULTRA-RÁPIDO COM SUPORTE À NUVEM ---
@st.cache_data(ttl=180)
def obter_base_pde_otimizada():
    """Lê a base PDE direto da nuvem (Google Sheets) ou fallback local."""
    erro_cloud = None
    
    # 1. Tentar via pde_engine (obter_cliente_google_sheets direto)
    try:
        sh = pde_engine.obter_cliente_google_sheets()
        if sh:
            ws_names = [w.title for w in sh.worksheets()]
            alvo_ws = "PDE" if "PDE" in ws_names else ("Página1" if "Página1" in ws_names else ws_names[0])
            ws = sh.worksheet(alvo_ws)
            valores = ws.get_all_values()
            if valores and len(valores) > 1:
                header = [str(c).strip() for c in valores[0]]
                linhas = valores[1:]
                df_cloud = pd.DataFrame(linhas, columns=header)
                ok, msg = pde_engine.auditar_e_sincronizar_pde(df_cloud, fonte="GOOGLE_SHEETS")
                if ok:
                    df_res = pde_engine.carregar_pde_instantaneo()
                    if df_res is not None and not df_res.empty:
                        return df_res, None
    except Exception as ex1:
        erro_cloud = str(ex1)

    # 2. Tentar via st.connection("gsheets")
    try:
        from streamlit_gsheets import GSheetsConnection
        conn_gs = st.connection("gsheets", type=GSheetsConnection)
        df_nuvem = conn_gs.read(worksheet="PDE", ttl=180)
        if df_nuvem is not None and not df_nuvem.empty and len(df_nuvem.columns) >= 3:
            df_nuvem = df_nuvem.dropna(how='all')
            ok, msg = pde_engine.auditar_e_sincronizar_pde(df_nuvem, fonte="GOOGLE_SHEETS")
            if ok:
                df_res = pde_engine.carregar_pde_instantaneo()
                if df_res is not None and not df_res.empty:
                    return df_res, None
    except Exception as ex2:
        if not erro_cloud:
            erro_cloud = str(ex2)

    # 3. Fallback: carregar instantâneo do SQLite ou CSV local
    df_local = pde_engine.carregar_pde_instantaneo()
    if df_local is not None and not df_local.empty:
        return df_local, None
        
    return None, erro_cloud

df_pde, erro_pde = obter_base_pde_otimizada()

# Blindagem universal de nomes de colunas
if df_pde is not None and not df_pde.empty:
    if "C.C" not in df_pde.columns:
        cands = [c for c in df_pde.columns if any(k in str(c).upper() for k in ["CUSTO", "C.C", "CC"])]
        df_pde["C.C"] = df_pde[cands[0]] if cands else ""
    if "FUNÇÃO" not in df_pde.columns:
        cands = [c for c in df_pde.columns if "FUN" in str(c).upper()]
        df_pde["FUNÇÃO"] = df_pde[cands[0]] if cands else ""
    if "ENCARREGADO" not in df_pde.columns:
        cands = [c for c in df_pde.columns if "ENCARR" in str(c).upper() and "MATR" not in str(c).upper()]
        df_pde["ENCARREGADO"] = df_pde[cands[0]] if cands else ""
    if "STATUS" not in df_pde.columns:
        df_pde["STATUS"] = "ATIVO"
    if "MATRICULA" not in df_pde.columns:
        cands = [c for c in df_pde.columns if "MATR" in str(c).upper() and "ENCARR" not in str(c).upper()]
        df_pde["MATRICULA"] = df_pde[cands[0]] if cands else ""

if df_pde is None or df_pde.empty:
    st.error("⚠️ Banco de dados PDE não carregado.")
    if erro_pde:
        st.warning(f"Detalhe: {erro_pde}")
    st.info("💡 Cole as chaves de acesso no menu do Streamlit Cloud: **Settings ➔ Secrets**.")
    if st.button("🔄 Tentar Novamente", type="primary"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

# Filtrar apenas ativos
# Deduplicação de colunas e filtragem robusta de ativos
df_pde = df_pde.loc[:, ~df_pde.columns.duplicated()]
col_st = df_pde["STATUS"].iloc[:, 0] if isinstance(df_pde["STATUS"], pd.DataFrame) else df_pde["STATUS"]
df_ativos = df_pde[col_st.astype(str).str.strip().str.upper() == "ATIVO"].copy()

# Lista limpa e completa de encarregados ativos
mapa_enc_mat = pde_engine.obter_mapa_encarregados_matricula()
cargos_gestao = ["SUPERVIS", "COORDENAD", "GERENTE", "DIRETOR", "ENGENHEIR"]

# 1. Encarregados que possuem equipe ativa de 3 ou mais colaboradores (sem Abmael)
col_funcao = "FUNÇÃO" if "FUNÇÃO" in df_pde.columns else ("FUNCAO" if "FUNCAO" in df_pde.columns else None)
mapa_funcao_pde = dict(zip(df_pde["NOME"].astype(str).str.strip().str.upper(), df_pde[col_funcao])) if col_funcao else {}

contagem_equipes = df_ativos["ENCARREGADO"].dropna().astype(str).str.strip().str.upper().value_counts()
encs_com_equipe_minima = set(contagem_equipes[contagem_equipes >= 3].index)

enc_da_coluna = set([
    str(e).strip() for e in df_ativos["ENCARREGADO"].dropna().unique() 
    if str(e).strip() and str(e) not in ["-", "0", "nan", "NAN", "SEM ENCARREGADO"]
    and str(e).strip().upper() in encs_com_equipe_minima
    and "ABMAEL" not in str(e).upper()
    and not any(cg in str(e).upper() for cg in cargos_gestao)
    and not any(cg in str(mapa_funcao_pde.get(str(e).strip().upper(), "")).upper() for cg in cargos_gestao)
])

# Lista unificada apenas dos encarregados com equipe ativa na obra (mínimo 3 colaboradores)
encarregados_ativos = sorted(list(enc_da_coluna))

# Opção de deixar sem encarregado
OPCAO_SEM_ENC = "🚫 (Deixar Sem Encarregado)"
opcoes_enc_com_desvincular = [OPCAO_SEM_ENC] + encarregados_ativos

# Centros de Custo formatados com Nome e [PB / RB]
def _carregar_mapa_centros_de_custo():
    mapa = {}
    if os.path.exists("PDE.csv"):
        try:
            df_raw = pd.read_csv("PDE.csv", encoding='latin1', low_memory=False)
            col_cc = [c for c in df_raw.columns if "CUSTO" in str(c).upper() or "C.C" in str(c).upper() or "CC" in str(c).upper()]
            col_cont = [c for c in df_raw.columns if "CONTRATO" in str(c).upper()]
            col_disc = [c for c in df_raw.columns if "DISCIPLINA" in str(c).upper()]
            if col_cc:
                for cc, grp in df_raw.groupby(col_cc[0]):
                    cc_clean = str(cc).strip()
                    if not cc_clean or cc_clean in ["-", "0", "nan", "NAN"]:
                        continue
                    contrato = "RB"
                    if col_cont:
                        contratos = [str(c).strip().upper() for c in grp[col_cont[0]].dropna().unique() if str(c).strip() not in ["-", "nan", ""]]
                        contrato = contratos[0] if contratos else ("PB" if "125.02." in cc_clean else "RB")
                    else:
                        contrato = "PB" if "125.02." in cc_clean else "RB"
                    disc = str(grp[col_disc[0]].dropna().value_counts().index[0]).strip().upper() if col_disc and not grp[col_disc[0]].dropna().empty else ""
                    label = f"{cc_clean} — [{contrato}] {disc}" if disc else f"{cc_clean} — [{contrato}]"
                    mapa[cc_clean] = label
        except Exception:
            pass
    if not mapa and df_ativos is not None and not df_ativos.empty:
        for cc, grp in df_ativos.groupby("C.C"):
            cc_clean = str(cc).strip()
            if not cc_clean or cc_clean in ["-", "0", "nan", "NAN"]:
                continue
            contrato = "PB" if "125.02." in cc_clean else "RB"
            disc = str(grp["DISCIPLINA"].dropna().value_counts().index[0]).strip().upper() if not grp["DISCIPLINA"].dropna().empty else ""
            label = f"{cc_clean} — [{contrato}] {disc}" if disc else f"{cc_clean} — [{contrato}]"
            mapa[cc_clean] = label
    return mapa

mapa_cc_labels = _carregar_mapa_centros_de_custo()

def formatar_cc_display(cc_codigo):
    """Retorna o código do CC com seu contrato [PB/RB] e nome da disciplina."""
    c = str(cc_codigo).strip()
    if not c or c in ["-", "0", "nan", "NAN"]:
        return "Sem C.C"
    return mapa_cc_labels.get(c, f"{c} — [{'PB' if '125.02.' in c else 'RB'}]")

# Lista de opções de C.C formatadas para os selectboxes
ccs_unicos = sorted([
    str(c).strip() for c in df_ativos["C.C"].dropna().unique() 
    if str(c).strip() and str(c) not in ["-", "0", "nan", "NAN"]
])
opcoes_cc_formatadas = [formatar_cc_display(c) for c in ccs_unicos]

# Identificação do Apontador com Memória Automática no Celular (LocalStorage + QueryParams)
param_apontador = st.query_params.get("apontador", "").strip()

if "nome_apontador" not in st.session_state:
    if param_apontador and param_apontador != "Apontador de Campo":
        st.session_state.nome_apontador = param_apontador
    else:
        st.session_state.nome_apontador = ""

if "editando_apontador" not in st.session_state:
    st.session_state.editando_apontador = not bool(st.session_state.nome_apontador)

# Injetar JavaScript para ler e gravar no LocalStorage do celular
st.markdown("""
<script>
    const savedApontador = localStorage.getItem('enesa_apontador_nome');
    const urlParams = new URLSearchParams(window.location.search);
    if (savedApontador && !urlParams.has('apontador')) {
        urlParams.set('apontador', savedApontador);
        window.location.search = urlParams.toString();
    }
</script>
""", unsafe_allow_html=True)

col_ap1, col_ap2 = st.columns([2, 1])
with col_ap1:
    if not st.session_state.editando_apontador and st.session_state.nome_apontador:
        st.markdown(f"""
        <div style="background: rgba(14, 165, 233, 0.1); border: 1px solid rgba(14, 165, 233, 0.3); border-radius: 10px; padding: 8px 12px; margin-top: 4px;">
            <div style="font-size: 11px; color: #38bdf8; font-weight: 600;">👤 APONTADOR IDENTIFICADO:</div>
            <div style="font-size: 14px; font-weight: 700; color: #f8fafc;">{st.session_state.nome_apontador} <span style="font-size: 11px; color: #10b981; font-weight: normal;">(Salvo no Celular ✅)</span></div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("✏️ Trocar Apontador", key="btn_trocar_apontador_id"):
            st.session_state.editando_apontador = True
            st.rerun()
    else:
        nome_input = st.text_input("👤 Seu Nome ou Matrícula:", value=st.session_state.nome_apontador, placeholder="Ex: Carlos Silva ou 12500...", key="txt_apontador_id")
        if st.button("💾 Lembrar Meu Nome neste Celular", type="primary", key="btn_salvar_nome_apontador"):
            if nome_input.strip():
                st.session_state.nome_apontador = nome_input.strip()
                st.session_state.editando_apontador = False
                st.query_params["apontador"] = nome_input.strip()
                st.markdown(f"""
                <script>
                    localStorage.setItem('enesa_apontador_nome', '{nome_input.strip()}');
                </script>
                """, unsafe_allow_html=True)
                st.toast("Nome gravado com sucesso no celular!", icon="💾")
                st.rerun()
            else:
                st.warning("Por favor, digite seu nome.")

with col_ap2:
    st.metric("👥 Efetivo Ativo", f"{len(df_ativos):,}".replace(",", "."))

# Link 4G para compartilhar no WhatsApp da equipe
arq_link_4g = os.path.join(pde_engine.PASTA_BASE, "link_4g_apontadores.txt")
if os.path.exists(arq_link_4g):
    try:
        with open(arq_link_4g, "r", encoding="utf-8") as f:
            url_link_4g = f.read().strip()
        if url_link_4g.startswith("http"):
            with st.expander("📶 Compartilhar Link 4G com outros Apontadores"):
                st.code(url_link_4g, language="text")
                st.caption("Envie este link no WhatsApp para qualquer apontador acessar do celular pelo 4G/5G sem precisar de Wi-Fi.")
    except Exception:
        pass


# --- NAVEGAÇÃO PRINCIPAL (4 ABAS DIRETAS) ---
tab_remanejar, tab_equipe, tab_imprimir, tab_historico = st.tabs([
    "🔄 Remanejar",
    "📋 Ver Equipes",
    "🖨️ Imprimir RDC",
    "📜 Histórico"
])

# ====================================================================
# ABA 1: REMANEJAR (FLUXO DIRETO POR ENCARREGADO OU POR BUSCA)
# ====================================================================
with tab_remanejar:
    
    modo_remanejo = st.radio(
        "Como deseja remanejar?",
        ["👷 Selecionar Encarregado e Colocar Pessoas", "🔍 Buscar Colaborador (1 a 1)"],
        horizontal=True,
        key="radio_modo_remanejo"
    )

    
    # -------------------------------------------------------------
    # FLUXO 1: SELECIONAR ENCARREGADO E ADICIONAR PESSOAS
    # -------------------------------------------------------------
    if modo_remanejo == "👷 Selecionar Encarregado e Colocar Pessoas":
        st.markdown("##### 1. Escolha o Encarregado de Destino:")
        enc_alvo = st.selectbox(
            "Selecione o Encarregado:",
            opcoes_enc_com_desvincular,
            key="sel_enc_alvo_remanejo"
        )
        
        # Obter dados do encarregado alvo
        if enc_alvo == OPCAO_SEM_ENC:
            st.markdown("""
            <div class="leader-box" style="border-color: rgba(239, 68, 68, 0.4); background: rgba(239, 68, 68, 0.08);">
                <div style="font-weight: 700; color: #f87171; font-size: 1.05rem;">🚫 MODO DESVINCULAÇÃO</div>
                <div style="font-size: 0.85rem; color: #cbd5e1; margin-top: 2px;">Os colaboradores selecionados ficarão <b>sem encarregado</b> (Coluna K vazia).</div>
            </div>
            """, unsafe_allow_html=True)
            cc_sugerido_fmt = "(Manter Atual)"
            membros_atuais = df_ativos[df_ativos["ENCARREGADO"].isna() | df_ativos["ENCARREGADO"].isin(["", "-", "nan", "NAN", "SEM ENCARREGADO"])]
        else:
            mat_lider = mapa_enc_mat.get(enc_alvo.upper(), "")
            membros_atuais = df_ativos[df_ativos["ENCARREGADO"] == enc_alvo]
            
            # Buscar CC mais comum do líder
            cc_lider_series = membros_atuais["C.C"].dropna()
            cc_padrao_raw = cc_lider_series.iloc[0] if not cc_lider_series.empty else (ccs_unicos[0] if ccs_unicos else "")
            cc_sugerido_fmt = formatar_cc_display(cc_padrao_raw)
            
            st.markdown(f"""
            <div class="leader-box">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <div style="font-weight: 700; color: #38bdf8; font-size: 1.1rem;">👷 {enc_alvo}</div>
                        <div style="font-size: 0.85rem; color: #94a3b8; margin-top: 2px;">Matrícula: <b style="color:#f8fafc;">{mat_lider}</b></div>
                        <div style="font-size: 0.8rem; color: #38bdf8; margin-top: 4px;">🏢 C.C Padrão: <b>{cc_sugerido_fmt}</b></div>
                    </div>
                    <div style="text-align: right;">
                        <span class="badge-active" style="font-size: 13px; padding: 4px 12px;">👥 {len(membros_atuais)} pessoas</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("##### 2. Selecione os Colaboradores para este Encarregado:")
        
        col_c1, col_c2 = st.columns(2)
        opcoes_cc_com_manter = ["(Manter Atual)"] + opcoes_cc_formatadas
        idx_cc_padrao = opcoes_cc_com_manter.index(cc_sugerido_fmt) if cc_sugerido_fmt in opcoes_cc_com_manter else 0
        
        with col_c1:
            cc_escolhido_fmt = st.selectbox(
                "🏢 Centro de Custo (C.C — PB/RB):",
                opcoes_cc_com_manter,
                index=idx_cc_padrao,
                key="sel_cc_remanejo_multi"
            )
        with col_c2:
            obs_multi = st.text_input("📝 Observação:", placeholder="Ex: Remanejamento...", key="txt_obs_remanejo_multi")

        # Lista de colaboradores formatada com Nome, Matrícula, Função e Encarregado Atual
        opcoes_colabs_multiselect = [
            f"{r['MATRICULA']} — {r['NOME']} ({r['FUNÇÃO']}) [Atual: {r['ENCARREGADO'] if r['ENCARREGADO'] and r['ENCARREGADO'] != '-' else 'Sem Enc.'} | {formatar_cc_display(r['C.C'])}]"
            for _, r in df_ativos.sort_values(by="NOME").iterrows()
        ]
        
        selecionados = st.multiselect(
            "Digite o nome ou matrícula para adicionar colaboradores (quantos quiser):",
            opcoes_colabs_multiselect,
            placeholder="🔎 Busque e selecione 1, 2, 5 ou mais colaboradores...",
            key="ms_colaboradores_remanejo"
        )
        
        if selecionados:
            st.info(f"📌 **{len(selecionados)} colaborador(es)** selecionado(s).")
            
            if enc_alvo == OPCAO_SEM_ENC:
                btn_txt = f"🚫 Desvincular {len(selecionados)} Colaborador(es) (Deixar Sem Encarregado)"
                btn_tipo = "secondary"
            else:
                btn_txt = f"⚡ Colocar {len(selecionados)} Colaborador(es) com {enc_alvo}"
                btn_tipo = "primary"
                
            if st.button(btn_txt, type=btn_tipo, use_container_width=True, key="btn_confirmar_remanejo_multi"):
                enc_final = "" if enc_alvo == OPCAO_SEM_ENC else enc_alvo
                cc_codigo_limpo = cc_escolhido_fmt.split(" — ")[0].strip() if cc_escolhido_fmt != "(Manter Atual)" else None
                
                with st.spinner("Atualizando planilha e salvando no sistema..."):
                    lista_alt = []
                    for item_str in selecionados:
                        mat_str = item_str.split(" — ")[0].strip()
                        nome_colab = ""
                        membros_match = df_ativos[df_ativos["MATRICULA"] == mat_str]
                        if not membros_match.empty:
                            nome_colab = str(membros_match.iloc[0]["NOME"])
                            
                        item_dict = {
                            "MATRICULA": mat_str,
                            "NOME": nome_colab,
                            "NOVO_ENCARREGADO": enc_final
                        }
                        if cc_codigo_limpo:
                            item_dict["NOVO_CC"] = cc_codigo_limpo
                        lista_alt.append(item_dict)
                        
                    ok, msg = pde_engine.aplicar_correcoes_em_massa(lista_alt, sync_sheets=True)
                    if ok:
                        for item_dict in lista_alt:
                            mat_it = item_dict["MATRICULA"]
                            row_it = df_ativos[df_ativos["MATRICULA"] == mat_it].iloc[0]
                            pde_engine.registrar_movimentacao_equipe(
                                matricula=mat_it,
                                nome=row_it["NOME"],
                                de_enc=str(row_it.get("ENCARREGADO", "")),
                                para_enc=enc_final if enc_final else "(Sem Encarregado)",
                                de_cc=str(row_it.get("C.C", "")),
                                para_cc=cc_codigo_limpo if cc_codigo_limpo else str(row_it.get("C.C", "")),
                                apontador=st.session_state.nome_apontador
                            )
                        st.success(f"🎉 **{len(selecionados)} colaborador(es)** alocado(s) com sucesso!\n\n{msg}")
                        st.balloons()
                        st.cache_data.clear()
                        time.sleep(1.5)
                        st.rerun()
                    else:
                        st.error(f"Erro: {msg}")

        # Mostrar membros que já estão na equipe deste encarregado
        if enc_alvo != OPCAO_SEM_ENC and not membros_atuais.empty:
            st.markdown("---")
            with st.expander(f"📋 Ver quem já está na equipe de {enc_alvo} ({len(membros_atuais)} pessoas)"):
                for _, r_m in membros_atuais.sort_values(by="NOME").iterrows():
                    st.markdown(f"""
                    <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.04); border-radius:8px; padding:8px 12px; margin-bottom:6px;">
                        <div>
                            <b style="color:#f8fafc; font-size:13px;">{r_m['NOME']}</b>
                            <div style="font-size:11px; color:#94a3b8;">Mat: {r_m['MATRICULA']} • {r_m['FUNÇÃO']}</div>
                            <div style="font-size:11px; color:#38bdf8;">🏢 {formatar_cc_display(r_m['C.C'])}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button(f"🚫 Tirar da Equipe", key=f"btn_quick_tirar_{r_m['MATRICULA']}"):
                        ok_t, msg_t = pde_engine.aplicar_correcoes_em_massa([{
                            "MATRICULA": r_m["MATRICULA"],
                            "NOME": r_m["NOME"],
                            "NOVO_ENCARREGADO": ""
                        }], sync_sheets=True)
                        if ok_t:
                            pde_engine.registrar_movimentacao_equipe(
                                matricula=r_m["MATRICULA"],
                                nome=r_m["NOME"],
                                de_enc=enc_alvo,
                                para_enc="(Sem Encarregado)",
                                de_cc=r_m["C.C"],
                                para_cc=r_m["C.C"],
                                apontador=st.session_state.nome_apontador
                            )
                            st.warning(f"🚫 **{r_m['NOME']}** desvinculado!\n\n{msg_t}")
                            st.cache_data.clear()
                            time.sleep(1.5)
                            st.rerun()

    # -------------------------------------------------------------
    # FLUXO 2: BUSCA POR COLABORADOR (1 A 1)
    # -------------------------------------------------------------
    else:
        st.markdown("##### 🔍 Localizar e Transferir Colaborador")
        busca_ind = st.text_input("🔎 Digite Nome ou Matrícula:", placeholder="Ex: 12500123 ou FRANCISCO...", key="txt_busca_ind_v2")
        
        if busca_ind.strip():
            bi = busca_ind.strip().upper()
            res_ind = df_ativos[
                df_ativos["NOME"].astype(str).str.contains(bi, na=False) |
                df_ativos["MATRICULA"].astype(str).str.contains(bi, na=False)
            ]
            
            if res_ind.empty:
                st.warning("Nenhum colaborador ativo encontrado com esse termo.")
            else:
                st.caption(f"Encontrados **{len(res_ind)}** colaboradores:")
                
                opcoes_colab = [f"{r['MATRICULA']} — {r['NOME']} ({r['FUNÇÃO']})" for _, r in res_ind.iterrows()]
                escolha = st.selectbox("Selecione o Colaborador:", opcoes_colab, key="sel_colab_ind_v2")
                
                mat_escolhida = escolha.split(" — ")[0].strip()
                row_c = res_ind[res_ind["MATRICULA"] == mat_escolhida].iloc[0]
                
                enc_atual = str(row_c.get("ENCARREGADO", "")).strip()
                tem_enc = enc_atual and enc_atual not in ["-", "0", "nan", "NAN", "SEM ENCARREGADO"]
                enc_display = enc_atual if tem_enc else "Sem Encarregado"
                
                cc_atual_raw = str(row_c.get("C.C", "")).strip()
                cc_atual_display = formatar_cc_display(cc_atual_raw)
                disc_atual = str(row_c.get("DISCIPLINA", "")).strip() or "-"
                
                badge_enc_class = "badge-active" if tem_enc else "badge-none"
                
                st.markdown(f"""
                <div class="worker-card">
                    <div style="font-size: 1.1rem; font-weight: 700; color: #f8fafc;">{row_c['NOME']}</div>
                    <div style="font-size: 0.85rem; color: #94a3b8; margin-top: 2px;">Matrícula: <b style="color:#e2e8f0;">{row_c['MATRICULA']}</b> • {row_c['FUNÇÃO']}</div>
                    <div style="display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap;">
                        <span class="{badge_enc_class}">👷 Líder Atual: {enc_display}</span>
                        <span class="badge-cc">🏢 C.C: {cc_atual_display}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("##### 🚀 Destino da Transferência")
                
                idx_enc_padrao = opcoes_enc_com_desvincular.index(enc_atual) if enc_atual in opcoes_enc_com_desvincular else 0
                novo_enc_escolhido = st.selectbox("👷 Novo Encarregado:", opcoes_enc_com_desvincular, index=idx_enc_padrao, key="sel_novo_enc_ind_v2")
                
                idx_cc_ind = opcoes_cc_formatadas.index(cc_atual_display) if cc_atual_display in opcoes_cc_formatadas else 0
                novo_cc_escolhido_fmt = st.selectbox("🏢 Novo Centro de Custo (C.C — PB/RB):", opcoes_cc_formatadas, index=idx_cc_ind, key="sel_novo_cc_ind_v2")
                
                obs_ind = st.text_input("📝 Motivo / Observação (Opcional):", placeholder="Ex: Reforço, remanejamento...", key="txt_obs_ind_v2")
                
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    btn_confirmar = st.button("⚡ Confirmar Alteração", type="primary", use_container_width=True, key="btn_transf_ind_v2")
                with col_b2:
                    btn_remover_enc = st.button("🚫 Tirar Encarregado", use_container_width=True, key="btn_remover_enc_ind_v2")
                    
                if btn_confirmar or btn_remover_enc:
                    enc_final = "" if (btn_remover_enc or novo_enc_escolhido == OPCAO_SEM_ENC) else novo_enc_escolhido
                    cc_final_raw = novo_cc_escolhido_fmt.split(" — ")[0].strip()
                    
                    with st.spinner("Gravando alteração e sincronizando..."):
                        lista_alt = [{
                            "MATRICULA": row_c["MATRICULA"],
                            "NOME": row_c["NOME"],
                            "NOVO_ENCARREGADO": enc_final,
                            "NOVO_CC": cc_final_raw
                        }]
                        ok, msg = pde_engine.aplicar_correcoes_em_massa(lista_alt, sync_sheets=True)
                        if ok:
                            pde_engine.registrar_movimentacao_equipe(
                                matricula=row_c["MATRICULA"],
                                nome=row_c["NOME"],
                                de_enc=enc_display,
                                para_enc=enc_final if enc_final else "(Sem Encarregado)",
                                de_cc=cc_atual_raw,
                                para_cc=cc_final_raw,
                                apontador=st.session_state.nome_apontador
                            )
                            if enc_final:
                                st.success(f"🎉 **{row_c['NOME']}** vinculado a **{enc_final}**!\n\n{msg}")
                            else:
                                st.warning(f"🚫 **{row_c['NOME']}** agora está **sem encarregado**!\n\n{msg}")
                            st.balloons()
                            st.cache_data.clear()
                            time.sleep(1.5)
                            st.rerun()
                        else:
                            st.error(f"Erro ao salvar: {msg}")
        else:
            st.info("👆 Digite o nome ou a matrícula no campo acima para pesquisar.")

# ====================================================================
# ABA 2: VISÃO POR EQUIPE (Raio-X do Encarregado & Desvinculação Rápida)
# ====================================================================
with tab_equipe:
    st.markdown("##### 📋 Integrantes por Encarregado")
    
    enc_visao = st.selectbox("Selecione o Encarregado:", encarregados_ativos, key="sel_enc_visao")
    
    if enc_visao:
        equipe = df_ativos[df_ativos["ENCARREGADO"] == enc_visao].sort_values(by="NOME")
        
        st.markdown(f"""
        <div style="background: rgba(14, 165, 233, 0.1); border: 1px solid rgba(14, 165, 233, 0.3); border-radius: 12px; padding: 12px; margin-bottom: 12px; display:flex; justify-content:space-between; align-items:center;">
            <div>
                <b style="color: #38bdf8; font-size: 1.1rem;">{enc_visao}</b>
                <div style="font-size: 0.8rem; color: #94a3b8;">Matrícula: {mapa_enc_mat.get(enc_visao, 'N/A')}</div>
            </div>
            <div style="font-size: 1.3rem; font-weight: 700; color: #10b981;">{len(equipe)} <span style="font-size: 0.8rem; color: #94a3b8;">pessoas</span></div>
        </div>
        """, unsafe_allow_html=True)
        
        if equipe.empty:
            st.info("Nenhum colaborador vinculado a este encarregado no momento.")
        else:
            for idx, r_eq in equipe.iterrows():
                with st.expander(f"👤 {r_eq['NOME']} ({r_eq['FUNÇÃO']})"):
                    st.write(f"**Matrícula:** {r_eq['MATRICULA']}")
                    st.write(f"**C.C:** {formatar_cc_display(r_eq['C.C'])}")
                    
                    col_eq_a1, col_eq_a2 = st.columns([2, 1])
                    with col_eq_a1:
                        outro_enc = st.selectbox("Transferir para:", [e for e in encarregados_ativos if e != enc_visao], key=f"sel_transf_eq_{r_eq['MATRICULA']}")
                    with col_eq_a2:
                        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                        btn_tr_direta = st.button("Mudar", key=f"btn_tr_{r_eq['MATRICULA']}", use_container_width=True)
                        
                    btn_desvinc_direta = st.button("🚫 Desvincular da Equipe (Deixar Sem Enc.)", key=f"btn_desv_{r_eq['MATRICULA']}", use_container_width=True)
                    
                    if btn_tr_direta or btn_desvinc_direta:
                        novo_enc_direto = outro_enc if btn_tr_direta else ""
                        
                        ok, msg = pde_engine.aplicar_correcoes_em_massa([{
                            "MATRICULA": r_eq["MATRICULA"],
                            "NOME": r_eq["NOME"],
                            "NOVO_ENCARREGADO": novo_enc_direto
                        }], sync_sheets=True)
                        if ok:
                            pde_engine.registrar_movimentacao_equipe(
                                matricula=r_eq["MATRICULA"],
                                nome=r_eq["NOME"],
                                de_enc=enc_visao,
                                para_enc=novo_enc_direto if novo_enc_direto else "(Sem Encarregado)",
                                de_cc=r_eq["C.C"],
                                para_cc=r_eq["C.C"],
                                apontador=st.session_state.nome_apontador
                            )
                            if novo_enc_direto:
                                st.success(f"Transferido para **{novo_enc_direto}**!\n\n{msg}")
                            else:
                                st.warning(f"Desvinculado de **{enc_visao}**!\n\n{msg}")
                            st.cache_data.clear()
                            time.sleep(1.5)
                            st.rerun()

# ====================================================================
# ABA 3: IMPRIMIR RDC (PADRÃO OFICIAL ENESA)
# ====================================================================
with tab_imprimir:
    st.markdown("##### 🖨️ Emissão e Impressão de RDC")
    st.caption("Gere os formulários oficiais em PDF ou Excel prontos para imprimir e levar a campo.")
    
    # 1. Modo de Seleção de Encarregados
    modo_selecao_enc = st.radio("Como deseja selecionar?", ["1 Encarregado", "Múltiplos / Todos"], horizontal=True, key="modo_sel_enc_imp")

    if modo_selecao_enc == "1 Encarregado":
        enc_unico = st.selectbox("Escolha o Encarregado:", encarregados_ativos, key="sel_enc_imp_unico")
        lista_alvo_imp = [enc_unico] if enc_unico else []
    else:
        sel_todos = st.checkbox("Selecionar Todos os Encarregados Ativos", key="chk_sel_todos_imp")
        if sel_todos:
            lista_alvo_imp = encarregados_ativos
            st.info(f"✅ Todos os **{len(lista_alvo_imp)}** encarregados ativos selecionados!")
        else:
            lista_alvo_imp = st.multiselect(
                "Escolha os Encarregados:",
                encarregados_ativos,
                default=[encarregados_ativos[0]] if encarregados_ativos else [],
                key="multi_enc_imp"
            )

    if not lista_alvo_imp:
        st.warning("⚠️ Selecione pelo menos um encarregado.")
    else:
        equipe_imp = df_ativos[df_ativos["ENCARREGADO"].isin(lista_alvo_imp)].copy()
        
        # Card Resumo do RDC
        texto_enc_resumo = ", ".join(lista_alvo_imp[:2]) + (f" e mais {len(lista_alvo_imp)-2}" if len(lista_alvo_imp) > 2 else "")
        st.markdown(f"""
        <div class="leader-box" style="margin-top: 10px; border-color: rgba(14, 165, 233, 0.4);">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div style="font-weight: 700; color: #38bdf8; font-size: 1rem;">📄 RDC · RELATÓRIO DIÁRIO DE CAMPO</div>
                <div class="badge-active">A4 OFICIAL</div>
            </div>
            <div style="margin-top: 8px; font-size: 13px; color: #cbd5e1; line-height: 1.5;">
                <div><b>Encarregado(s):</b> {texto_enc_resumo} ({len(lista_alvo_imp)})</div>
                <div><b>Data:</b> ✍️ <i>Preenchimento manual a caneta no canteiro</i></div>
                <div><b>Efetivo Total:</b> {len(equipe_imp)} colaboradores vinculados</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Preview da equipe
        with st.expander(f"👥 Visualizar Efetivo ({len(equipe_imp)} colaboradores)", expanded=False):
            if not equipe_imp.empty:
                colunas_prev = ["MATRICULA", "NOME", "FUNÇÃO", "ENCARREGADO"]
                colunas_existentes = [c for c in colunas_prev if c in equipe_imp.columns]
                st.dataframe(
                    equipe_imp[colunas_existentes].reset_index(drop=True),
                    hide_index=True,
                    use_container_width=True
                )
            else:
                st.info("Nenhum colaborador encontrado para o(s) encarregado(s) selecionado(s).")
                
        st.markdown("###### 📥 Opções de Impressão e Download:")
        
        rotulo_btn_pdf = f"🖨️ GERAR PDF — {lista_alvo_imp[0]}" if len(lista_alvo_imp) == 1 else "🖨️ GERAR PDF ÚNICO (Impressão A4)"
        if st.button(rotulo_btn_pdf, type="primary", use_container_width=True, key="btn_pdf_unico_imp"):
            with st.spinner("Gerando PDF oficial de alta qualidade..."):
                pdf_unico_bytes, qtd_gerada = rdc_impressao.gerar_pdf_unico_lote(
                    df_base=df_ativos,
                    lista_encarregados=lista_alvo_imp,
                    nome_empresa="ENESA Engenharia",
                    data_rdc=""
                )
                if pdf_unico_bytes and qtd_gerada > 0:
                    if len(lista_alvo_imp) == 1:
                        nome_limpo_enc = lista_alvo_imp[0].replace(" ", "_").strip().upper()
                        nome_pdf_arq = f"RDC_{nome_limpo_enc}.pdf"
                        rotulo_dl = f"⬇️ Baixar RDC de {lista_alvo_imp[0]} (.pdf)"
                    elif len(lista_alvo_imp) <= 2:
                        nomes_curtos = "_".join([e.replace(" ", "_").strip().upper() for e in lista_alvo_imp])
                        nome_pdf_arq = f"RDC_{nomes_curtos}.pdf"
                        rotulo_dl = f"⬇️ Baixar PDF Único ({len(lista_alvo_imp)} Encarregados)"
                    else:
                        nome_pdf_arq = f"RDC_LOTE_{len(lista_alvo_imp)}_ENCARREGADOS.pdf"
                        rotulo_dl = f"⬇️ Baixar PDF Único ({len(lista_alvo_imp)} Encarregados)"
                        
                    st.download_button(
                        label=rotulo_dl,
                        data=pdf_unico_bytes,
                        file_name=nome_pdf_arq,
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True,
                        key="dl_pdf_unico_mobile"
                    )
                    st.success(f"✅ Arquivo `{nome_pdf_arq}` pronto! Toque acima para salvar.")
                else:
                    st.error("Não foi possível gerar o PDF. Verifique se os encarregados possuem equipe.")

        col_b_exc, col_b_zip = st.columns(2)
        with col_b_exc:
            if st.button("🟢 Gerar Excel (.xlsx)", use_container_width=True, key="btn_excel_imp"):
                if len(lista_alvo_imp) == 1:
                    enc_alvo_exc = lista_alvo_imp[0]
                    nome_limpo_enc = enc_alvo_exc.replace(" ", "_").strip().upper()
                    eq_exc = df_ativos[df_ativos["ENCARREGADO"] == enc_alvo_exc]
                    wb_res = rdc_impressao.preencher_excel(eq_exc, enc_alvo_exc, data_rdc="")
                    if wb_res:
                        buf_exc = io.BytesIO()
                        wb_res.save(buf_exc)
                        wb_res.close()
                        buf_exc.seek(0)
                        nome_exc_arq = f"RDC_{nome_limpo_enc}.xlsx"
                        st.download_button(
                            f"⬇️ Baixar {nome_exc_arq}",
                            data=buf_exc,
                            file_name=nome_exc_arq,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="dl_exc_mobile"
                        )
                        st.success(f"✅ Planilha `{nome_exc_arq}` gerada!")
                    else:
                        st.error("Modelo MODELO.xlsx não encontrado no servidor.")
                else:
                    with st.spinner("Gerando planilhas em Lote..."):
                        zip_exc_bytes, qtd_exc = rdc_impressao.gerar_zip_lote(
                            df_base=df_ativos,
                            lista_encarregados=lista_alvo_imp,
                            formato="excel",
                            data_rdc=""
                        )
                        if zip_exc_bytes:
                            st.download_button(
                                f"⬇️ Baixar {qtd_exc} Excels (.zip)",
                                data=zip_exc_bytes,
                                file_name=f"LOTE_EXCEL_RDC_{qtd_exc}Encarregados.zip",
                                mime="application/zip",
                                use_container_width=True,
                                key="dl_zip_exc_mobile"
                            )
                            st.success(f"✅ {qtd_exc} Planilhas Excel compactadas no ZIP!")
                        else:
                            st.warning("Nenhuma planilha gerada.")

        with col_b_zip:
            if st.button("🚀 Pacote ZIP (PDFs)", use_container_width=True, key="btn_zip_pdf_imp"):
                with st.spinner("Compactando PDFs individuais em ZIP..."):
                    zip_pdf_bytes, qtd_zip = rdc_impressao.gerar_zip_lote(
                        df_base=df_ativos,
                        lista_encarregados=lista_alvo_imp,
                        formato="pdf",
                        nome_empresa="ENESA Engenharia",
                        data_rdc=""
                    )
                    if zip_pdf_bytes:
                        if len(lista_alvo_imp) == 1:
                            nome_zip_arq = f"RDC_{lista_alvo_imp[0].replace(' ', '_').strip().upper()}.zip"
                        else:
                            nome_zip_arq = f"LOTE_PDF_RDC_{qtd_zip}Encarregados.zip"
                        st.download_button(
                            f"⬇️ Baixar {qtd_zip} PDFs (.zip)",
                            data=zip_pdf_bytes,
                            file_name=nome_zip_arq,
                            mime="application/zip",
                            use_container_width=True,
                            key="dl_zip_pdf_mobile"
                        )
                        st.success(f"✅ {qtd_zip} PDFs individuais gerados no ZIP!")
                    else:
                        st.warning("Nenhuma PDF gerado.")

# ====================================================================
# ABA 4: HISTÓRICO DE MOVIMENTAÇÕES (Audit Trail)
# ====================================================================
with tab_historico:
    st.markdown("##### 📜 Histórico de Alterações Realizadas")
    
    df_hist = pde_engine.obter_historico_movimentacoes(limite=50)
    
    if df_hist.empty:
        st.info("Nenhuma movimentação registrada hoje ainda.")
    else:
        st.caption(f"Últimas **{len(df_hist)}** transferências de campo:")
        for _, r_h in df_hist.iterrows():
            st.markdown(f"""
            <div style="background: rgba(15, 23, 42, 0.6); border-left: 4px solid #0ea5e9; border-radius: 8px; padding: 10px 14px; margin-bottom: 10px; font-size: 13px;">
                <div style="display: flex; justify-content: space-between; color: #94a3b8; font-size: 11px;">
                    <span>🕒 {r_h['DATA_HORA']}</span>
                    <span>👤 {r_h['APONTADOR']}</span>
                </div>
                <div style="font-weight: 600; color: #f8fafc; margin: 4px 0;">{r_h['NOME']} ({r_h['MATRICULA']})</div>
                <div style="color: #cbd5e1;">
                    De: <span style="color:#ef4444;">{r_h['DE_ENCARREGADO']}</span> ➔ Para: <span style="color:#10b981; font-weight:600;">{r_h['PARA_ENCARREGADO']}</span>
                </div>
                {f"<div style='font-size: 11px; color: #38bdf8; margin-top: 2px;'>🏢 C.C: {formatar_cc_display(r_h['PARA_CC'])}</div>" if r_h.get('PARA_CC') else ''}
            </div>
            """, unsafe_allow_html=True)
            
        if st.button("🔄 Atualizar Histórico", use_container_width=True):
            st.rerun()
