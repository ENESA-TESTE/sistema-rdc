# -*- coding: utf-8 -*-
import base64
import json
import urllib.parse
import streamlit as st
import pandas as pd
import datetime
import time
import os
import re
import io
import importlib
import pde_engine
importlib.reload(pde_engine)
import rdc_impressao
importlib.reload(rdc_impressao)

# --- CONFIGURAÇÃO MOBILE-FIRST COM ÍCONE SGO OFICIAL ---
_sgo_fav_mob = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sgo_favicon.png")
B64_FAVICON_SGO = "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"
if not os.path.exists(_sgo_fav_mob) or os.path.getsize(_sgo_fav_mob) < 100:
    try:
        with open(_sgo_fav_mob, "wb") as _ff:
            _ff.write(base64.b64decode(B64_FAVICON_SGO))
    except Exception:
        pass

st.set_page_config(
    page_title="SGO · Apontador de Campo",
    page_icon=(_sgo_fav_mob if os.path.exists(_sgo_fav_mob) else "📱"),
    layout="centered",
    initial_sidebar_state="collapsed"
)

_manifest_dict = {
    "name": "SGO RDC & PDE · Apontador de Campo",
    "short_name": "SGO Campo",
    "start_url": "/",
    "display": "standalone",
    "background_color": "#0F172A",
    "theme_color": "#0F172A",
    "icons": [
        {
            "src": f"data:image/png;base64,{B64_FAVICON_SGO}",
            "sizes": "256x256",
            "type": "image/png",
            "purpose": "any maskable"
        }
    ]
}
_manifest_uri = "data:application/manifest+json," + urllib.parse.quote(json.dumps(_manifest_dict))

# ====================================================================
# 🎨 DESIGN SYSTEM PREMIUM MOBILE (4 MELHORIAS DE LAYOUT INTEGRADAS)
# ====================================================================
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"], .stApp {{
        font-family: 'Inter', sans-serif;
        background: radial-gradient(circle at 50% 0%, #1E293B 0%, #0F172A 75%) !important;
        color: #F8FAFC !important;
    }}
    
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header {{visibility: hidden;}}
    
    /* 1. Aproveitamento máximo da tela do celular (sem margens mortas) */
    .block-container {{
        padding-top: 0.75rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 0.7rem !important;
        padding-right: 0.7rem !important;
        max-width: 660px !important;
    }}

    /* 2. Header Unificado All-in-One */
    .sgo-unified-header {{
        background: linear-gradient(145deg, rgba(30, 41, 59, 0.96), rgba(15, 23, 42, 0.98));
        border: 1px solid rgba(56, 189, 248, 0.32);
        border-radius: 16px;
        padding: 11px 14px;
        margin-bottom: 8px;
        box-shadow: 0 10px 25px rgba(2, 6, 23, 0.55);
    }}
    .sgo-header-top {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding-bottom: 8px;
        border-bottom: 1px solid rgba(148, 163, 184, 0.14);
    }}
    .sgo-header-user {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding-top: 8px;
        gap: 8px;
    }}

    /* 3. Barra de Abas Estilo "App Nativo" (Botões Pílula / Segmented Control) */
    div[data-baseweb="tab-list"] {{
        display: grid !important;
        grid-template-columns: repeat(5, 1fr) !important;
        gap: 5px !important;
        background: rgba(15, 23, 42, 0.85) !important;
        padding: 5px !important;
        border-radius: 14px !important;
        border: 1px solid rgba(56, 189, 248, 0.24) !important;
        margin-bottom: 10px !important;
        box-shadow: inset 0 2px 6px rgba(0, 0, 0, 0.35) !important;
    }}
    button[data-baseweb="tab"] {{
        border-radius: 10px !important;
        padding: 7px 2px !important;
        height: auto !important;
        min-height: 44px !important;
        margin: 0 !important;
        background: transparent !important;
        border: none !important;
        color: #94A3B8 !important;
        font-size: 11px !important;
        font-weight: 600 !important;
        line-height: 1.15 !important;
        white-space: normal !important;
        text-align: center !important;
        transition: all 0.18s ease !important;
    }}
    button[data-baseweb="tab"] p {{
        font-size: 11px !important;
        font-weight: 700 !important;
        line-height: 1.15 !important;
        margin: 0 !important;
    }}
    button[data-baseweb="tab"][aria-selected="true"] {{
        background: linear-gradient(135deg, #0284C7 0%, #0EA5E9 100%) !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 12px rgba(14, 165, 233, 0.42) !important;
    }}
    div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] {{
        display: none !important;
    }}

    /* 4. Cards e Badges Coloridas por Turno, Matrícula e Contrato */
    .worker-card {{
        background: linear-gradient(145deg, rgba(30, 41, 59, 0.88), rgba(15, 23, 42, 0.92));
        border: 1px solid rgba(148, 163, 184, 0.2);
        border-radius: 14px;
        padding: 13px 14px;
        margin-bottom: 10px;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.25);
    }}
    
    .leader-box {{
        background: linear-gradient(135deg, rgba(14, 165, 233, 0.13), rgba(30, 41, 59, 0.85));
        border: 1px solid rgba(56, 189, 248, 0.38);
        border-radius: 14px;
        padding: 13px 14px;
        margin-bottom: 12px;
    }}

    .badge-mat {{
        background: rgba(15, 23, 42, 0.9);
        border: 1px solid rgba(148, 163, 184, 0.35);
        color: #E2E8F0;
        border-radius: 8px;
        padding: 2px 8px;
        font-size: 11px;
        font-weight: 700;
        font-family: monospace;
        display: inline-block;
    }}
    
    .badge-active {{
        background: rgba(16, 185, 129, 0.16);
        border: 1px solid rgba(16, 185, 129, 0.42);
        color: #34D399;
        border-radius: 999px;
        padding: 3px 10px;
        font-size: 11px;
        font-weight: 700;
        display: inline-block;
    }}
    
    .badge-cc {{
        background: rgba(56, 189, 248, 0.14);
        border: 1px solid rgba(56, 189, 248, 0.36);
        color: #38BDF8;
        border-radius: 8px;
        padding: 2px 9px;
        font-size: 11px;
        font-weight: 600;
        display: inline-block;
    }}

    .badge-t1 {{
        background: rgba(56, 189, 248, 0.16);
        border: 1px solid rgba(56, 189, 248, 0.45);
        color: #38BDF8;
        border-radius: 8px;
        padding: 2px 9px;
        font-size: 11px;
        font-weight: 700;
        display: inline-block;
    }}
    .badge-t2 {{
        background: rgba(245, 158, 11, 0.16);
        border: 1px solid rgba(245, 158, 11, 0.45);
        color: #FBBF24;
        border-radius: 8px;
        padding: 2px 9px;
        font-size: 11px;
        font-weight: 700;
        display: inline-block;
    }}
    .badge-t3 {{
        background: rgba(168, 85, 247, 0.18);
        border: 1px solid rgba(168, 85, 247, 0.45);
        color: #C084FC;
        border-radius: 8px;
        padding: 2px 9px;
        font-size: 11px;
        font-weight: 700;
        display: inline-block;
    }}
    
    .badge-none {{
        background: rgba(239, 68, 68, 0.15);
        border: 1px solid rgba(239, 68, 68, 0.38);
        color: #F87171;
        border-radius: 8px;
        padding: 2px 9px;
        font-size: 11px;
        font-weight: 600;
        display: inline-block;
    }}

    /* Inputs e Botões Confortáveis para Toque */
    .stTextInput > div > div > input {{
        border-radius: 12px !important;
        min-height: 45px !important;
        background: rgba(15, 23, 42, 0.78) !important;
        border: 1px solid rgba(148, 163, 184, 0.28) !important;
        color: #F8FAFC !important;
        font-size: 14px !important;
    }}
    .stTextInput > div > div > input:focus {{
        border-color: #38BDF8 !important;
        box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.22) !important;
    }}
    
    .stButton > button {{
        border-radius: 12px !important;
        min-height: 44px !important;
        font-weight: 700 !important;
        font-size: 13.5px !important;
        letter-spacing: 0.2px !important;
    }}
    .stButton > button[kind="primary"] {{
        background: linear-gradient(135deg, #0284C7 0%, #0EA5E9 100%) !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(14, 165, 233, 0.35) !important;
    }}
</style>

<!-- PWA META TAGS & WEB MANIFEST COM ÍCONE SGO OFICIAL -->
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="SGO Campo">
<meta name="application-name" content="SGO Campo">
<meta name="theme-color" content="#0F172A">
<link rel="apple-touch-icon" href="data:image/png;base64,{B64_FAVICON_SGO}">
<link rel="icon" type="image/png" href="data:image/png;base64,{B64_FAVICON_SGO}">
<link rel="manifest" href="{_manifest_uri}">
""", unsafe_allow_html=True)

def badge_turno_html(turno_raw):
    t = str(turno_raw or "").strip().upper()
    if "2" in t:
        return '<span class="badge-t2">🌆 2° TURNO</span>'
    if "3" in t:
        return '<span class="badge-t3">🌙 3° TURNO</span>'
    return f'<span class="badge-t1">☀️ {t if t else "1° TURNO"}</span>'

# --- MOTOR DE SINCRONIZAÇÃO EM TEMPO REAL (5 SEGUNDOS) PARA O CELULAR ---
import threading as _th_mob
import hashlib as _hl_mob

@st.cache_resource(show_spinner=False)
def obter_monitor_mobile_tempo_real():
    estado = {
        "versao": 0,
        "df_pde": None,
        "hash_pde": "",
        "ultima_sync_str": "Conectando...",
        "forcar_agora": False,
        "rodando": False,
        "_sh_cached": None,
        "_sh_ts": 0.0,
    }

    def _ciclo_mob():
        try:
            agora = time.time()
            if estado["_sh_cached"] is None or (agora - estado["_sh_ts"]) > 900:
                estado["_sh_cached"] = pde_engine.obter_cliente_google_sheets()
                estado["_sh_ts"] = agora
            sh = estado["_sh_cached"]
            if not sh:
                return
            res_b = sh.values_batch_get(["PDE!A1:Z3500"])
            vrs = res_b.get("valueRanges", []) if isinstance(res_b, dict) else []
            if vrs:
                vals = vrs[0].get("values", [])
                if vals and len(vals) > 1:
                    h_novo = _hl_mob.md5(repr(vals).encode("utf-8", errors="ignore")).hexdigest()
                    if h_novo != estado["hash_pde"]:
                        hdr = [str(c).strip() for c in vals[0]]
                        nc = len(hdr)
                        rows = [r + [""] * max(0, nc - len(r)) if len(r) < nc else r[:nc] for r in vals[1:]]
                        df_cloud = pd.DataFrame(rows, columns=hdr).dropna(how="all")
                        if not df_cloud.empty:
                            pde_engine.auditar_e_sincronizar_pde(df_cloud, fonte="GOOGLE_SHEETS")
                            df_res = pde_engine.carregar_pde_instantaneo()
                            if df_res is not None and not df_res.empty:
                                estado["df_pde"] = df_res
                                estado["hash_pde"] = h_novo
                                estado["versao"] += 1
            estado["ultima_sync_str"] = datetime.datetime.now().strftime("%H:%M:%S")
        except Exception:
            estado["_sh_cached"] = None

    def _loop_mob():
        while True:
            _ciclo_mob()
            for _ in range(5):
                if estado.get("forcar_agora"):
                    estado["forcar_agora"] = False
                    break
                time.sleep(1)

    if not estado["rodando"]:
        estado["rodando"] = True
        _th_mob.Thread(target=_loop_mob, daemon=True).start()

    return estado

_MON_MOB = obter_monitor_mobile_tempo_real()

def _invalidar_cache_apos_gravacao():
    """Força recarga imediata do PDE local/nuvem após salvar alterações."""
    try:
        _MON_MOB["df_pde"] = None
        _MON_MOB["hash_pde"] = ""
        _MON_MOB["forcar_agora"] = True
    except Exception:
        pass
    st.cache_data.clear()

# --- CARREGAR BASE PDE EM TEMPO REAL ---
def obter_base_pde_otimizada():
    mon = obter_monitor_mobile_tempo_real()
    if mon.get("df_pde") is not None and not mon["df_pde"].empty:
        st.session_state["_hash_mob_pde"] = mon.get("hash_pde", "")
        return mon["df_pde"].copy(), None
    df_local = pde_engine.carregar_pde_instantaneo()
    if df_local is not None and not df_local.empty:
        return df_local, None
    return None, "Sincronizando com a nuvem..."

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
    if "TURNO" not in df_pde.columns:
        cands_t = [c for c in df_pde.columns if "TURNO" in str(c).upper()]
        df_pde["TURNO"] = df_pde[cands_t[0]] if cands_t else "1° TURNO"
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
df_pde = df_pde.loc[:, ~df_pde.columns.duplicated()]
col_st = df_pde["STATUS"].iloc[:, 0] if isinstance(df_pde["STATUS"], pd.DataFrame) else df_pde["STATUS"]
df_ativos = df_pde[col_st.astype(str).str.strip().str.upper() == "ATIVO"].copy()

# ====================================================================
# 🔐 PORTÃO DE SEGURANÇA OBRIGATÓRIO: LOGIN CORPORATIVO PREMIUM
# ====================================================================
SENHA_PADRAO_APONTADOR = "enesa2026"

def _gerar_token_auth_mat(mat_str):
    m_clean = str(mat_str).strip().replace(".0", "")
    return _hl_mob.md5(f"{m_clean}:{SENHA_PADRAO_APONTADOR}:sgo_v2".encode("utf-8")).hexdigest()[:16]

mapa_matricula_colaborador = {}
for _df_src in [df_pde, df_ativos]:
    for _, _row_u in _df_src.iterrows():
        _m_raw = str(_row_u.get("MATRICULA", "")).strip().replace(".0", "")
        if not _m_raw or _m_raw in ["-", "0", "nan", "NAN"]:
            continue
        _info_u = {
            "MATRICULA": _m_raw,
            "NOME": str(_row_u.get("NOME", "")).strip().upper(),
            "FUNCAO": str(_row_u.get("FUNÇÃO", _row_u.get("FUNCAO", ""))).strip().upper() or "COLABORADOR",
            "STATUS": str(_row_u.get("STATUS", "ATIVO")).strip().upper()
        }
        mapa_matricula_colaborador[_m_raw] = _info_u
        if _m_raw.isdigit():
            mapa_matricula_colaborador[str(int(_m_raw))] = _info_u

def _buscar_colaborador_por_matricula(mat_input):
    m_limpa = re.sub(r"[^0-9A-Za-z]", "", str(mat_input).strip().replace(".0", ""))
    if not m_limpa:
        return None
    achado = mapa_matricula_colaborador.get(m_limpa)
    if not achado and m_limpa.isdigit():
        achado = mapa_matricula_colaborador.get(str(int(m_limpa)))
    return achado

if "apontador_autenticado" not in st.session_state:
    st.session_state.apontador_autenticado = False
    st.session_state.mat_apontador = ""
    st.session_state.nome_apontador_curto = ""
    st.session_state.funcao_apontador = ""
    st.session_state.nome_apontador = ""

_q_mat = st.query_params.get("auth_mat", "").strip()
_q_tk = st.query_params.get("auth_tk", "").strip()
if not st.session_state.apontador_autenticado and _q_mat and _q_tk:
    _colab_q = _buscar_colaborador_por_matricula(_q_mat)
    if _colab_q and _q_tk == _gerar_token_auth_mat(_colab_q["MATRICULA"]):
        st.session_state.apontador_autenticado = True
        st.session_state.mat_apontador = _colab_q["MATRICULA"]
        st.session_state.nome_apontador_curto = _colab_q["NOME"]
        st.session_state.funcao_apontador = _colab_q["FUNCAO"]
        st.session_state.nome_apontador = f"{_colab_q['NOME']} (Mat: {_colab_q['MATRICULA']})"

# --------------------------------------------------------------------
# TELA DE LOGIN ESTILO "APP BANCÁRIO / CORPORATIVO" (QUANDO DESLOGADO)
# --------------------------------------------------------------------
if not st.session_state.apontador_autenticado:
    st.markdown("""
    <script>
        const savedMat = localStorage.getItem('enesa_auth_mat');
        const savedTk = localStorage.getItem('enesa_auth_tk');
        const urlParams = new URLSearchParams(window.location.search);
        if (savedMat && savedTk && !urlParams.has('auth_mat')) {
            urlParams.set('auth_mat', savedMat);
            urlParams.set('auth_tk', savedTk);
            window.location.search = urlParams.toString();
        }
    </script>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div style="background: linear-gradient(160deg, rgba(30, 41, 59, 0.95) 0%, rgba(15, 23, 42, 0.98) 100%); border: 1px solid rgba(56, 189, 248, 0.38); border-radius: 22px; padding: 26px 20px 18px 20px; margin: 14px 0 14px 0; box-shadow: 0 20px 45px rgba(2, 6, 23, 0.7); text-align: center;">
        <div style="width: 68px; height: 68px; margin: 0 auto 12px auto; border-radius: 18px; padding: 3px; background: linear-gradient(135deg, #38BDF8, #10B981); box-shadow: 0 8px 24px rgba(56, 189, 248, 0.35);">
            <img src="data:image/png;base64,{B64_FAVICON_SGO}" style="width: 100%; height: 100%; border-radius: 15px; display: block;">
        </div>
        <div style="display: inline-block; background: rgba(56, 189, 248, 0.14); border: 1px solid rgba(56, 189, 248, 0.35); color: #38BDF8; padding: 3px 12px; border-radius: 999px; font-size: 10.5px; font-weight: 700; letter-spacing: 0.8px; text-transform: uppercase;">
            🔐 PORTAL OFICIAL ENESA · SGO RDC &amp; PDE
        </div>
        <h2 style="margin: 10px 0 4px 0; color: #F8FAFC; font-family: 'Outfit', sans-serif; font-size: 1.45rem; font-weight: 800;">Apontador de Campo</h2>
        <p style="margin: 0; font-size: 12.5px; color: #94A3B8; line-height: 1.45;">
            Entre com sua <b>Matrícula</b> e a <b>Senha Padrão</b>.<br>O sistema reconhece seu nome automaticamente no histórico.
        </p>
    </div>
    """, unsafe_allow_html=True)

    with st.form("form_login_apontador_obrigatorio", clear_on_submit=False):
        mat_login_input = st.text_input(
            "🪪 Sua Matrícula ENESA:",
            placeholder="Digite apenas números (Ex: 12501473)",
            key="login_mat_input"
        )
        senha_login_input = st.text_input(
            "🔑 Senha de Acesso:",
            type="password",
            placeholder="Digite a senha padrão da obra...",
            key="login_senha_input"
        )
        st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
        btn_entrar_login = st.form_submit_button(
            "🔓 Acessar Painel de Apontamento",
            type="primary",
            use_container_width=True
        )

    if btn_entrar_login:
        mat_limpa_tentativa = str(mat_login_input).strip()
        senha_limpa_tentativa = str(senha_login_input).strip().lower()

        if not mat_limpa_tentativa:
            st.error("⚠️ Digite o número da sua **Matrícula**.")
        elif senha_limpa_tentativa != SENHA_PADRAO_APONTADOR:
            st.error("❌ **Senha incorreta!** Solicite a senha padrão com a coordenação.")
        else:
            colab_encontrado = _buscar_colaborador_por_matricula(mat_limpa_tentativa)
            if not colab_encontrado:
                st.error(f"❌ Matrícula **`{mat_limpa_tentativa}`** não encontrada na base do PDE! Verifique os números digitados.")
            else:
                mat_oficial = colab_encontrado["MATRICULA"]
                nome_oficial = colab_encontrado["NOME"]
                func_oficial = colab_encontrado["FUNCAO"]
                tk_gerado = _gerar_token_auth_mat(mat_oficial)

                st.session_state.apontador_autenticado = True
                st.session_state.mat_apontador = mat_oficial
                st.session_state.nome_apontador_curto = nome_oficial
                st.session_state.funcao_apontador = func_oficial
                st.session_state.nome_apontador = f"{nome_oficial} (Mat: {mat_oficial})"

                st.query_params["auth_mat"] = mat_oficial
                st.query_params["auth_tk"] = tk_gerado

                st.markdown(f"""
                <script>
                    localStorage.setItem('enesa_auth_mat', '{mat_oficial}');
                    localStorage.setItem('enesa_auth_tk', '{tk_gerado}');
                </script>
                """, unsafe_allow_html=True)

                st.success(f"✅ Olá, **{nome_oficial}**! Carregando seu painel...")
                time.sleep(0.4)
                st.rerun()

    with st.expander("📲 Como instalar este App na tela inicial do celular", expanded=False):
        st.markdown("""
        <div style="font-size: 12.5px; color: #CBD5E1; line-height: 1.55;">
            🤖 <b>Android (Chrome):</b> Toque nos <b>3 pontinhos (⋮)</b> no topo direito ➔ <b>"Instalar aplicativo"</b>.<br>
            🍎 <b>iPhone (Safari):</b> Toque em <b>Compartilhar (⎋)</b> no rodapé ➔ <b>"Adicionar à Tela de Início"</b>.
        </div>
        """, unsafe_allow_html=True)

    st.stop()

# ====================================================================
# 1. TOPO COMPACTO "TUDO-EM-UM" (LOGO + AO VIVO + CRACHÁ + EFETIVO)
# ====================================================================
st.markdown(f"""
<script>
    localStorage.setItem('enesa_auth_mat', '{st.session_state.mat_apontador}');
    localStorage.setItem('enesa_auth_tk', '{_gerar_token_auth_mat(st.session_state.mat_apontador)}');
</script>
""", unsafe_allow_html=True)

@st.fragment(run_every=5)
def _cabecalho_mobile_unificado(total_ativos_obra):
    mon = obter_monitor_mobile_tempo_real()
    h_nuvem = mon.get("hash_pde", "")
    h_local = st.session_state.get("_hash_mob_pde", "")
    if h_nuvem and h_nuvem != h_local:
        st.session_state["_hash_mob_pde"] = h_nuvem
        st.cache_data.clear()
        st.rerun(scope="app")

    hora_sync = mon.get("ultima_sync_str") or datetime.datetime.now().strftime("%H:%M:%S")
    efetivo_fmt = f"{total_ativos_obra:,}".replace(",", ".")
    nome_logado = st.session_state.get("nome_apontador_curto", "")
    mat_logada = st.session_state.get("mat_apontador", "")
    func_logada = st.session_state.get("funcao_apontador", "")

    st.markdown(f"""
    <div class="sgo-unified-header">
        <div class="sgo-header-top">
            <div style="display: flex; align-items: center; gap: 10px;">
                <img src="data:image/png;base64,{B64_FAVICON_SGO}" style="width: 36px; height: 36px; border-radius: 9px; border: 1px solid rgba(56,189,248,0.4);">
                <div>
                    <div style="font-size: 9.5px; color: #38BDF8; font-weight: 700; letter-spacing: 0.7px; text-transform: uppercase;">SGO RDC &amp; PDE · ENESA</div>
                    <div style="font-size: 1.05rem; font-weight: 800; color: #F8FAFC; font-family: 'Outfit', sans-serif; line-height: 1.15;">📱 Apontador de Campo</div>
                </div>
            </div>
            <div style="text-align: right; display: flex; flex-direction: column; align-items: flex-end; gap: 3px;">
                <div style="display: flex; gap: 5px; align-items: center;">
                    <span class="badge-cc" style="font-size: 10.5px; padding: 2px 8px;">👥 {efetivo_fmt} ativos</span>
                    <span class="badge-active" style="font-size: 10px; padding: 2px 8px;">🟢 {hora_sync}</span>
                </div>
            </div>
        </div>
        <div class="sgo-header-user">
            <div style="min-width: 0; flex: 1;">
                <div style="font-size: 10px; color: #34D399; font-weight: 700; letter-spacing: 0.4px;">🔐 EM USO POR:</div>
                <div style="font-size: 13px; font-weight: 700; color: #F8FAFC; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">👤 {nome_logado}</div>
                <div style="font-size: 11px; color: #94A3B8;">🪪 Mat: <b style="color:#38BDF8;">{mat_logada}</b> • {func_logada}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

_cabecalho_mobile_unificado(len(df_ativos))

# Lista limpa e completa de TODOS os encarregados ativos (inclusive os que estão com 0 pessoas / sem equipe!)
mapa_enc_mat = pde_engine.obter_mapa_encarregados_matricula()
cargos_gestao = ["SUPERVIS", "COORDENAD", "GERENTE", "DIRETOR", "ENGENHEIR"]

col_funcao = "FUNÇÃO" if "FUNÇÃO" in df_pde.columns else ("FUNCAO" if "FUNCAO" in df_pde.columns else None)
mapa_funcao_pde = dict(zip(df_pde["NOME"].astype(str).str.strip().str.upper(), df_pde[col_funcao])) if col_funcao else {}
nomes_desligados = set(
    df_pde[df_pde["STATUS"].astype(str).str.strip().str.upper() != "ATIVO"]["NOME"]
    .dropna().astype(str).str.strip().str.upper()
)

# 1. Todos que já aparecem na coluna ENCARREGADO (mesmo com 1 ou 2 pessoas)
enc_da_coluna = set([
    str(e).strip().upper() for e in df_ativos["ENCARREGADO"].dropna().unique()
    if str(e).strip() and str(e).strip().upper() not in ["-", "0", "NAN", "NONE", "SEM ENCARREGADO"]
])

# 2. Todos os colaboradores ATIVOS cuja FUNÇÃO seja ENCARREGADO ou LÍDER (mesmo com 0 pessoas na equipe!)
enc_por_funcao = set()
if col_funcao and col_funcao in df_ativos.columns:
    mask_func_enc = df_ativos[col_funcao].astype(str).str.upper().str.contains("ENCARREGAD|LIDER|LÍDER", na=False, regex=True)
    enc_por_funcao = set(
        df_ativos.loc[mask_func_enc, "NOME"].dropna().astype(str).str.strip().str.upper().unique()
    )

# 3. Encarregados salvos no encarregados_f1.json (desde que não estejam desligados)
enc_do_f1_json = set()
arq_f1_j = os.path.join(pde_engine.PASTA_BASE, "encarregados_f1.json")
if os.path.exists(arq_f1_j):
    try:
        with open(arq_f1_j, "r", encoding="utf-8") as _ff1:
            dados_f1 = json.load(_ff1)
            if isinstance(dados_f1, list):
                enc_do_f1_json = set(str(x).strip().upper() for x in dados_f1 if str(x).strip())
            elif isinstance(dados_f1, dict):
                for v_list in dados_f1.values():
                    if isinstance(v_list, list):
                        enc_do_f1_json.update(str(x).strip().upper() for x in v_list if str(x).strip())
    except Exception:
        pass

todos_enc_bruto = (enc_da_coluna | enc_por_funcao | enc_do_f1_json) - nomes_desligados

encarregados_ativos = sorted([
    e for e in todos_enc_bruto
    if e and e not in ["-", "0", "NAN", "NONE", "SEM ENCARREGADO"]
    and "ABMAEL" not in e.upper()
    and not any(cg in e.upper() for cg in cargos_gestao)
    and not any(cg in str(mapa_funcao_pde.get(e.upper(), "")).upper() for cg in cargos_gestao)
])

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

ccs_unicos = sorted([
    str(c).strip() for c in df_ativos["C.C"].dropna().unique() 
    if str(c).strip() and str(c) not in ["-", "0", "nan", "NAN"]
])
opcoes_cc_formatadas = [formatar_cc_display(c) for c in ccs_unicos]

# ====================================================================
# ⭐ BARRA COMPACTA: MEUS ENCARREGADOS + TROCAR USUÁRIO (SEM EMPURRAR A TELA)
# ====================================================================
ARQ_FAVORITOS_APONTADOR = os.path.join(pde_engine.PASTA_BASE, "favoritos_apontadores.json")

def carregar_favoritos_apontadores():
    if os.path.exists(ARQ_FAVORITOS_APONTADOR):
        try:
            with open(ARQ_FAVORITOS_APONTADOR, "r", encoding="utf-8") as _ff:
                return json.load(_ff)
        except Exception:
            return {}
    return {}

def salvar_favoritos_apontador(chave_ap, lista_encs):
    dados = carregar_favoritos_apontadores()
    chave = (chave_ap or "PADRAO").strip().upper()
    dados[chave] = [str(e).strip() for e in lista_encs if str(e).strip()]
    try:
        with open(ARQ_FAVORITOS_APONTADOR, "w", encoding="utf-8") as _ff:
            json.dump(dados, _ff, ensure_ascii=False, indent=2)
    except Exception:
        pass

chave_apontador_atual = (st.session_state.get("mat_apontador") or st.session_state.get("nome_apontador") or "PADRAO").strip().upper()
mapa_todos_favoritos = carregar_favoritos_apontadores()

param_favs = st.query_params.get("favs", "").strip()
favs_salvos_servidor = mapa_todos_favoritos.get(chave_apontador_atual, [])
if not favs_salvos_servidor:
    favs_salvos_servidor = mapa_todos_favoritos.get(st.session_state.get("nome_apontador_curto", "").strip().upper(), [])
if not favs_salvos_servidor and param_favs:
    favs_salvos_servidor = [x.strip() for x in param_favs.split("|") if x.strip() in encarregados_ativos]
    if favs_salvos_servidor:
        salvar_favoritos_apontador(chave_apontador_atual, favs_salvos_servidor)

meus_encarregados_fav = [e for e in favs_salvos_servidor if e in encarregados_ativos]

col_top_fav, col_top_out = st.columns([2.5, 1])
with col_top_fav:
    with st.expander(f"⭐ Meus Encarregados ({len(meus_encarregados_fav)}) · 📲 App", expanded=False):
        st.caption(f"Rota salva na matrícula **{st.session_state.mat_apontador}**:")
        novos_favs_sel = st.multiselect(
            "Seus Encarregados:",
            options=encarregados_ativos,
            default=meus_encarregados_fav,
            placeholder="⭐ Escolha seus encarregados...",
            key="ms_meus_encarregados_fav"
        )
        c_f1, c_f2 = st.columns([2, 1])
        with c_f1:
            if st.button("💾 Salvar Rota", type="primary", use_container_width=True, key="btn_salvar_favs_ap"):
                salvar_favoritos_apontador(chave_apontador_atual, novos_favs_sel)
                st.query_params["favs"] = "|".join(novos_favs_sel)
                st.toast(f"⭐ {len(novos_favs_sel)} encarregado(s) favoritados!", icon="⭐")
                st.rerun()
        with c_f2:
            if meus_encarregados_fav and st.button("🗑️ Limpar", use_container_width=True, key="btn_limpar_favs_ap"):
                salvar_favoritos_apontador(chave_apontador_atual, [])
                if "favs" in st.query_params:
                    del st.query_params["favs"]
                st.rerun()
        st.markdown("---")
        st.caption("📲 **Instalar no celular:** No Chrome toque em **⋮ ➔ Instalar aplicativo**. No iPhone toque em **Compartilhar ➔ Adicionar à Tela de Início**.")
with col_top_out:
    if st.button("🚪 Sair", key="btn_logout_apontador", use_container_width=True):
        st.session_state.apontador_autenticado = False
        st.session_state.mat_apontador = ""
        st.session_state.nome_apontador_curto = ""
        st.session_state.funcao_apontador = ""
        st.session_state.nome_apontador = ""
        if "auth_mat" in st.query_params:
            del st.query_params["auth_mat"]
        if "auth_tk" in st.query_params:
            del st.query_params["auth_tk"]
        st.markdown("""
        <script>
            localStorage.removeItem('enesa_auth_mat');
            localStorage.removeItem('enesa_auth_tk');
        </script>
        """, unsafe_allow_html=True)
        st.rerun()

if meus_encarregados_fav:
    outros_encarregados = [e for e in encarregados_ativos if e not in meus_encarregados_fav]
    encarregados_ordenados_top = meus_encarregados_fav + outros_encarregados
    
    filtro_visao_rota = st.radio(
        "Filtro Rápido de Encarregados:",
        [f"⭐ Minha Rota ({len(meus_encarregados_fav)})", f"🌐 Todos ({len(encarregados_ativos)})"],
        horizontal=True,
        label_visibility="collapsed",
        key="radio_filtro_meus_encs"
    )
    if filtro_visao_rota.startswith("⭐"):
        lista_enc_ativa_tela = meus_encarregados_fav
    else:
        lista_enc_ativa_tela = encarregados_ordenados_top
else:
    lista_enc_ativa_tela = encarregados_ativos

def formatar_enc_estrela(nome_enc):
    if nome_enc == OPCAO_SEM_ENC:
        return nome_enc
    qtd_eq = int((df_ativos["ENCARREGADO"].astype(str).str.strip().str.upper() == str(nome_enc).strip().upper()).sum())
    prefixo = "⭐ " if nome_enc in meus_encarregados_fav else "👷 "
    sufixo = f"({qtd_eq} pess.)" if qtd_eq > 0 else "(0 pess. - Sem Equipe)"
    return f"{prefixo}{nome_enc} {sufixo}"

opcoes_enc_com_desvincular = [OPCAO_SEM_ENC] + lista_enc_ativa_tela

OPCOES_TURNO_OFICIAIS = ["1° TURNO", "2° TURNO", "3° TURNO"]

# ====================================================================
# 2. NAVEGAÇÃO PRINCIPAL ESTILO "APP NATIVO" (5 PÍLULAS LADO A LADO)
# ====================================================================
tab_remanejar, tab_turno, tab_equipe, tab_imprimir, tab_historico = st.tabs([
    "🔄\nRemanejar",
    "🕒\nTurno",
    "📋\nEquipes",
    "🖨️\nRDC",
    "📜\nHistórico"
])

# ====================================================================
# ABA 1: REMANEJAR (FLUXO DIRETO POR ENCARREGADO OU POR BUSCA)
# ====================================================================
with tab_remanejar:
    
    modo_remanejo = st.radio(
        "Modo de remanejamento:",
        ["👷 Por Encarregado (Lote)", "🔍 Buscar 1 a 1"],
        horizontal=True,
        label_visibility="collapsed",
        key="radio_modo_remanejo"
    )

    # -------------------------------------------------------------
    # FLUXO 1: SELECIONAR ENCARREGADO E ADICIONAR PESSOAS
    # -------------------------------------------------------------
    if modo_remanejo.startswith("👷"):
        idx_padrao_rem = 1 if (meus_encarregados_fav and len(opcoes_enc_com_desvincular) > 1) else 0
        enc_alvo = st.selectbox(
            "1. Encarregado de Destino:",
            opcoes_enc_com_desvincular,
            index=idx_padrao_rem,
            format_func=formatar_enc_estrela,
            key="sel_enc_alvo_remanejo"
        )
        
        if enc_alvo == OPCAO_SEM_ENC:
            st.markdown("""
            <div class="leader-box" style="border-color: rgba(239, 68, 68, 0.45); background: rgba(239, 68, 68, 0.1);">
                <div style="font-weight: 700; color: #F87171; font-size: 1rem;">🚫 MODO DESVINCULAÇÃO</div>
                <div style="font-size: 0.82rem; color: #CBD5E1; margin-top: 2px;">Os colaboradores selecionados ficarão <b>sem encarregado</b> (Coluna K vazia).</div>
            </div>
            """, unsafe_allow_html=True)
            cc_sugerido_fmt = "(Manter Atual)"
            membros_atuais = df_ativos[df_ativos["ENCARREGADO"].isna() | df_ativos["ENCARREGADO"].isin(["", "-", "nan", "NAN", "SEM ENCARREGADO"])]
        else:
            mat_lider = mapa_enc_mat.get(enc_alvo.upper(), "")
            membros_atuais = df_ativos[df_ativos["ENCARREGADO"] == enc_alvo]
            
            cc_lider_series = membros_atuais["C.C"].dropna()
            cc_padrao_raw = cc_lider_series.iloc[0] if not cc_lider_series.empty else (ccs_unicos[0] if ccs_unicos else "")
            cc_sugerido_fmt = formatar_cc_display(cc_padrao_raw)
            
            st.markdown(f"""
            <div class="leader-box">
                <div style="display: flex; justify-content: space-between; align-items: flex-start; gap: 8px;">
                    <div>
                        <div style="font-weight: 800; color: #38BDF8; font-size: 1.05rem; font-family: 'Outfit', sans-serif;">👷 {enc_alvo}</div>
                        <div style="display: flex; flex-wrap: wrap; gap: 6px; margin-top: 6px; align-items: center;">
                            <span class="badge-mat">#{mat_lider or 'S/M'}</span>
                            <span class="badge-cc">🏢 {cc_sugerido_fmt}</span>
                        </div>
                    </div>
                    <span class="badge-active" style="font-size: 12px; padding: 4px 10px; white-space: nowrap;">👥 {len(membros_atuais)} pess.</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

        col_c1, col_c2 = st.columns(2)
        opcoes_cc_com_manter = ["(Manter Atual)"] + opcoes_cc_formatadas
        idx_cc_padrao = opcoes_cc_com_manter.index(cc_sugerido_fmt) if cc_sugerido_fmt in opcoes_cc_com_manter else 0
        
        with col_c1:
            cc_escolhido_fmt = st.selectbox(
                "🏢 Centro de Custo:",
                opcoes_cc_com_manter,
                index=idx_cc_padrao,
                key="sel_cc_remanejo_multi"
            )
        with col_c2:
            turno_escolhido_multi = st.selectbox(
                "🕒 Turno (Col. G):",
                ["(Manter Atual)"] + OPCOES_TURNO_OFICIAIS,
                key="sel_turno_remanejo_multi"
            )
        obs_multi = st.text_input("📝 Observação (Opcional):", placeholder="Ex: Remanejamento de frente...", key="txt_obs_remanejo_multi")

        opcoes_colabs_multiselect = [
            f"{r['MATRICULA']} — {r['NOME']} ({r['FUNÇÃO']}) [Atual: {r['ENCARREGADO'] if r['ENCARREGADO'] and r['ENCARREGADO'] != '-' else 'Sem Enc.'} | {formatar_cc_display(r['C.C'])}]"
            for _, r in df_ativos.sort_values(by="NOME").iterrows()
        ]
        mapa_mat_para_opcao = {
            str(r["MATRICULA"]).strip().replace(".0", ""): f"{r['MATRICULA']} — {r['NOME']} ({r['FUNÇÃO']}) [Atual: {r['ENCARREGADO'] if r['ENCARREGADO'] and r['ENCARREGADO'] != '-' else 'Sem Enc.'} | {formatar_cc_display(r['C.C'])}]"
            for _, r in df_ativos.iterrows()
        }

        lote_matriculas_txt = st.text_input(
            "⚡ Colar / Digitar Várias Matrículas (separadas por espaço ou vírgula):",
            placeholder="Ex: 12502449 12502458 12502460",
            key="txt_lote_matriculas_remanejo"
        )
        opcoes_auto_lote = []
        mats_nao_encontradas = []
        if lote_matriculas_txt.strip():
            tokens_mat = [t.strip().replace(".0", "") for t in re.split(r"[\s,;\n\t|]+", lote_matriculas_txt.strip()) if t.strip()]
            for tk in tokens_mat:
                op_match = mapa_mat_para_opcao.get(tk)
                if not op_match and tk.isdigit():
                    op_match = mapa_mat_para_opcao.get(str(int(tk)))
                if op_match:
                    if op_match not in opcoes_auto_lote:
                        opcoes_auto_lote.append(op_match)
                else:
                    mats_nao_encontradas.append(tk)
            if opcoes_auto_lote:
                st.success(f"✅ **{len(opcoes_auto_lote)} matrícula(s)** reconhecida(s) automaticamente!")
                html_lote_rec = "".join([
                    f"<div style='font-size:12px; color:#E2E8F0; padding:4px 0; border-bottom:1px solid rgba(148,163,184,0.12);'>✔️ <b>{op.split(' [')[0]}</b></div>"
                    for op in opcoes_auto_lote
                ])
                st.markdown(f"<div style='background:rgba(16,185,129,0.09); border:1px solid rgba(16,185,129,0.32); border-radius:10px; padding:8px 12px; margin-bottom:10px;'>{html_lote_rec}</div>", unsafe_allow_html=True)
            if mats_nao_encontradas:
                st.warning(f"⚠️ Matrícula(s) não encontrada(s) nos ativos: {', '.join(mats_nao_encontradas)}")

        selecionados_manual = st.multiselect(
            "2. Ou busque por Nome / Matrícula:",
            opcoes_colabs_multiselect,
            default=opcoes_auto_lote,
            placeholder="🔎 Digite nome ou matrícula...",
            key="ms_colaboradores_remanejo"
        )
        selecionados = list(dict.fromkeys(opcoes_auto_lote + selecionados_manual))
        
        if selecionados:
            if enc_alvo == OPCAO_SEM_ENC:
                btn_txt = f"🚫 Desvincular {len(selecionados)} Colaborador(es)"
                btn_tipo = "secondary"
            else:
                btn_txt = f"⚡ Alocar {len(selecionados)} Colaborador(es) com {enc_alvo}"
                btn_tipo = "primary"
                
            if st.button(btn_txt, type=btn_tipo, use_container_width=True, key="btn_confirmar_remanejo_multi"):
                enc_final = "" if enc_alvo == OPCAO_SEM_ENC else enc_alvo
                cc_codigo_limpo = cc_escolhido_fmt.split(" — ")[0].strip() if cc_escolhido_fmt != "(Manter Atual)" else None
                
                with st.spinner("Atualizando planilha e registrando no seu nome..."):
                    lista_alt = []
                    for item_str in selecionados:
                        mat_str = item_str.split(" — ")[0].strip()
                        nome_colab = ""
                        membros_match = df_ativos[df_ativos["MATRICULA"].astype(str) == mat_str]
                        if not membros_match.empty:
                            nome_colab = str(membros_match.iloc[0]["NOME"])
                            
                        item_dict = {
                            "MATRICULA": mat_str,
                            "NOME": nome_colab,
                            "NOVO_ENCARREGADO": enc_final
                        }
                        if cc_codigo_limpo:
                            item_dict["NOVO_CC"] = cc_codigo_limpo
                        if turno_escolhido_multi != "(Manter Atual)":
                            item_dict["NOVO_TURNO"] = turno_escolhido_multi
                        lista_alt.append(item_dict)
                        
                    ok, msg = pde_engine.aplicar_correcoes_em_massa(lista_alt, sync_sheets=True)
                    if ok:
                        for item_dict in lista_alt:
                            mat_it = item_dict["MATRICULA"]
                            m_rows = df_ativos[df_ativos["MATRICULA"].astype(str) == mat_it]
                            if not m_rows.empty:
                                row_it = m_rows.iloc[0]
                                pde_engine.registrar_movimentacao_equipe(
                                    matricula=mat_it,
                                    nome=row_it["NOME"],
                                    de_enc=str(row_it.get("ENCARREGADO", "")),
                                    para_enc=enc_final if enc_final else "(Sem Encarregado)",
                                    de_cc=str(row_it.get("C.C", "")),
                                    para_cc=cc_codigo_limpo if cc_codigo_limpo else str(row_it.get("C.C", "")),
                                    apontador=st.session_state.nome_apontador,
                                    obs=obs_multi
                                )
                        st.success(f"🎉 **{len(selecionados)} colaborador(es)** alocado(s) por **{st.session_state.nome_apontador}**!\n\n{msg}")
                        st.balloons()
                        _invalidar_cache_apos_gravacao()
                        time.sleep(1.2)
                        st.rerun()
                    else:
                        st.error(f"Erro: {msg}")

        # Mostrar membros que já estão na equipe deste encarregado
        if enc_alvo != OPCAO_SEM_ENC and not membros_atuais.empty:
            with st.expander(f"📋 Ver equipe atual de {enc_alvo} ({len(membros_atuais)} pessoas)"):
                for _, r_m in membros_atuais.sort_values(by="NOME").iterrows():
                    b_turno_m = badge_turno_html(r_m.get("TURNO", "1° TURNO"))
                    st.markdown(f"""
                    <div class="worker-card" style="padding: 10px 12px; margin-bottom: 6px;">
                        <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px;">
                            <div>
                                <div style="font-weight:700; color:#F8FAFC; font-size:13px;">{r_m['NOME']}</div>
                                <div style="font-size:11px; color:#94A3B8; margin-top:2px;">{r_m['FUNÇÃO']}</div>
                            </div>
                            <span class="badge-mat">#{r_m['MATRICULA']}</span>
                        </div>
                        <div style="display:flex; flex-wrap:wrap; gap:6px; margin-top:7px;">
                            {b_turno_m}
                            <span class="badge-cc">🏢 {formatar_cc_display(r_m['C.C'])}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button(f"🚫 Tirar {r_m['NOME'].split()[0]} da Equipe", key=f"btn_quick_tirar_{r_m['MATRICULA']}", use_container_width=True):
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
                                apontador=st.session_state.nome_apontador,
                                obs="Desvinculado da equipe"
                            )
                            st.warning(f"🚫 **{r_m['NOME']}** desvinculado!\n\n{msg_t}")
                            _invalidar_cache_apos_gravacao()
                            time.sleep(1.2)
                            st.rerun()

    # -------------------------------------------------------------
    # FLUXO 2: BUSCA POR COLABORADOR (1 A 1)
    # -------------------------------------------------------------
    else:
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
                opcoes_colab = [f"{r['MATRICULA']} — {r['NOME']} ({r['FUNÇÃO']})" for _, r in res_ind.iterrows()]
                escolha = st.selectbox(f"Colaborador ({len(res_ind)} encontrados):", opcoes_colab, key="sel_colab_ind_v2")
                
                mat_escolhida = escolha.split(" — ")[0].strip()
                row_c = res_ind[res_ind["MATRICULA"].astype(str) == mat_escolhida].iloc[0]
                
                enc_atual = str(row_c.get("ENCARREGADO", "")).strip()
                tem_enc = enc_atual and enc_atual not in ["-", "0", "nan", "NAN", "SEM ENCARREGADO"]
                enc_display = enc_atual if tem_enc else "Sem Encarregado"
                
                cc_atual_raw = str(row_c.get("C.C", "")).strip()
                cc_atual_display = formatar_cc_display(cc_atual_raw)
                turno_atual_colab = str(row_c.get("TURNO", "1° TURNO")).strip().upper() or "1° TURNO"
                
                badge_enc_class = "badge-active" if tem_enc else "badge-none"
                
                st.markdown(f"""
                <div class="worker-card">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px;">
                        <div>
                            <div style="font-size: 1.05rem; font-weight: 800; color: #F8FAFC;">{row_c['NOME']}</div>
                            <div style="font-size: 0.82rem; color: #94A3B8; margin-top: 2px;">{row_c['FUNÇÃO']}</div>
                        </div>
                        <span class="badge-mat">#{row_c['MATRICULA']}</span>
                    </div>
                    <div style="display: flex; gap: 6px; margin-top: 10px; flex-wrap: wrap;">
                        <span class="{badge_enc_class}">👷 {enc_display}</span>
                        {badge_turno_html(turno_atual_colab)}
                        <span class="badge-cc">🏢 {cc_atual_display}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                idx_enc_padrao = opcoes_enc_com_desvincular.index(enc_atual) if enc_atual in opcoes_enc_com_desvincular else 0
                novo_enc_escolhido = st.selectbox("👷 Novo Encarregado:", opcoes_enc_com_desvincular, index=idx_enc_padrao, format_func=formatar_enc_estrela, key="sel_novo_enc_ind_v2")
                
                c_i1, c_i2 = st.columns(2)
                with c_i1:
                    idx_cc_ind = opcoes_cc_formatadas.index(cc_atual_display) if cc_atual_display in opcoes_cc_formatadas else 0
                    novo_cc_escolhido_fmt = st.selectbox("🏢 Novo C.C:", opcoes_cc_formatadas, index=idx_cc_ind, key="sel_novo_cc_ind_v2")
                with c_i2:
                    opcoes_turno_ind = ["(Manter Atual)"] + OPCOES_TURNO_OFICIAIS
                    idx_turno_ind = opcoes_turno_ind.index(turno_atual_colab) if turno_atual_colab in opcoes_turno_ind else 0
                    novo_turno_ind = st.selectbox("🕒 Turno (Col. G):", opcoes_turno_ind, index=idx_turno_ind, key="sel_novo_turno_ind_v2")
                
                obs_ind = st.text_input("📝 Observação (Opcional):", placeholder="Ex: Reforço, remanejamento...", key="txt_obs_ind_v2")
                
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    btn_confirmar = st.button("⚡ Salvar Alteração", type="primary", use_container_width=True, key="btn_transf_ind_v2")
                with col_b2:
                    btn_remover_enc = st.button("🚫 Deixar S/ Enc.", use_container_width=True, key="btn_remover_enc_ind_v2")
                    
                if btn_confirmar or btn_remover_enc:
                    enc_final = "" if (btn_remover_enc or novo_enc_escolhido == OPCAO_SEM_ENC) else novo_enc_escolhido
                    cc_final_raw = novo_cc_escolhido_fmt.split(" — ")[0].strip()
                    
                    with st.spinner("Gravando alteração e sincronizando..."):
                        lista_alt = [{
                            "MATRICULA": row_c["MATRICULA"],
                            "NOME": row_c["NOME"],
                            "NOVO_ENCARREGADO": enc_final,
                            "NOVO_CC": cc_final_raw,
                            "NOVO_TURNO": (novo_turno_ind if novo_turno_ind != "(Manter Atual)" else None)
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
                                apontador=st.session_state.nome_apontador,
                                obs=obs_ind
                            )
                            if enc_final:
                                st.success(f"🎉 **{row_c['NOME']}** vinculado a **{enc_final}**!\n\n{msg}")
                            else:
                                st.warning(f"🚫 **{row_c['NOME']}** agora está **sem encarregado**!\n\n{msg}")
                            st.balloons()
                            _invalidar_cache_apos_gravacao()
                            time.sleep(1.2)
                            st.rerun()
                        else:
                            st.error(f"Erro ao salvar: {msg}")
        else:
            st.info("👆 Digite o nome ou a matrícula acima para localizar o colaborador.")

# ====================================================================
# ABA 2: GESTÃO DE TURNO (COLUNA G — 1° TURNO, 2° TURNO E 3° TURNO)
# ====================================================================
with tab_turno:
    s_turno_norm = df_ativos["TURNO"].fillna("").astype(str).str.strip().str.upper()
    qtd_t1 = int(s_turno_norm.str.contains("1", na=False).sum())
    qtd_t2 = int(s_turno_norm.str.contains("2", na=False).sum())
    qtd_t3 = int(s_turno_norm.str.contains("3", na=False).sum())

    # Cards horizontais coloridos para os 3 turnos (não quebram no celular)
    st.markdown(f"""
    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-bottom: 12px;">
        <div style="background: rgba(56, 189, 248, 0.12); border: 1px solid rgba(56, 189, 248, 0.35); border-radius: 12px; padding: 8px 6px; text-align: center;">
            <div style="font-size: 10.5px; color: #38BDF8; font-weight: 700;">☀️ 1° TURNO</div>
            <div style="font-size: 1.35rem; font-weight: 800; color: #F8FAFC; font-family: 'Outfit', sans-serif;">{qtd_t1}</div>
        </div>
        <div style="background: rgba(245, 158, 11, 0.12); border: 1px solid rgba(245, 158, 11, 0.35); border-radius: 12px; padding: 8px 6px; text-align: center;">
            <div style="font-size: 10.5px; color: #FBBF24; font-weight: 700;">🌆 2° TURNO</div>
            <div style="font-size: 1.35rem; font-weight: 800; color: #F8FAFC; font-family: 'Outfit', sans-serif;">{qtd_t2}</div>
        </div>
        <div style="background: rgba(168, 85, 247, 0.14); border: 1px solid rgba(168, 85, 247, 0.38); border-radius: 12px; padding: 8px 6px; text-align: center;">
            <div style="font-size: 10.5px; color: #C084FC; font-weight: 700;">🌙 3° TURNO</div>
            <div style="font-size: 1.35rem; font-weight: 800; color: #F8FAFC; font-family: 'Outfit', sans-serif;">{qtd_t3}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    modo_turno = st.radio(
        "Modo de alteração de turno:",
        ["👷 Por Equipe do Encarregado", "🔎 Colaboradores Avulsos / Lote"],
        horizontal=True,
        label_visibility="collapsed",
        key="radio_modo_aba_turno"
    )

    if modo_turno.startswith("👷"):
        enc_turno_sel = st.selectbox(
            "1. Escolha o Encarregado:",
            lista_enc_ativa_tela,
            format_func=formatar_enc_estrela,
            key="sel_enc_aba_turno"
        )
        eq_turno_df = df_ativos[df_ativos["ENCARREGADO"] == enc_turno_sel].sort_values(by="NOME")

        if eq_turno_df.empty:
            st.info("Nenhum colaborador ativo vinculado a este encarregado.")
        else:
            marcar_toda_eq_turno = st.checkbox(
                f"✅ Aplicar para TODA a equipe de {enc_turno_sel} ({len(eq_turno_df)} pessoas)",
                value=True,
                key="chk_toda_eq_turno"
            )
            if marcar_toda_eq_turno:
                mats_alvo_turno = list(eq_turno_df["MATRICULA"].astype(str))
            else:
                opcoes_eq_turno = [
                    f"{r['MATRICULA']} — {r['NOME']} ({r['FUNÇÃO']}) [Atual: {str(r.get('TURNO', '')).strip() or 'S/ Turno'}]"
                    for _, r in eq_turno_df.iterrows()
                ]
                escolhidos_eq_t = st.multiselect(
                    "Selecione os colaboradores da equipe:",
                    opcoes_eq_turno,
                    default=opcoes_eq_turno,
                    key="ms_eq_aba_turno"
                )
                mats_alvo_turno = [x.split(" — ")[0].strip() for x in escolhidos_eq_t]

            novo_turno_eq = st.radio(
                "2. Novo Turno (Coluna G):",
                OPCOES_TURNO_OFICIAIS,
                horizontal=True,
                key="radio_novo_turno_eq"
            )

            if st.button(
                f"💾 Gravar {novo_turno_eq} na Coluna G ({len(mats_alvo_turno)} pessoas)",
                type="primary",
                use_container_width=True,
                key="btn_salvar_turno_eq"
            ):
                if not mats_alvo_turno:
                    st.warning("Selecione pelo menos 1 colaborador.")
                else:
                    with st.spinner(f"Atualizando Coluna G para {novo_turno_eq}..."):
                        lista_alt_t = []
                        for m_t in mats_alvo_turno:
                            r_match = df_ativos[df_ativos["MATRICULA"].astype(str) == str(m_t)]
                            nome_m = str(r_match.iloc[0]["NOME"]) if not r_match.empty else ""
                            lista_alt_t.append({
                                "MATRICULA": m_t,
                                "NOME": nome_m,
                                "NOVO_TURNO": novo_turno_eq
                            })
                        ok_t, msg_t = pde_engine.aplicar_correcoes_em_massa(lista_alt_t, sync_sheets=True)
                        if ok_t:
                            for item_t in lista_alt_t:
                                pde_engine.registrar_movimentacao_equipe(
                                    matricula=item_t["MATRICULA"],
                                    nome=item_t["NOME"],
                                    de_enc=enc_turno_sel,
                                    para_enc=enc_turno_sel,
                                    de_cc="-",
                                    para_cc="-",
                                    apontador=st.session_state.nome_apontador,
                                    obs=f"Turno (Col G) -> {novo_turno_eq}"
                                )
                            st.success(f"🎉 **{len(lista_alt_t)} colaborador(es)** atualizados para **{novo_turno_eq}**!\n\n{msg_t}")
                            _invalidar_cache_apos_gravacao()
                            time.sleep(1.2)
                            st.rerun()
                        else:
                            st.error(f"Erro ao atualizar turno: {msg_t}")

            with st.expander(f"👀 Conferir Turnos de {enc_turno_sel} ({len(eq_turno_df)} pessoas)"):
                for _, r_et in eq_turno_df.iterrows():
                    t_at = str(r_et.get("TURNO", "")).strip() or "1° TURNO"
                    st.markdown(f"""
                    <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.04); border-radius:10px; padding:8px 12px; margin-bottom:6px;">
                        <div>
                            <b style="color:#F8FAFC; font-size:13px;">{r_et['NOME']}</b>
                            <div style="font-size:11px; color:#94A3B8;">#{r_et['MATRICULA']} • {r_et['FUNÇÃO']}</div>
                        </div>
                        {badge_turno_html(t_at)}
                    </div>
                    """, unsafe_allow_html=True)

    else:
        opcoes_todos_turno = [
            f"{r['MATRICULA']} — {r['NOME']} ({r['FUNÇÃO']}) [Turno Atual: {str(r.get('TURNO', '')).strip() or 'S/ Turno'} | Enc: {r.get('ENCARREGADO', '-')}]"
            for _, r in df_ativos.sort_values(by="NOME").iterrows()
        ]
        mapa_mat_turno_op = {
            str(r["MATRICULA"]).strip().replace(".0", ""): f"{r['MATRICULA']} — {r['NOME']} ({r['FUNÇÃO']}) [Turno Atual: {str(r.get('TURNO', '')).strip() or 'S/ Turno'} | Enc: {r.get('ENCARREGADO', '-')}]"
            for _, r in df_ativos.iterrows()
        }
        lote_mat_turno_txt = st.text_input(
            "⚡ Colar / Digitar Várias Matrículas (separadas por espaço ou vírgula):",
            placeholder="Ex: 12502449 12502458 12502460",
            key="txt_lote_mat_aba_turno"
        )
        auto_lote_turno = []
        if lote_mat_turno_txt.strip():
            for tk in [t.strip().replace(".0", "") for t in re.split(r"[\s,;\n\t|]+", lote_mat_turno_txt.strip()) if t.strip()]:
                op_m = mapa_mat_turno_op.get(tk) or (mapa_mat_turno_op.get(str(int(tk))) if tk.isdigit() else None)
                if op_m and op_m not in auto_lote_turno:
                    auto_lote_turno.append(op_m)
            if auto_lote_turno:
                st.success(f"✅ **{len(auto_lote_turno)} matrícula(s)** reconhecida(s)!")
                html_lote_t = "".join([
                    f"<div style='font-size:12px; color:#E2E8F0; padding:4px 0; border-bottom:1px solid rgba(148,163,184,0.12);'>✔️ <b>{op.split(' [')[0]}</b></div>"
                    for op in auto_lote_turno
                ])
                st.markdown(f"<div style='background:rgba(16,185,129,0.09); border:1px solid rgba(16,185,129,0.32); border-radius:10px; padding:8px 12px; margin-bottom:10px;'>{html_lote_t}</div>", unsafe_allow_html=True)

        sel_avulsos_manual = st.multiselect(
            "Ou busque por Nome / Matrícula:",
            opcoes_todos_turno,
            default=auto_lote_turno,
            placeholder="🔎 Digite nome ou matrícula...",
            key="ms_avulsos_aba_turno"
        )
        sel_avulsos_turno = list(dict.fromkeys(auto_lote_turno + sel_avulsos_manual))
        novo_turno_avulso = st.radio(
            "2. Novo Turno (Coluna G):",
            OPCOES_TURNO_OFICIAIS,
            horizontal=True,
            key="radio_novo_turno_avulso"
        )
        if sel_avulsos_turno:
            if st.button(
                f"💾 Gravar {novo_turno_avulso} na Coluna G ({len(sel_avulsos_turno)} selecionado(s))",
                type="primary",
                use_container_width=True,
                key="btn_salvar_turno_avulso"
            ):
                with st.spinner(f"Gravando {novo_turno_avulso} na Coluna G..."):
                    lista_alt_av = []
                    for item_s in sel_avulsos_turno:
                        m_av = item_s.split(" — ")[0].strip()
                        r_m = df_ativos[df_ativos["MATRICULA"].astype(str) == m_av]
                        n_av = str(r_m.iloc[0]["NOME"]) if not r_m.empty else ""
                        enc_av = str(r_m.iloc[0].get("ENCARREGADO", "")) if not r_m.empty else ""
                        lista_alt_av.append({
                            "MATRICULA": m_av,
                            "NOME": n_av,
                            "ENC_ATUAL": enc_av,
                            "NOVO_TURNO": novo_turno_avulso
                        })
                    ok_av, msg_av = pde_engine.aplicar_correcoes_em_massa(lista_alt_av, sync_sheets=True)
                    if ok_av:
                        for it_av in lista_alt_av:
                            pde_engine.registrar_movimentacao_equipe(
                                matricula=it_av["MATRICULA"],
                                nome=it_av["NOME"],
                                de_enc=it_av["ENC_ATUAL"],
                                para_enc=it_av["ENC_ATUAL"],
                                de_cc="-",
                                para_cc="-",
                                apontador=st.session_state.nome_apontador,
                                obs=f"Turno (Col G) -> {novo_turno_avulso}"
                            )
                        st.success(f"🎉 **{len(lista_alt_av)} colaborador(es)** alterados para **{novo_turno_avulso}**!\n\n{msg_av}")
                        _invalidar_cache_apos_gravacao()
                        time.sleep(1.2)
                        st.rerun()
                    else:
                        st.error(f"Erro ao salvar: {msg_av}")

# ====================================================================
# ABA 3: VISÃO POR EQUIPE (Raio-X do Encarregado & Desvinculação Rápida)
# ====================================================================
with tab_equipe:
    enc_visao = st.selectbox("Selecione o Encarregado:", lista_enc_ativa_tela, format_func=formatar_enc_estrela, key="sel_enc_visao")
    
    if enc_visao:
        equipe = df_ativos[df_ativos["ENCARREGADO"] == enc_visao].sort_values(by="NOME")
        
        st.markdown(f"""
        <div class="leader-box" style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <b style="color: #38BDF8; font-size: 1.05rem; font-family: 'Outfit', sans-serif;">👷 {enc_visao}</b>
                <div style="margin-top: 4px;"><span class="badge-mat">#{mapa_enc_mat.get(enc_visao, 'N/A')}</span></div>
            </div>
            <div style="text-align: right;">
                <span class="badge-active" style="font-size: 13px; padding: 5px 12px;">👥 {len(equipe)} integrantes</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if equipe.empty:
            st.info("Nenhum colaborador vinculado a este encarregado no momento.")
        else:
            for idx, r_eq in equipe.iterrows():
                t_eq = str(r_eq.get("TURNO", "1° TURNO")).strip() or "1° TURNO"
                with st.expander(f"👤 {r_eq['NOME']} • #{r_eq['MATRICULA']}"):
                    st.markdown(f"""
                    <div style="display:flex; flex-wrap:wrap; gap:6px; margin-bottom:8px;">
                        <span class="badge-mat">#{r_eq['MATRICULA']}</span>
                        {badge_turno_html(t_eq)}
                        <span class="badge-cc">🏢 {formatar_cc_display(r_eq['C.C'])}</span>
                    </div>
                    <div style="font-size:12px; color:#94A3B8; margin-bottom:8px;">Função: <b style="color:#F8FAFC;">{r_eq['FUNÇÃO']}</b></div>
                    """, unsafe_allow_html=True)
                    
                    col_eq_a1, col_eq_a2 = st.columns([2, 1])
                    with col_eq_a1:
                        outro_enc = st.selectbox("Transferir para:", [e for e in encarregados_ativos if e != enc_visao], key=f"sel_transf_eq_{r_eq['MATRICULA']}")
                    with col_eq_a2:
                        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                        btn_tr_direta = st.button("Mudar", key=f"btn_tr_{r_eq['MATRICULA']}", type="primary", use_container_width=True)
                        
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
                            _invalidar_cache_apos_gravacao()
                            time.sleep(1.2)
                            st.rerun()

# ====================================================================
# ABA 4: IMPRIMIR RDC (PADRÃO OFICIAL ENESA)
# ====================================================================
with tab_imprimir:
    opcoes_modo_imp = ["⭐ Minha Rota", "1 Encarregado", "Múltiplos / Todos"] if meus_encarregados_fav else ["1 Encarregado", "Múltiplos / Todos"]
    modo_selecao_enc = st.radio("Modo de seleção:", opcoes_modo_imp, horizontal=True, label_visibility="collapsed", key="modo_sel_enc_imp")

    if modo_selecao_enc == "⭐ Minha Rota":
        lista_alvo_imp = meus_encarregados_fav
        st.success(f"⭐ **{len(lista_alvo_imp)} encarregado(s)** da sua rota selecionados!")
    elif modo_selecao_enc == "1 Encarregado":
        enc_unico = st.selectbox("Escolha o Encarregado:", lista_enc_ativa_tela, format_func=formatar_enc_estrela, key="sel_enc_imp_unico")
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
        
        texto_enc_resumo = ", ".join(lista_alvo_imp[:2]) + (f" e mais {len(lista_alvo_imp)-2}" if len(lista_alvo_imp) > 2 else "")
        st.markdown(f"""
        <div class="leader-box" style="margin-top: 6px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div style="font-weight: 800; color: #38BDF8; font-size: 0.98rem;">📄 RDC · RELATÓRIO DIÁRIO DE CAMPO</div>
                <div class="badge-active">A4 OFICIAL</div>
            </div>
            <div style="margin-top: 8px; font-size: 12.5px; color: #CBD5E1; line-height: 1.55;">
                <div><b>Encarregado(s):</b> {texto_enc_resumo} ({len(lista_alvo_imp)})</div>
                <div><b>Emitido por:</b> {st.session_state.nome_apontador}</div>
                <div><b>Efetivo Total:</b> {len(equipe_imp)} colaboradores vinculados</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander(f"👥 Conferir Efetivo ({len(equipe_imp)} colaboradores)", expanded=False):
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
            if st.button("🟢 Excel (.xlsx)", use_container_width=True, key="btn_excel_imp"):
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
# ABA 5: HISTÓRICO DE MOVIMENTAÇÕES E AUDITORIA DE QUEM ESTÁ USANDO
# ====================================================================
with tab_historico:
    df_hist = pde_engine.obter_historico_movimentacoes(limite=80, buscar_nuvem=True)

    if df_hist.empty:
        st.info("Nenhuma movimentação registrada ainda.")
    else:
        contagem_por_ap = df_hist["APONTADOR"].fillna("Não identificado").astype(str).value_counts()
        with st.expander(f"👥 Auditoria: Quem realizou alterações ({len(contagem_por_ap)} usuário(s))", expanded=True):
            for ap_nome, qtd_movs in contagem_por_ap.items():
                st.markdown(f"""
                <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.04); border-radius:8px; padding:6px 12px; margin-bottom:5px; font-size:12.5px;">
                    <span style="color:#F8FAFC; font-weight:600;">👤 {ap_nome}</span>
                    <span class="badge-cc">{qtd_movs} alt.</span>
                </div>
                """, unsafe_allow_html=True)

        opcoes_filtro_ap = ["Todos os Colaboradores"] + list(contagem_por_ap.index)
        filtro_ap_hist = st.selectbox("Filtrar por quem fez a alteração:", opcoes_filtro_ap, key="sel_filtro_ap_hist")
        df_hist_view = df_hist if filtro_ap_hist == "Todos os Colaboradores" else df_hist[df_hist["APONTADOR"].astype(str) == filtro_ap_hist]

        st.caption(f"Exibindo **{len(df_hist_view)}** registro(s) de auditoria:")
        for _, r_h in df_hist_view.iterrows():
            obs_h = str(r_h.get("OBSERVACAO", "")).strip()
            cc_h = str(r_h.get("PARA_CC", "")).strip()
            st.markdown(f"""
            <div class="worker-card" style="border-left: 4px solid #0EA5E9; padding: 10px 14px;">
                <div style="display: flex; justify-content: space-between; align-items: center; color: #94A3B8; font-size: 11px; margin-bottom: 5px; gap: 6px; flex-wrap: wrap;">
                    <span>🕒 {r_h['DATA_HORA']}</span>
                    <span class="badge-cc" style="font-weight:700;">👤 {r_h['APONTADOR']}</span>
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; margin: 4px 0;">
                    <b style="color: #F8FAFC; font-size: 13.5px;">👷 {r_h['NOME']}</b>
                    <span class="badge-mat">#{r_h['MATRICULA']}</span>
                </div>
                <div style="color: #CBD5E1; font-size: 12.5px; margin-top: 3px;">
                    De: <span style="color:#F87171;">{r_h['DE_ENCARREGADO']}</span> ➔ Para: <span style="color:#34D399; font-weight:700;">{r_h['PARA_ENCARREGADO']}</span>
                </div>
                {f"<div style='margin-top: 5px;'><span class='badge-cc'>🏢 {formatar_cc_display(cc_h)}</span></div>" if cc_h and cc_h != '-' else ''}
                {f"<div style='font-size: 11px; color: #FBBF24; margin-top: 4px;'>📝 {obs_h}</div>" if obs_h else ''}
            </div>
            """, unsafe_allow_html=True)

    if st.button("🔄 Atualizar Histórico", use_container_width=True, key="btn_refresh_hist"):
        st.rerun()
