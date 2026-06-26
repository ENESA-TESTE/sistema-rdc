import os
import re
import datetime
import zipfile
import io
import openpyxl
import pandas as pd
import streamlit as st
import altair as alt
import base64
import json
import time
import tempfile
import plotly.express as px
try:
    from google import genai
except ImportError:
    pass

from streamlit_gsheets import GSheetsConnection

# --- CONFIGURAÇÃO DA PÁGINA ---
# Força a criação do arquivo de tema Escuro automaticamente
config_dir = ".streamlit"
os.makedirs(config_dir, exist_ok=True)
config_path = os.path.join(config_dir, "config.toml")
if not os.path.exists(config_path):
    with open(config_path, "w", encoding="utf-8") as f:
        f.write('[theme]\nbase="dark"\nprimaryColor="#f39c12"\nbackgroundColor="#1e1e1e"\nsecondaryBackgroundColor="#2b2b2b"\ntextColor="#e0e4ea"\n')

caminho_nome_site = "nome_empresa.txt"
if os.path.exists(caminho_nome_site):
    with open(caminho_nome_site, "r", encoding="utf-8") as f:
        nome_site = f.read().strip()
    if not nome_site:
        nome_site = "ENESA Engenharia"
else:
    nome_site = "ENESA Engenharia"

st.set_page_config(page_title=f"Sistema RDC & PDE - {nome_site}", layout="wide", initial_sidebar_state="expanded")

# Injeção de CSS para ajustes de interface
st.markdown("""
    <style>
        /* Tentar forçar o carregamento da fonte oficial de ícones do Google */
        @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,0,0');
        
        /* Esconder o botão de collapse da sidebar para evitar texto quebrado */
        [data-testid="stSidebarCollapseButton"] {
            display: none !important;
        }
        
        /* Ocultar texto quebrado "arrow_down" do st.expander caso a fonte não carregue */
        summary .material-symbols-rounded,
        .st-emotion-cache-1t8fpt5 .material-symbols-rounded,
        [data-testid="stExpander"] .material-symbols-rounded {
            display: none !important;
            color: transparent !important;
        }
        
        /* Esconder o menu superior chato do Streamlit (Deploy, Rerun, etc) */
        header { visibility: hidden !important; display: none !important; }
        [data-testid="stHeader"] { display: none !important; }
        [data-testid="stToolbar"] { display: none !important; }
        #MainMenu { display: none !important; }
        footer { display: none !important; }
        .stApp > header { display: none !important; }
        
        /* Remover a linha colorida no topo e subir o layout */
        [data-testid="stDecoration"] { display: none !important; }
        .stApp { margin-top: -60px; }
        
        /* Travar a largura da barra lateral (esconder o arrastador) */
        [data-testid="stSidebarResizer"] {
            display: none !important;
        }
    </style>
""", unsafe_allow_html=True)

# Paleta Premium Dark Mode (Glassmorphism & Neon Subtle)
cor_fundo = "#020617" # Slate 950
cor_fundo_grad = "linear-gradient(135deg, #020617 0%, #0f172a 100%)"
cor_card = "rgba(30, 41, 59, 0.45)" # Slate 800 glass
cor_borda = "rgba(255, 255, 255, 0.08)"
cor_texto = "#f8fafc"
cor_texto_sub = "#94a3b8"
cor_azul = "#3b82f6"
cor_azul_hover = "#2563eb"
cor_destaque = "#0ea5e9" # Light blue neon
cor_verde = "#10b981"

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@500;600;700&display=swap');
    
    @keyframes fadeIn {{
        from {{ opacity: 0; transform: translateY(15px); }}
        to {{ opacity: 1; transform: translateY(0); }}
    }}
    
    /* === BASE === */
    .stApp {{
        font-family: 'Inter', sans-serif !important;
        background: {cor_fundo_grad};
        background-attachment: fixed;
        color: {cor_texto};
    }}
    
    .block-container {{
        animation: fadeIn 0.8s cubic-bezier(0.16, 1, 0.3, 1);
        max-width: 1200px;
        padding-top: 1rem;
    }}

    /* === CABEÇALHO === */
    .enesa-header {{
        text-align: center;
        margin-top: -30px;
        margin-bottom: 28px;
        padding: 32px 20px;
        background: {cor_card};
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border-radius: 16px;
        border: 1px solid {cor_borda};
        border-left: 5px solid {cor_destaque};
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        transition: transform 0.3s ease;
    }}
    .enesa-header:hover {{
        transform: translateY(-2px);
    }}
    
    /* === FUNDO === */
    .stApp {{
        background: radial-gradient(circle at top, {cor_fundo} 0%, #020617 100%) !important;
        background-image: 
            linear-gradient(rgba(14, 165, 233, 0.03) 1px, transparent 1px),
            linear-gradient(90deg, rgba(14, 165, 233, 0.03) 1px, transparent 1px) !important;
        background-size: 30px 30px !important;
    }}
    
    /* === TIPOGRAFIA === */
    h1, h2, h3 {{
        color: {cor_texto} !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        letter-spacing: -0.5px;
    }}
    p, span, label, div {{
        font-family: 'Inter', sans-serif !important;
    }}
    
    /* === BOTÃO PRINCIPAL === */
    div.stButton > button[data-baseweb="button"] {{
        background: linear-gradient(135deg, {cor_azul} 0%, {cor_destaque} 100%);
        color: white;
        border: none;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600;
        font-size: 0.95rem;
        letter-spacing: 0.5px;
        border-radius: 10px;
        padding: 12px 24px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 4px 15px rgba(14, 165, 233, 0.2);
    }}
    div.stButton > button[data-baseweb="button"]:hover {{
        transform: translateY(-2px) scale(1.02);
        box-shadow: 0 8px 25px rgba(14, 165, 233, 0.4);
    }}
    div.stButton > button[data-baseweb="button"]:active {{
        transform: translateY(1px) scale(0.98);
    }}
    
    /* === SIDEBAR === */
    [data-testid="stSidebar"] {{
        background-color: rgba(15, 23, 42, 0.8) !important;
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-right: 1px solid rgba(14, 165, 233, 0.2);
        box-shadow: 4px 0 25px rgba(14, 165, 233, 0.15);
    }}
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {{
        color: {cor_texto} !important;
        font-family: 'Outfit', sans-serif !important;
        text-shadow: 0 0 15px rgba(14, 165, 233, 0.5);
    }}
    
    /* === ADMIN AVATAR GLOW === */
    .admin-avatar {{
        width: 60px; height: 60px;
        border-radius: 50%;
        background: {cor_azul};
        display: flex; align-items: center; justify-content: center;
        margin: 20px auto;
        box-shadow: 0 0 20px {cor_azul};
        color: white; font-weight: bold; font-size: 20px;
        border: 2px solid white;
    }}
    
    /* === INPUTS & CONTAINERS === */
    .stTextInput input, .stSelectbox > div > div, .stTextArea textarea {{
        border-radius: 10px !important;
        border: 1px solid {cor_borda} !important;
        background-color: rgba(15, 23, 42, 0.5) !important;
        color: {cor_texto} !important;
        transition: all 0.3s ease !important;
        backdrop-filter: blur(4px);
    }}
    .stTextInput input:focus, .stSelectbox > div > div:focus-within, .stTextArea textarea:focus {{
        border-color: {cor_destaque} !important;
        box-shadow: 0 0 0 3px rgba(14, 165, 233, 0.2) !important;
        background-color: rgba(30, 41, 59, 0.8) !important;
    }}
    
    /* === ABAS === */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        background-color: {cor_card};
        backdrop-filter: blur(10px);
        border-radius: 12px;
        padding: 6px;
        border: 1px solid {cor_borda};
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 8px;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 500;
        font-size: 0.95rem;
        padding: 10px 20px;
        color: {cor_texto_sub};
        transition: all 0.3s ease;
        border: 1px solid transparent;
    }}
    .stTabs [data-baseweb="tab"]:hover {{
        color: {cor_texto};
        background-color: rgba(255, 255, 255, 0.05);
    }}
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(135deg, {cor_azul} 0%, {cor_destaque} 100%) !important;
        color: white !important;
        box-shadow: 0 4px 15px rgba(14, 165, 233, 0.3);
        border: 1px solid rgba(255,255,255,0.1);
    }}
    
    /* === MÉTRICAS (GLASS CARDS) === */
    [data-testid="stMetric"] {{
        background: {cor_card};
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid {cor_borda};
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative;
        overflow: hidden;
    }}
    [data-testid="stMetric"]::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; width: 100%; height: 100%;
        background: linear-gradient(135deg, rgba(255,255,255,0.05) 0%, rgba(255,255,255,0) 100%);
        pointer-events: none;
    }}
    [data-testid="stMetric"]:hover {{
        transform: translateY(-5px);
        box-shadow: 0 12px 25px rgba(0,0,0,0.3);
        border-color: rgba(14, 165, 233, 0.3);
    }}
    [data-testid="stMetricLabel"] {{
        color: {cor_texto_sub} !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 500 !important;
        letter-spacing: 0.5px;
    }}
    [data-testid="stMetricValue"] {{
        color: {cor_texto} !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 700 !important;
        font-size: 2.2rem !important;
        background: linear-gradient(to right, #fff, #94a3b8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }}
    
    /* === DATAFRAMES === */
    .stDataFrame {{
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid {cor_borda};
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }}
    
    /* === FILE UPLOADER === */
    [data-testid="stFileUploader"] {{
        border-radius: 12px;
        background: {cor_card};
        border: 1px dashed rgba(14, 165, 233, 0.4);
        padding: 10px;
        transition: all 0.3s ease;
    }}
    [data-testid="stFileUploader"]:hover {{
        border-color: {cor_destaque};
        background: rgba(30, 41, 59, 0.6);
    }}
    [data-testid="stFileUploaderDropzone"] button {{
        color: transparent !important;
        position: relative;
        min-width: 140px !important;
    }}
    [data-testid="stFileUploaderDropzone"] button::after {{
        content: "Procurar arquivos";
        color: white;
        position: absolute;
        left: 50%;
        top: 50%;
        transform: translate(-50%, -50%);
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600;
        font-size: 14px;
        white-space: nowrap;
    }}
    
    /* === EXPANDERS === */
    [data-testid="stExpander"] {{
        background: {cor_card};
        backdrop-filter: blur(10px);
        border-radius: 12px !important;
        border: 1px solid {cor_borda} !important;
        overflow: hidden;
    }}
    [data-testid="stExpander"] summary {{
        font-family: 'Outfit', sans-serif !important;
        font-weight: 500;
    }}
    
    /* === ALERTAS === */
    .stAlert {{
        border-radius: 12px;
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255,255,255,0.1);
    }}
    
    /* === SCROLLBAR SUAVE === */
    ::-webkit-scrollbar {{
        width: 8px;
        height: 8px;
    }}
    ::-webkit-scrollbar-track {{
        background: {cor_fundo};
    }}
    ::-webkit-scrollbar-thumb {{
        background: rgba(14, 165, 233, 0.5);
        border-radius: 4px;
    }}
    ::-webkit-scrollbar-thumb:hover {{
        background: {cor_destaque};
    }}
    </style>
""", unsafe_allow_html=True)

# --- CHECAR LOGIN POR LINK RÁPIDO (QR CODE) ---
try:
    if "pwd" in st.query_params and st.query_params["pwd"] == "Campo@2026":
        st.session_state.logged_in = True
        st.session_state.role = "encarregado"
except Exception:
    pass

# --- LOGIN / AUTENTICAÇÃO ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None

if not st.session_state.logged_in:
    col_log1, col_log2, col_log3 = st.columns([1, 2, 1])
    with col_log2:
        st.markdown(f"""
            <div class='enesa-header'>
                <h1 style='color: {cor_azul} !important; font-size: 2.2rem;'>🔒 Acesso Restrito</h1>
                <p style='color: {cor_texto_sub};'>Sistema RDC & PDE - {nome_site}</p>
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            senha_input = st.text_input("Digite a senha de acesso:", type="password")
            btn_login = st.form_submit_button("Entrar no Sistema", use_container_width=True)
            
            if btn_login:
                senha_admin = "Enesa@2026"
                senha_encarregado = "Campo@2026"
                senha_visualizador = "Visualizar@2026"
                if "senha_global" in st.secrets:
                    senha_admin = st.secrets["senha_global"]
                    
                if senha_input == senha_admin:
                    st.session_state.logged_in = True
                    st.session_state.role = "admin"
                    st.rerun()
                elif senha_input == senha_encarregado:
                    st.session_state.logged_in = True
                    st.session_state.role = "encarregado"
                    st.rerun()
                elif senha_input == senha_visualizador:
                    st.session_state.logged_in = True
                    st.session_state.role = "visualizador"
                    st.rerun()
                else:
                    st.error("❌ Senha incorreta! Tente novamente.")
    st.stop() # Bloqueia a renderização do restante do script

# === ANIMAÇÃO DE BOAS-VINDAS ===
if "welcome_shown" not in st.session_state:
    st.session_state.welcome_shown = True
    _hora = datetime.datetime.now().hour
    if _hora < 12:
        _saudacao = "Bom dia"
        _emoji_hora = "☀️"
    elif _hora < 18:
        _saudacao = "Boa tarde"
        _emoji_hora = "🌤️"
    else:
        _saudacao = "Boa noite"
        _emoji_hora = "🌙"
    
    _role_nome = {"admin": "Administrador", "visualizador": "Visualizador", "user": "Encarregado", "encarregado": "Encarregado"}
    _nome_display = st.session_state.get("nome_completo", _role_nome.get(st.session_state.get("role_usuario", st.session_state.get("role", "user")), "Usuário"))
    
    # Tentar contar efetivo
    _efetivo_msg = ""
    try:
        _pasta = os.path.dirname(os.path.abspath(__file__))
        _csv_path = os.path.join(_pasta, "BASE_ATUAL.csv")
        if os.path.exists(_csv_path):
            _df_temp = pd.read_csv(_csv_path)
            _total = len(_df_temp)
            _efetivo_msg = f"<br><span style='font-size: 16px; color: #94a3b8;'>📊 Hoje temos <b style=\"color: #0ea5e9;\">{_total:,}</b> colaboradores na base ativa</span>"
    except Exception:
        pass
    
    st.markdown(f"""
    <div id="welcome-card" style="text-align: center; padding: 30px; background: linear-gradient(145deg, rgba(14, 165, 233, 0.15), rgba(139, 92, 246, 0.15)); border: 1px solid rgba(14, 165, 233, 0.4); border-radius: 16px; margin: 20px auto; max-width: 600px; box-shadow: 0 0 40px rgba(14, 165, 233, 0.2); animation: welcomeSlide 0.8s ease-out;">
        <div style="font-size: 48px; margin-bottom: 10px;">{_emoji_hora}</div>
        <div style="font-size: 28px; font-weight: 800; color: #f8fafc; text-shadow: 0 0 15px rgba(255,255,255,0.2);">{_saudacao}, {_nome_display}!</div>
        <div style="font-size: 14px; color: #64748b; margin-top: 5px;">{datetime.datetime.now().strftime('%A, %d de %B de %Y').capitalize()}</div>
        {_efetivo_msg}
    </div>
    <style>
        @keyframes welcomeSlide {{
            0% {{ opacity: 0; transform: translateY(-30px); }}
            100% {{ opacity: 1; transform: translateY(0); }}
        }}
    </style>
    """, unsafe_allow_html=True)

# Removido o header global daqui para aparecer apenas após o login.

# =================================================================
# CAMINHOS E CONFIGURAÇÕES
# =================================================================
pasta_base = os.path.dirname(os.path.abspath(__file__))
caminho_logo = os.path.join(pasta_base, "logo.png")
caminho_pde_padrao = os.path.join(pasta_base, "PDE.csv")
caminho_modelo_padrao = os.path.join(pasta_base, "MODELO.xlsx")
caminho_modelo_salvo = os.path.join(pasta_base, "MODELO_SALVO.xlsx")

# Caminhos para PERSISTÊNCIA (salvar a base do usuário)
caminho_base_salva_csv = os.path.join(pasta_base, "BASE_ATUAL.csv")
caminho_hist_cc = os.path.join(pasta_base, "historico_cc.csv")
caminho_base_salva_xlsx = os.path.join(pasta_base, "BASE_ATUAL.xlsx")

celula_encarregado = "I4"
celula_matricula = "B9"
celula_nome = "C9"
celula_funcao = "H9"



# =================================================================
# FUNÇÕES UTILITÁRIAS
# =================================================================
def extrair_coordenadas(celula_str):
    match = re.match(r"([A-Z]+)([0-9]+)", celula_str.strip().upper())
    if match:
        return match.group(1), int(match.group(2))
    return None, None

def salvar_historico(wb, nome_arquivo):
    try:
        hoje = datetime.datetime.now()
        pasta_hist = os.path.join(pasta_base, "Historico_RDC", str(hoje.year), f"{hoje.month:02d}_{hoje.strftime('%B')}")
        os.makedirs(pasta_hist, exist_ok=True)
        caminho_hist = os.path.join(pasta_hist, f"{hoje.strftime('%d_%H%M')}_{nome_arquivo}")
        wb.save(caminho_hist)
    except Exception:
        pass

def salvar_modelo_no_disco(arquivo_modelo_up):
    try:
        arquivo_modelo_up.seek(0)
        conteudo = arquivo_modelo_up.read()
        arquivo_modelo_up.seek(0)
        with open(caminho_modelo_salvo, "wb") as f:
            f.write(conteudo)
        return True
    except Exception:
        return False

def obter_caminho_modelo():
    if os.path.exists(caminho_modelo_salvo):
        return caminho_modelo_salvo
    elif os.path.exists(caminho_modelo_padrao):
        return caminho_modelo_padrao
    return None

def preencher_excel(equipe, encarregado_selecionado):
    try:
        caminho = obter_caminho_modelo()
        if caminho is None:
            st.error("⚠️ Ficheiro MODELO.xlsx não encontrado! Faça upload na barra lateral.")
            return None
        
        wb = openpyxl.load_workbook(caminho)
            
        from copy import copy
        ws = wb.active
        celula_enc = ws[celula_encarregado]
        celula_enc.value = encarregado_selecionado
        
        # Ajuste dinâmico de tamanho para nomes longos
        tamanho = 16
        if len(encarregado_selecionado) > 25:
            tamanho = 14
        if len(encarregado_selecionado) > 35:
            tamanho = 12
            
        if celula_enc.font:
            nova_fonte = copy(celula_enc.font)
            nova_fonte.size = tamanho
            nova_fonte.bold = True
            celula_enc.font = nova_fonte
        letra_mat, _ = extrair_coordenadas(celula_matricula)
        letra_nom, linha_nom = extrair_coordenadas(celula_nome)
        letra_fun, _ = extrair_coordenadas(celula_funcao)
        linha_atual = linha_nom if linha_nom else 9
        
        for _, row in equipe.iterrows():
            if letra_mat:
                ws[f"{letra_mat}{linha_atual}"] = str(row.get("MATRICULA", ""))
            if letra_nom:
                ws[f"{letra_nom}{linha_atual}"] = str(row.get("NOME", ""))
            if letra_fun:
                ws[f"{letra_fun}{linha_atual}"] = str(row.get("FUNÇÃO", ""))
            linha_atual += 1
            
        ws.print_area = "A1:R61"
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.horizontalCentered = True
        ws.page_setup.verticalCentered = True
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5
        
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        return wb
    except Exception as e:
        st.error(f"⚠️ Erro ao preencher o modelo Excel: {e}")
        return None

def ler_arquivo_seguro(arquivo, nome_arquivo=""):
    try:
        if nome_arquivo.endswith(".xlsx") or nome_arquivo.endswith(".xls"):
            return pd.read_excel(arquivo)
        else:
            try:
                return pd.read_csv(arquivo, sep=";", encoding="latin-1")
            except Exception:
                pass
            try:
                if hasattr(arquivo, 'seek'):
                    arquivo.seek(0)
                return pd.read_csv(arquivo, sep=",", encoding="latin-1")
            except Exception:
                pass
            try:
                if hasattr(arquivo, 'seek'):
                    arquivo.seek(0)
                return pd.read_csv(arquivo, sep=";", encoding="utf-8")
            except Exception:
                pass
            if hasattr(arquivo, 'seek'):
                arquivo.seek(0)
            return pd.read_csv(arquivo)
    except Exception as e:
        st.sidebar.error(f"❌ Não foi possível ler o arquivo: {e}")
        return None

def salvar_base_localmente(arquivo_upload):
    try:
        nome = arquivo_upload.name
        arquivo_upload.seek(0)
        conteudo = arquivo_upload.read()
        arquivo_upload.seek(0)
        
        if nome.endswith(".xlsx") or nome.endswith(".xls"):
            with open(caminho_base_salva_xlsx, "wb") as f:
                f.write(conteudo)
            if os.path.exists(caminho_base_salva_csv):
                os.remove(caminho_base_salva_csv)
        else:
            with open(caminho_base_salva_csv, "wb") as f:
                f.write(conteudo)
            if os.path.exists(caminho_base_salva_xlsx):
                os.remove(caminho_base_salva_xlsx)
        return True
    except Exception:
        return False

def preparar_dataframe(df):
    mapeamento = {}
    for col in df.columns:
        col_clean = str(col).strip().upper()
        if "ENCARREGADO" in col_clean:
            mapeamento[col] = "ENCARREGADO"
        elif col_clean == "NOME":
            mapeamento[col] = "NOME"
        elif col_clean in ["MATRICULA", "MATRÍCULA"]:
            mapeamento[col] = "MATRICULA"
        elif col_clean in ["FUNÇÃO", "FUNCAO", "FUNCOES", "FUNÇÕES"]:
            mapeamento[col] = "FUNÇÃO"
        elif "CENTRO DE CUSTO" in col_clean or col_clean == "C.C" or col_clean == "CC":
            mapeamento[col] = "C.C"
        elif col_clean in ["DISCIPLINA"]:
            mapeamento[col] = "DISCIPLINA"
        elif "MÃO DE OBRA" in col_clean or "MAO DE OBRA" in col_clean:
            mapeamento[col] = "MÃO DE OBRA"
    
    df = df.rename(columns=mapeamento)
    
    # Remover colunas duplicadas que podem ser geradas após o rename ou pela injeção da nuvem, mantendo a mais recente (last)
    df = df.loc[:, ~df.columns.duplicated(keep='last')]
    
    for c in ["MATRICULA", "NOME", "FUNÇÃO", "ENCARREGADO"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
        df[c] = df[c].replace(["nan", "NaN", "None", "0.0", "0", "#N/D", "#N/A", "#REF!", "-"], "")

    for c in ["C.C", "DISCIPLINA", "MÃO DE OBRA"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
        df[c] = df[c].replace(["nan", "NaN", "None", "0.0", "0", "#N/D", "#N/A", "#REF!", "-"], "")

    df["MATRICULA"] = df["MATRICULA"].str.replace(".0", "", regex=False)
    
    # -------------------------------------------------------------------------
    # NOVO: NORMALIZAR NOMES DE C.C VINDOS DO ERP EM FORMATO DE TEXTO
    # Ex: "PB - Dutos" -> "125.02.002"
    # -------------------------------------------------------------------------
    def normalizar_cc(valor):
        val = str(valor).upper()
        if not val or val.startswith("125."):
            return valor
            
        prefixo = ""
        if "PB" in val: prefixo = "125.02"
        elif "RB" in val: prefixo = "125.01"
            
        sufixo = ""
        if "DUTO" in val: sufixo = "002"
        elif "EQUIPA" in val: sufixo = "001"
        elif "TUBU" in val and "PRESS" not in val: sufixo = "003"
        elif "ESTRUTURA" in val: sufixo = "004"
        elif "PRECIP" in val: sufixo = "005"
        elif "PRESS" in val and "MEC" in val: sufixo = "006"
        elif "PRESS" in val and "TUBU" in val: sufixo = "007"
        elif "PRESS" in val and "FORN" in val: sufixo = "008"
        elif "PINTURA" in val: sufixo = "009"
        elif "COMIS" in val: sufixo = "010"
        elif "ASSISTIDA" in val: sufixo = "011"
        elif "LAVAGEM" in val: sufixo = "012"
        elif "SOPRAGEM" in val: sufixo = "013"
        elif "ANDAIME" in val: sufixo = "014"
        elif "OPERADOR" in val: sufixo = "015"
        elif "ESCOPO" in val: sufixo = "016"
        elif "GERENC" in val: sufixo = "101"
        elif "PRODUC" in val or "PRODUÇ" in val: sufixo = "102"
        elif "QUALIDADE" in val: sufixo = "103"
        elif "PLANEJAMENTO" in val: sufixo = "104"
        elif "ADMINISTRA" in val: sufixo = "105"
        elif "MEDICINA" in val or "SEGURAN" in val: sufixo = "106"
        elif "INFRAESTRUTURA" in val: sufixo = "107"
        elif "ALMOXARIFADO" in val and "ENESA" in val: sufixo = "108"
        elif "ALMOXARIFADO" in val: sufixo = "109"
        elif "ELETRICA PROV" in val: sufixo = "110"
        elif "TOPOGRAFIA" in val: sufixo = "111"
        elif "CARGA" in val or "MOVIMENTA" in val: sufixo = "112"
        elif "MEDICAO" in val or "CUSTO" in val: sufixo = "113"
        
        if prefixo and sufixo:
            return f"{prefixo}.{sufixo}"
        return valor

    df["C.C"] = df["C.C"].apply(normalizar_cc)
    
    # -------------------------------------------------------------------------
    # NOVO: FORÇAR A DISCIPLINA A SER SEMPRE CORRETA DE ACORDO COM O C.C ATUAL
    # -------------------------------------------------------------------------
    mapa_sufixo_disciplina = {
        '001': 'EQUIPAMENTOS', '002': 'DUTOS', '003': 'TUBULACAO', 
        '004': 'ESTRUTURA METALICA', '005': 'PRECIPITADOR', '006': 'PRESSAO - MECANICA', 
        '007': 'PRESSAO - TUBULACAO', '008': 'PRESSAO - FORNALHA', '009': 'PINTURA', 
        '010': 'COMISSIONAMENTO', '011': 'OP. ASSISTIDA', '012': 'LAVAGEM QUIMICA', 
        '013': 'SOPRAGEM', '014': 'ANDAIME', '015': 'OPERADORES', '016': 'FORA DE ESCOPO',
        '101': 'GERENCIA', '102': 'PRODUCAO', '103': 'GARANTIA DA QUALIDADE',
        '104': 'PLANEJAMENTO', '105': 'ADMINISTRACAO', '106': 'SEGURANCA E MEDICINA DO TRABALHO',
        '107': 'INFRAESTRUTURA', '108': 'ALMOXARIFADO ENESA', '109': 'ALMOXARIFADO MATERIAIS',
        '110': 'MANUT. ELETRICA PROVISORIA', '111': 'TOPOGRAFIA', '112': 'MOVIMENTACAO DE CARGAS',
        '113': 'MEDICAO/CUSTO/CONTRATOS'
    }

    def corrigir_disciplina(row):
        cc_val = str(row["C.C"]).strip()
        sufixo = cc_val.split('.')[-1] if '.' in cc_val else cc_val
        if sufixo in mapa_sufixo_disciplina:
            return mapa_sufixo_disciplina[sufixo]
        return str(row.get("DISCIPLINA", "")).upper()

    df["DISCIPLINA"] = df.apply(corrigir_disciplina, axis=1)
    # -------------------------------------------------------------------------
    colunas_ordenadas = ["MATRICULA", "NOME", "FUNÇÃO", "C.C", "ENCARREGADO"]
    outras_cols = [c for c in df.columns if c not in colunas_ordenadas]
    df = df[colunas_ordenadas + outras_cols]
    df = df[df["NOME"].str.strip() != ""]
    
    return df

# =================================================================
# INTEGRAÇÃO GOOGLE DRIVE (BACKUP NUVEM)
# =================================================================
def backup_google_drive(file_path, mime_type, file_name):
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        
        # Tentar pegar as credenciais que já usamos pro Sheets
        if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            creds_info = st.secrets["connections"]["gsheets"]
            # Precisamos do escopo do Drive
            scopes = ['https://www.googleapis.com/auth/drive.file']
            creds = service_account.Credentials.from_service_account_info(creds_info, scopes=scopes)
            
            drive_service = build('drive', 'v3', credentials=creds)
            
            # Buscar pasta "RDO_Backups"
            pasta_nome = "RDO_Backups"
            query = f"name='{pasta_nome}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
            response = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
            pastas = response.get('files', [])
            
            if not pastas:
                # Criar pasta se não existe
                folder_metadata = {
                    'name': pasta_nome,
                    'mimeType': 'application/vnd.google-apps.folder'
                }
                folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
                folder_id = folder.get('id')
            else:
                folder_id = pastas[0].get('id')
                
            file_metadata = {
                'name': file_name,
                'parents': [folder_id]
            }
            media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
            arquivo_salvo = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            return True, f"Backup salvo no Drive com ID: {arquivo_salvo.get('id')}"
        else:
            return False, "Credenciais do Google não encontradas no secrets.toml"
    except ImportError:
        return False, "Biblioteca google-api-python-client não instalada."
    except Exception as e:
        return False, f"Erro no Drive: {e}"

# =================================================================
# SESSION STATE
# =================================================================
if 'df' not in st.session_state:
    st.session_state.df = None
if 'df_ia' not in st.session_state:
    st.session_state.df_ia = pd.DataFrame(columns=['ITEM', 'DATA', 'DISCIPLINA', 'ENCARREGADO', 'TURNO', 'DDS', 'ATIVIDADE', 'PROBLEMAS', 'LOCAL', 'AREA'])
if 'df_historico_f1' not in st.session_state:
    st.session_state.df_historico_f1 = pd.DataFrame(columns=["DATA", "ENCARREGADO"])
if 'mostrar_upload' not in st.session_state:
    st.session_state.mostrar_upload = False

# =================================================================
# SISTEMA DE LOGIN (BLOQUEIO GLOBAL) E COOKIES
# =================================================================
import extra_streamlit_components as stx

cookie_manager = stx.CookieManager()

caminho_usuarios = "usuarios.json"
import json
if not os.path.exists(caminho_usuarios):
    with open(caminho_usuarios, "w", encoding="utf-8") as f:
        json.dump({"admin": {"senha": "123", "nome": "Administrador", "role": "admin"}}, f)

def carregar_usuarios():
    if not os.path.exists(caminho_usuarios): return {}
    with open(caminho_usuarios, "r", encoding="utf-8") as f:
        return json.load(f)

def salvar_usuarios(users):
    with open(caminho_usuarios, "w", encoding="utf-8") as f:
        json.dump(users, f)

if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None
if "role_usuario" not in st.session_state:
    st.session_state.role_usuario = None
if "nome_completo" not in st.session_state:
    st.session_state.nome_completo = None

usuarios_db = carregar_usuarios()

# Tentativa de auto-login via Cookie
cookie_user = cookie_manager.get("rdc_user_session")
if st.session_state.usuario_logado is None and cookie_user and cookie_user in usuarios_db:
    st.session_state.usuario_logado = cookie_user
    st.session_state.role_usuario = usuarios_db[cookie_user].get("role", "user")
    st.session_state.nome_completo = usuarios_db[cookie_user].get("nome", cookie_user)

if st.session_state.usuario_logado is None:
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        # Se houver logo, mostra logo acima da caixa
        if os.path.exists(caminho_logo):
            col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
            with col_l2:
                st.image(caminho_logo, use_container_width=True)
            st.markdown("<br>", unsafe_allow_html=True)
            
        with st.container():
            st.markdown(f"""
            <div style="text-align: center; margin-bottom: 20px;">
                <span class="material-symbols-rounded" style="font-size: 54px; color: {cor_destaque}; text-shadow: 0 0 20px rgba(14, 165, 233, 0.6);">rocket_launch</span>
                <h3 style='color: {cor_texto}; margin-bottom: 0px; font-weight: 700; font-size: 26px; letter-spacing: -0.5px;'>{nome_site}</h3>
                <p style='color: {cor_azul}; margin-bottom: 5px; font-size: 13px; font-weight: 600; letter-spacing: 3px;'>ACESSO RESTRITO</p>
            </div>
            """, unsafe_allow_html=True)
            
            user_input = st.text_input("Usuário (Login):", placeholder="Digite sua credencial")
            pass_input = st.text_input("Senha:", type="password", placeholder="••••••••")
            lembrar_me = st.checkbox("Manter conectado", value=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Entrar no Sistema", type="primary", use_container_width=True):
                user_clean = user_input.strip().upper()
                pass_clean = pass_input.strip()
                
                user_encontrado = None
                for key_db in usuarios_db.keys():
                    if key_db.strip().upper() == user_clean:
                        user_encontrado = key_db
                        break
                
                if user_encontrado and usuarios_db[user_encontrado]["senha"] == pass_clean:
                    st.session_state.usuario_logado = user_encontrado
                    st.session_state.role_usuario = usuarios_db[user_encontrado].get("role", "user")
                    st.session_state.nome_completo = usuarios_db[user_encontrado].get("nome", user_encontrado)
                    
                    if lembrar_me:
                        cookie_manager.set("rdc_user_session", user_encontrado, expires_at=datetime.datetime.now() + datetime.timedelta(days=30))
                        
                    time.sleep(1) # Tempo para o cookie assentar
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos. Verifique espaços em branco ou letras erradas.")
    st.stop() # Bloqueia todo o resto do sistema!


# =================================================================


# =================================================================
# CABEÇALHO GLOBAL (Mostrado apenas se logado)
# =================================================================
st.markdown(f"""
    <div class="enesa-header">
        <span class="material-symbols-rounded" style="font-size: 38px; color: {cor_destaque}; vertical-align: middle; margin-right: 12px; text-shadow: 0 0 15px rgba(14, 165, 233, 0.4);">dashboard</span>
        <h1 style="margin: 0; font-size: 2rem; font-weight: 700; display: inline-block; vertical-align: middle;">Sistema de Gestão RDC & PDE</h1>
        <p style="color: {cor_texto_sub}; font-size: 0.95rem; margin: 6px 0 0 0; letter-spacing: 0.5px;">{nome_site} - Controle Operacional de Efetivo</p>
    </div>
""", unsafe_allow_html=True)



# =================================================================
# BARRA LATERAL
# =================================================================
with st.sidebar:
    if os.path.exists(caminho_logo):
        col1, col2, col3 = st.columns([1.5, 2, 1.5]) 
        with col2:
            st.image(caminho_logo, use_container_width=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
    # === BADGE DINÂMICO POR NÍVEL DE ACESSO ===
    _role = st.session_state.get("role_usuario", st.session_state.get("role", "user"))
    if _role == "admin":
        _avatar_emoji = "👨‍💻"
        _avatar_nome = "ADMINISTRADOR"
        _avatar_sub = "Acesso Total"
        _avatar_cor = "#0ea5e9"
    elif _role == "visualizador":
        _avatar_emoji = "👁️"
        _avatar_nome = "VISUALIZADOR"
        _avatar_sub = "Somente Leitura"
        _avatar_cor = "#8b5cf6"
    else:
        _avatar_emoji = "📋"
        _avatar_nome = "ENCARREGADO"
        _avatar_sub = "Acesso de Campo"
        _avatar_cor = "#10b981"
    
    html_avatar = f"""
    <div style="display: flex; align-items: center; gap: 15px; padding: 15px; background: rgba({14 if _role=='admin' else (139 if _role=='visualizador' else 16)}, {165 if _role=='admin' else (92 if _role=='visualizador' else 185)}, {233 if _role=='admin' else (246 if _role=='visualizador' else 129)}, 0.1); border-radius: 12px; border: 1px solid {_avatar_cor}40; margin-bottom: 25px; box-shadow: 0 0 20px {_avatar_cor}30; transition: transform 0.3s;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
        <div style="width: 50px; height: 50px; border-radius: 50%; background: linear-gradient(135deg, {_avatar_cor}, #8b5cf6); display: flex; justify-content: center; align-items: center; font-size: 24px; color: white; box-shadow: 0 0 15px {_avatar_cor}80;">
            {_avatar_emoji}
        </div>
        <div>
            <div style="font-size: 14px; font-weight: 800; color: #f8fafc; text-shadow: 0 0 10px rgba(255,255,255,0.3); letter-spacing: 0.5px;">{_avatar_nome}</div>
            <div style="font-size: 12px; color: {_avatar_cor}; font-weight: bold; margin-top: 2px; text-shadow: 0 0 5px {_avatar_cor}80;">{_avatar_sub}</div>
        </div>
    </div>
    """
    st.markdown(html_avatar, unsafe_allow_html=True)
    
    st.header("📂 Arquivos Base")
    
    if st.button("➕ Enviar Nova Base (PDE)", use_container_width=True):
        st.session_state.mostrar_upload = not st.session_state.mostrar_upload
        
    arquivo_pde = None
    arquivo_modelo = None
    
    if st.session_state.mostrar_upload:
        st.markdown("<div style='background-color: #22262e; padding: 10px; border-radius: 8px;'>", unsafe_allow_html=True)
        arquivo_pde = st.file_uploader("Base de Efetivo (.csv/.xlsx):", type=["csv", "xlsx"])
        arquivo_modelo = st.file_uploader("📄 Layout MODELO.xlsx:", type=["xlsx"])
        st.markdown("</div>", unsafe_allow_html=True)
        
        if arquivo_modelo is not None:
            if salvar_modelo_no_disco(arquivo_modelo):
                st.success("💾 Modelo salvo!")
    
    st.markdown("---")
    
    base_existe = os.path.exists(caminho_base_salva_csv) or os.path.exists(caminho_base_salva_xlsx)
    if base_existe:
        st.success("✅ Base salva no sistema.")
    else:
        st.info("ℹ️ Nenhuma base salva ainda.")
    
    st.markdown("---")
    
    st.markdown("---")
    st.markdown(f"👤 Bem-vindo(a), **{st.session_state.nome_completo}**")
    
    if st.button("Sair (Logout)", use_container_width=True):
        cookie_manager.delete("rdc_user_session")
        st.session_state.usuario_logado = None
        st.session_state.role_usuario = None
        st.session_state.nome_completo = None
        time.sleep(1)
        st.rerun()
        
    if st.session_state.role_usuario == "admin":
        st.markdown("---")
        st.markdown("#### ⚙️ Painel de Configurações")
        
        if st.toggle("🔑 Ver Usuários e Senhas"):
            usuarios_carregados = carregar_usuarios()
            dados_usuarios = []
            for u_nome, u_dados in sorted(usuarios_carregados.items()):
                dados_usuarios.append({
                    "Login": u_nome,
                    "Senha": u_dados.get("senha", ""),
                    "Acesso": u_dados.get("role", "user")
                })
            df_usuarios = pd.DataFrame(dados_usuarios)
            st.dataframe(df_usuarios, hide_index=True, use_container_width=True)
            
        novo_logo = st.file_uploader("Trocar Logo (PNG/JPG):", type=["png", "jpg", "jpeg"])
        if novo_logo:
            with open(caminho_logo, "wb") as f:
                f.write(novo_logo.getbuffer())
            st.success("Logo atualizado! Recarregue a página.")
            
        novo_nome_site = st.text_input("Nome da Empresa/Site:", value=nome_site)
        if st.button("Salvar Nome"):
            with open(caminho_nome_site, "w", encoding="utf-8") as f:
                f.write(novo_nome_site)
            st.success("Nome atualizado!")
            time.sleep(1)
            st.rerun()
                
        st.markdown("---")
        st.markdown("**💾 Backup Seguro**")
        
        # Função para gerar backup ZIP
        buffer_zip = io.BytesIO()
        with zipfile.ZipFile(buffer_zip, "w") as z:
            # Backup da Base de Efetivo
            if st.session_state.df is not None:
                buffer_pde = io.BytesIO()
                st.session_state.df.to_excel(buffer_pde, index=False, engine='openpyxl')
                z.writestr("BASE_EFETIVO.xlsx", buffer_pde.getvalue())
            
            # Backup do Histórico F1
            if "df_historico_f1" in st.session_state and not st.session_state.df_historico_f1.empty:
                buffer_f1 = io.BytesIO()
                st.session_state.df_historico_f1.to_excel(buffer_f1, index=False, engine='openpyxl')
                z.writestr("HISTORICO_F1.xlsx", buffer_f1.getvalue())
        
        st.download_button(
            label="📥 Baixar Backup (.zip)",
            data=buffer_zip.getvalue(),
            file_name=f"Backup_RDC_{datetime.datetime.now().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            use_container_width=True
        )



        st.markdown("---")
        st.markdown("#### 👥 Gestão de Usuários")
        with st.form("form_novo_usuario"):
            st.markdown("**Adicionar / Editar Usuário**")
            novo_user = st.text_input("Usuário (Login):")
            nova_senha = st.text_input("Senha:")
            novo_nome = st.text_input("Nome Completo:")
            nova_role = st.selectbox("Nível de Acesso:", ["encarregado", "visualizador", "admin"], help="Admin (Tudo), Visualizador (Dashboard), Encarregado (Apenas preenche RDC)")
            submit_user = st.form_submit_button("Salvar Usuário")
            if submit_user and novo_user and nova_senha:
                usuarios_db[novo_user] = {"senha": nova_senha, "nome": novo_nome, "role": nova_role}
                salvar_usuarios(usuarios_db)
                st.success(f"Usuário '{novo_user}' salvo!")
                time.sleep(1)
                st.rerun()
        
        st.markdown("**Usuários Cadastrados:**")
        for u, dados in sorted(usuarios_db.items()):
            role_atual = dados.get('role', 'encarregado')
            if role_atual == 'user': role_atual = 'encarregado'
            
            if u == "admin":
                st.markdown(f"<div style='margin-top: 5px; margin-bottom: 5px; font-size: 14px;'>👤 <b>{u}</b> <span style='color:#888; font-size:12px;'>(admin)</span></div>", unsafe_allow_html=True)
            else:
                col_u, col_role, col_del = st.columns([4, 3, 1])
                col_u.markdown(f"<div style='margin-top: 5px; font-size: 14px;'>👤 <b>{u}</b></div>", unsafe_allow_html=True)
                
                novo_role = col_role.selectbox(
                    "Acesso", 
                    ["encarregado", "visualizador", "admin"], 
                    index=["encarregado", "visualizador", "admin"].index(role_atual) if role_atual in ["encarregado", "visualizador", "admin"] else 0,
                    key=f"role_sel_{u}",
                    label_visibility="collapsed"
                )
                
                if novo_role != role_atual:
                    usuarios_db[u]["role"] = novo_role
                    salvar_usuarios(usuarios_db)
                    st.rerun()
                    
                if col_del.button("❌", key=f"del_{u}"):
                    del usuarios_db[u]
                    salvar_usuarios(usuarios_db)
                    st.rerun()

    st.markdown("---")
    # === BOTÃO DE LOGOUT ===
    st.divider()
    if st.button("🚪 Sair do Sistema", use_container_width=True, type="secondary", key="btn_logout_sidebar"):
        try:
            cookie_manager.delete("rdc_user_session")
        except Exception:
            pass
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
    
    st.markdown(
        f"""
        <div style='text-align: center; margin-top: 20px; font-size: 12px; color: #888;'>
            <p>Desenvolvido por</p>
            <p style='font-size: 16px; font-weight: bold; color: #ff4b4b; margin-top: -10px;'>Edson Garcia - 125</p>
            <p style='margin-top: -10px;'>v5.0 · {nome_site} (com ANTIGRAVITY)</p>
        </div>
        """, 
        unsafe_allow_html=True
    )

# =================================================================
# LÓGICA DE CARREGAMENTO DA NUVEM (GOOGLE SHEETS)
# =================================================================
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    conn = None

if arquivo_pde is not None:
    df_carregado = ler_arquivo_seguro(arquivo_pde, arquivo_pde.name)
    if df_carregado is not None:
        # PREPARAR PRIMEIRO para padronizar nomes de colunas (MATRÍCULA -> MATRICULA, etc.)
        df_carregado = preparar_dataframe(df_carregado)
        
        # PRESERVAR O C.C DA BASE ANTIGA DO APP
        if st.session_state.df is not None:
            # Prepara a base antiga para garantir que o nome da coluna é exatamente "C.C" e "MATRICULA"
            df_antigo = preparar_dataframe(st.session_state.df)
            
            if "C.C" in df_antigo.columns and "MATRICULA" in df_carregado.columns:
                try:
                    df_com_cc = df_antigo[df_antigo["C.C"].astype(str).str.strip() != ""]
                    
                    # O C.C do PDE agora é a verdade absoluta.
                    # Apenas herdamos o C.C do Encarregado caso o colaborador venha sem C.C no PDE.
                    # Passo 2: Herdar o C.C do Encarregado para novos colaboradores (ou transferidos sem C.C)
                    if "ENCARREGADO" in df_carregado.columns and "ENCARREGADO" in df_com_cc.columns:
                        mapa_enc_cc = df_com_cc.groupby("ENCARREGADO")["C.C"].agg(lambda x: x.value_counts().index[0] if len(x) > 0 else "").to_dict()
                        
                        def herdar_cc(row):
                            cc_atual = str(row.get("C.C", "")).strip()
                            if not cc_atual:
                                return mapa_enc_cc.get(row.get("ENCARREGADO", ""), "")
                            return cc_atual
                            
                        df_carregado["C.C"] = df_carregado.apply(herdar_cc, axis=1)
                        
                except Exception:
                    pass
        
        st.session_state.df = df_carregado
        if conn and not st.session_state.get('force_use_local', False):
            try:
                conn.update(worksheet="Página1", data=st.session_state.df)
                st.sidebar.success("☁️ Base salva! C.Cs antigos foram preservados com sucesso!")
            except Exception as e:
                st.sidebar.error(f"Erro Nuvem: {e}")

elif st.session_state.df is None:
    carregado_nuvem = False
    if conn and not st.session_state.get('force_use_local', False):
        try:
            df_gsheets = conn.read(worksheet="Página1", ttl=5)
            df_gsheets = df_gsheets.dropna(how='all')
            if not df_gsheets.empty:
                st.session_state.df = df_gsheets
                carregado_nuvem = True
        except Exception:
            pass
            
        try:
            df_f1 = conn.read(worksheet="Historico_F1", ttl=5)
            if not df_f1.empty:
                st.session_state.df_historico_f1 = df_f1.dropna(how='all')
        except Exception:
            pass
            
    # Resetar a flag
    if st.session_state.get('force_use_local', False):
        carregado_nuvem = True
        st.session_state.force_use_local = False
            
    if not carregado_nuvem:
        if os.path.exists(caminho_base_salva_xlsx):
            df_carregado = ler_arquivo_seguro(caminho_base_salva_xlsx, "BASE_ATUAL.xlsx")
            if df_carregado is not None:
                st.session_state.df = df_carregado
        elif os.path.exists(caminho_base_salva_csv):
            df_carregado = ler_arquivo_seguro(caminho_base_salva_csv, "BASE_ATUAL.csv")
            if df_carregado is not None:
                st.session_state.df = df_carregado
        elif os.path.exists(caminho_pde_padrao):
            df_carregado = ler_arquivo_seguro(caminho_pde_padrao, "PDE.csv")
            if df_carregado is not None:
                st.session_state.df = df_carregado

# =================================================================
# CONTEUDO PRINCIPAL
# =================================================================
if st.session_state.df is not None:
    df_atual = preparar_dataframe(st.session_state.df.copy())
    lista_encarregados_base = sorted([str(e) for e in df_atual["ENCARREGADO"].unique() if str(e).strip() != ""])

    # ====== HISTÓRICO DE C.C (FOTOGRAFIA DIÁRIA) ======
    try:
        hoje_hist = datetime.date.today().strftime("%Y-%m-%d")
        df_cc_valido = df_atual[df_atual["C.C"].astype(str).str.strip() != ""]
        if not df_cc_valido.empty:
            cc_counts = df_cc_valido["C.C"].value_counts().reset_index()
            cc_counts.columns = ["C.C", "Efetivo"]
            cc_counts["DATA"] = hoje_hist
            
            hist_cc_existente = pd.DataFrame()
            if os.path.exists(caminho_hist_cc):
                try:
                    hist_cc_existente = pd.read_csv(caminho_hist_cc)
                except:
                    pass
                
            if not hist_cc_existente.empty and "DATA" in hist_cc_existente.columns:
                hist_cc_existente = hist_cc_existente[hist_cc_existente["DATA"] != hoje_hist]
                hist_cc_novo = pd.concat([hist_cc_existente, cc_counts], ignore_index=True)
            else:
                hist_cc_novo = cc_counts
                
            hist_cc_novo.to_csv(caminho_hist_cc, index=False)
    except Exception as e:
        pass # Falha silenciosa para não quebrar o app
    # ==================================================

    # === LISTA DINÂMICA DE ENCARREGADOS (Carrega do JSON ou cria com a lista padrão) ===
    import json
    caminho_f1_json = os.path.join(os.path.dirname(__file__), "encarregados_f1.json")
    caminho_f1_excecoes = os.path.join(os.path.dirname(__file__), "f1_excecoes.csv")
    
    encarregados_f1_padrao = [
        "ABMAEL PEREIRA PAIVA", "JEAN PEDRO", "ANANIAS DE SOUSA NETO", "GILDO GONCALVES DA SILVA",
        "SIDNEI FERNANDES DA SILVA", "BARTOLOMEU FERNANDES", "FRANCINALDO DE SOUSA", "IZAIAS BAIA BELO",
        "SANDRO LIMA DE SOUZA", "ALOISIO FERREIRA SOUZA", "ARLINDO PEREIRA DA SILVA", "FAUZE CELIS RODRIGUES COSTA",
        "FRANCISCO PEREIRA LIMA", "JOAO PAULO DA COSTA QUARESMA", "JOSE ORLANDO DAS NEVES MADEIRA",
        "JOSE TARCISIO ARAUJO DA SILVA", "LEANDRO DA CRUZ DE SOUZA", "CLAUDIO LUCIANO ARGELINO",
        "EDVALDO CARVALHO ANGELIM", "ELDER MENDES JUNIOR", "MANOEL MARIA SARGES SOARES", "CLAUDIO CRUZ SOUSA",
        "CLIDENILDO GOMES DE ALMEIDA", "GRACINEI PEREIRA DOS SANTOS", "JAILSON MENDES DE OLIVEIRA",
        "JARBAS DA ROCHA GOMES", "JOSE MAURICIO RODRIGUES DA SILVA", "JOSE SARAIVA LOPES NETO",
        "JOSMAEL RODRIGUES PEREIRA", "ALEX PANTOJA DE OLIVEIRA", "ARILSON DIAS DO PRADO", "ELTON GOMES DOS SANTOS",
        "RICARDO SARMENTO FERREIRA", "WENISON DA SILVA CUNHA CORREIA", "FRANCISCO ALVES DA PENHA",
        "IVAN DO NASCIMENTO RAMOS", "ELDER MENDES", "GEAN LENO JOSE DE FREITAS", "JOSE EDUARDO FARIAS FERREIRA",
        "EDIMILSON NUNES VASCONCELOS", "LOURISVALDO AMARAL ARAUJO", "VALDEMIR BARBOSA REIS",
        "LUZINALDO AMARAL DE ARAUJO", "MAURO DE QUEIROZ ANDRADE", "ELIAS SOUSA DA COSTA", "ISAIAS SOUSA LISBOA",
        "ISMAEL CARLOS GOMES DA SILVA", "RAIMUNDO DA SILVA DOS SANTOS", "RAIMUNDO EUDE DA SILVA FREITAS",
        "RODOLFO DOS SANTOS COSTA", "ELISEU DA SILVA BISPO", "IRON MARQUES MOREIRA", "LUIZ CARLOS DE SOUZA",
        "ANTONIO TEIXEIRA BORBA", "JOSE FRANCIVAN MONTEIRO SANTOS", "JOSE WALKER CARNEIRO OLIVEIRA",
        "LEANDRO DA SILVA QUEIROZ", "SILVIO MANOEL DE ANDRADE", "EVERALDO DOS SANTOS SOARES",
        "FRANCISCO GRACIEL DE SOUSA MARTINS", "JAILSON SILVA DE GOIS", "JORGINALDO NUNES DA SILVA",
        "CLAUDIVAN OLIVEIRA DOS SANTOS", "GUILHERME HENRIQUE DE ARAUJO SOUSA", "LEANDRO MARTINS DA SILVA BORGES",
        "WEVERTON FERNANDES MARIANO", "JORGE DA COSTA SILVA", "RAIMUNDO FRAZAO DOS SANTOS",
        "JOSE RIBEIRO DO NASCIMENTO JUNIOR", "JOSE ROBERTO SALVADOR FILHO", "MARCUS ANTONIO DE SOUZA",
        "RAIMUNDO ROGERIO LEITE", "ROUBERVAL SANTOS DOS SANTOS", "CARLOS ALBERTO DA COSTA MOREIRA",
        "JOSE FELIPE DOS SANTOS", "JOSE GERIARDI FONSECA DE SENA", "JOSE HENRIQUE SILVA VIEIRA",
        "ODAIR MENEZES DA SILVA", "SIDNALDO SANTOS DE JESUS", "ANDERSON VICTALINO",
        "FRANCISCO AUGUSTO DE SOUSA BARROS", "GENILSON PEREIRA DE SOUSA", "HELENO MARQUES DE SOUZA NETO",
        "HEMERSON MONTEIRO DE OLIVEIRA", "JACKSON DEIBSON FELICIANO DA SILVA", "JARDELINO PEREIRA DA COSTA",
        "JOAO TIAGO OLIVEIRA DE AMORIM", "JOSE MARIA DA SILVA PESSOA", "LUCIO FABIO DA SILVA LEANDRO",
        "RAIMUNDO GONCALVES DOS SANTOS", "FABRICIO FIGUEIREDO", "RHOKSONY FERREIRA SILVEIRA",
        "FERNANDO DA CONCEIÇÃO", "ROGERIO BARROS DOS SANTOS", "SIQUEU SANTOS SOLEDADE",
        "SEBASTIAO CARLOS DE OLIVEIRA", "MANOEL NEPOMUCENO DOS SANTOS", "LUIZ RAMOS DE LIMA",
        "JORGE LUIS LOPES", "VALDINEI GOMES OLIVEIRA", "CARLOS DA SILVA OLIVEIRA"
    ]
    
    # Carregar ou criar o JSON
    if os.path.exists(caminho_f1_json):
        try:
            with open(caminho_f1_json, "r", encoding="utf-8") as f:
                encarregados_f1_oficial = json.load(f)
        except Exception:
            encarregados_f1_oficial = encarregados_f1_padrao
    else:
        encarregados_f1_oficial = encarregados_f1_padrao
        with open(caminho_f1_json, "w", encoding="utf-8") as f:
            json.dump(encarregados_f1_padrao, f, ensure_ascii=False, indent=2)
    
    lista_completa_encarregados = sorted([e.upper() for e in encarregados_f1_oficial])
    
    # Carregar exceções (Abonos)
    if "df_f1_excecoes" not in st.session_state:
        if os.path.exists(caminho_f1_excecoes):
            try:
                st.session_state.df_f1_excecoes = pd.read_csv(caminho_f1_excecoes)
            except Exception:
                st.session_state.df_f1_excecoes = pd.DataFrame(columns=["DATA", "ENCARREGADO", "MOTIVO"])
        else:
            st.session_state.df_f1_excecoes = pd.DataFrame(columns=["DATA", "ENCARREGADO", "MOTIVO"])

    def obter_saudacao():
        import datetime
        hora = datetime.datetime.now().hour
        if 5 <= hora < 12:
            return "Bom dia"
        elif 12 <= hora < 18:
            return "Boa tarde"
        else:
            return "Boa noite"

    def exibir_apresentacao(nome):
        if st.session_state.get("saudacao_vista", False):
            return
            
        st.session_state.saudacao_vista = True
        
        saudacao = obter_saudacao()
        primeiro_nome = nome.split()[0].title() if nome else "Equipe"
        
        st.markdown(f"""
        <style>
        @keyframes autoHideBanner {{
            0% {{ opacity: 1; max-height: 200px; padding: 20px; margin-bottom: 25px; border-left-width: 5px; }}
            70% {{ opacity: 1; max-height: 200px; padding: 20px; margin-bottom: 25px; border-left-width: 5px; }}
            90% {{ opacity: 0; max-height: 200px; padding: 20px; margin-bottom: 25px; border-left-width: 5px; }}
            100% {{ opacity: 0; max-height: 0px; padding: 0px; margin-bottom: 0px; border-left-width: 0px; overflow: hidden; }}
        }}
        .greeting-box {{
            background: linear-gradient(135deg, rgba(14,165,233,0.1), rgba(139,92,246,0.1)); 
            border-radius: 12px; 
            border-left: 5px solid #0ea5e9; 
            animation: autoHideBanner 5s forwards;
            overflow: hidden;
        }}
        </style>
        <div class='greeting-box'>
            <h2 style='margin-top: 0; margin-bottom: 5px; color: #f8fafc;'>{saudacao}, {primeiro_nome}! 👋</h2>
            <p style='margin: 0; color: #94a3b8; font-size: 16px;'>Em que posso ajudar hoje?</p>
        </div>
        """, unsafe_allow_html=True)

    # =================================================================
    # MODO ENCARREGADO (Lançamento Nativo com Formatação Original)
    # =================================================================
    if st.session_state.get("role") == "encarregado":
        exibir_apresentacao(st.session_state.get("nome_completo", "Encarregado"))
        st.markdown("### <span class='material-symbols-rounded' style='vertical-align: middle; color: #0ea5e9; font-size: 32px;'>edit_document</span> Lançamento de RDC Digital", unsafe_allow_html=True)
        st.caption("Preencha as informações do seu dia de trabalho seguindo as 3 etapas abaixo. Os dados serão salvos na nuvem.")
        
        with st.form("form_rdc_digital_encarregado"):
            tab_id, tab_local, tab_ativ = st.tabs(["1️⃣ Identificação", "2️⃣ Localização", "3️⃣ Atividades e Envio"])
            
            with tab_id:
                st.markdown("<p style='color: #94a3b8; font-size: 14px;'>Quem é você e qual seu turno?</p>", unsafe_allow_html=True)
                rdc_encarregado = st.selectbox("Selecione seu Nome (Encarregado):", [""] + lista_completa_encarregados)
                rdc_turno = st.selectbox("Turno de Trabalho:", ["DIURNO", "NOTURNO", "MISTO"])
                
            with tab_local:
                import datetime
                st.markdown("<p style='color: #94a3b8; font-size: 14px;'>Onde você trabalhou hoje?</p>", unsafe_allow_html=True)
                
                rdc_data = st.date_input("Data do Relatório:", datetime.date.today())
                
                area_options = ["PB", "RB", "ESP", "LAYDOWN 1", "LAYDOWN 2", "OUTRO (DIGITAR)"]
                area_sel = st.selectbox("Área / Local de Trabalho:", area_options)
                rdc_area = area_sel
                if area_sel == "OUTRO (DIGITAR)":
                    rdc_area = st.text_input("Qual Área/Local?", placeholder="Ex: Escritório, Almoxarifado...")
                
                disc_options = [
                    "EQUIPAMENTOS", "DUTOS", "TUBULACAO", "ESTRUTURA METALICA", "PRECIPITADOR", 
                    "PRESSAO - MECANICA", "PRESSAO - TUBULACAO", "PRESSAO - FORNALHA", "PINTURA", 
                    "COMISSIONAMENTO", "OP. ASSISTIDA", "LAVAGEM QUIMICA", "SOPRAGEM", "ANDAIME", 
                    "OPERADORES", "FORA DE ESCOPO", "GERENCIA", "PRODUCAO", "GARANTIA DA QUALIDADE", 
                    "PLANEJAMENTO", "ADMINISTRACAO", "SEGURANCA E MEDICINA DO TRABALHO", "INFRAESTRUTURA", 
                    "ALMOXARIFADO ENESA", "ALMOXARIFADO MATERIAIS", "MANUT. ELETRICA PROVISORIA", 
                    "TOPOGRAFIA", "MOVIMENTACAO DE CARGAS", "MEDICAO/CUSTO/CONTRATOS", "CIVIL", "MECÂNICA", "ELÉTRICA", "INSTRUMENTAÇÃO", "ISOLAMENTO", "OUTRA (DIGITAR)"
                ]
                disc_sel = st.selectbox("Disciplina Principal:", disc_options)
                
                rdc_disciplina = disc_sel
                if disc_sel == "OUTRA (DIGITAR)":
                    rdc_disciplina = st.text_input("Qual Disciplina?", placeholder="Ex: Tubulação, Solda...")
                    
            with tab_ativ:
                st.markdown("<p style='color: #94a3b8; font-size: 14px;'>O que foi executado?</p>", unsafe_allow_html=True)
                rdc_dds = st.text_input("Tópico do DDS do dia:")
                rdc_atividades = st.text_area("Atividades Executadas (Detalhe os serviços feitos pela equipe):", height=150)
                rdc_problemas = st.text_area("Problemas / Interrupções / Ocorrências (Opcional):", height=68)
                
                st.markdown("<br>", unsafe_allow_html=True)
                submit_rdc = st.form_submit_button("🚀 Salvar e Enviar RDC na Nuvem", use_container_width=True, type="primary")
            
            if submit_rdc:
                if not rdc_encarregado:
                    st.error("⚠️ Por favor, selecione o nome do Encarregado.")
                elif not rdc_atividades.strip():
                    st.error("⚠️ Por favor, preencha as Atividades Executadas.")
                elif disc_sel == "OUTRA (DIGITAR)" and not rdc_disciplina.strip():
                    st.error("⚠️ Digite a disciplina na caixa 'Qual Disciplina?'.")
                else:
                    rdc_json = [{
                        "ENCARREGADO": rdc_encarregado,
                        "DATA": rdc_data.strftime("%Y/%m/%d"),
                        "TURNO": rdc_turno,
                        "AREA": rdc_area.strip().upper(),
                        "DISCIPLINA": rdc_disciplina.strip().upper(),
                        "DDS": rdc_dds.strip(),
                        "ATIVIDADE": rdc_atividades.strip(),
                        "CALDEIRA": rdc_problemas.strip(),
                        "PROBLEMAS": rdc_problemas.strip()
                    }]
                    
                    import json
                    import requests
                    
                    WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxfE96gE7ckdmapBLBHJuoX2bvAt-2d76OUJNiSRsLgFCOiySeQhFOopp3DoC5Fn95D/exec"
                    
                    try:
                        with st.spinner("Enviando dados para a nuvem..."):
                            res = requests.post(WEBHOOK_URL, json=rdc_json, allow_redirects=True)
                        if res.status_code == 200:
                            st.success(f"✅ RDC Digital de {rdc_encarregado} salvo com sucesso na Nuvem!")
                            st.info("O relatório já foi enviado. Você pode fechar esta página.")
                        else:
                            st.error(f"❌ Erro ao enviar. Servidor retornou: {res.text}")
                    except Exception as e:
                        st.error(f"❌ Falha de conexão: {e}")
                        
        st.stop() # Finaliza o script para não mostrar as abas do admin

    mapa_area_sufixo = {
        'EQUIPAMENTO': '001', 'EQUIPAMENTOS': '001',
        'DUTO': '002', 'DUTOS': '002',
        'TUBULACAO': '003', 'TUBULAÇÃO': '003',
        'ESTRUTURA MET': '004', 'ESTRUTURA METALICA': '004', 'ESTRUTURA METÁLICA': '004',
        'PRECIPITADOR': '005', 'ESP': '005',
        'PRESSAO-MEC': '006', 'PRESSAO - MEC': '006', 'PARTE DE PRESSAO - MECANICA': '006',
        'PRESSAO-TUBULACAO': '007', 'PRESSAO - TUBULACAO': '007', 'PARTE DE PRESSAO - TUBULACAO': '007',
        'PRESSAO-FORNALHA': '008', 'PRESSAO - FORNALHA': '008', 'PARTE DE PRESSAO - FORNALHA': '008', 'PARTE DE PRESSAO - FORNALIA': '008',
        'PINTURA': '009',
        'COMISSIONAMENTO': '010', 'APOIO AO COMISSIONAMENTO': '010',
        'OPERACAO ASSISTIDA': '011', 'OPERAÇÃO ASSISTIDA': '011',
        'LAVAGEM QUIMICA': '012', 'LAVAGEM QUÍMICA': '012',
        'SOPRAGEM': '013',
        'ANDAIME': '014', 'ANDAIMES': '014',
        'OPERADOR': '015', 'OPERADORES E MOTORISTAS': '015', 'MOTORISTA': '015',
        'FORA DE ESCOPO': '016', 'SERVICOS FORA DE ESCOPO': '016', 'SERVIÇOS FORA DE ESCOPO': '016'
    }

    # === CONTROLE DE ABAS POR NÍVEL DE ACESSO ===
    user_role = st.session_state.get("role_usuario", st.session_state.get("role", "user"))
    
    if user_role == "admin":
        tab_dashboard, tab_resumo, tab_emissao, tab_cc, tab_f1, tab_ia, tab_ia_cc, tab_rdc_digital = st.tabs(["📊 Dashboard", "📅 Resumo Diário", "📝 Emissão de RDC", "💰 Controle de C.C", "🏎️ Competição F1", "🤖 Leitor de RDC (IA)", "🤖 IA - Atualizador de C.C", "📱 RDC Digital"])
    elif user_role == "visualizador":
        tab_dashboard, tab_cc, tab_f1 = st.tabs(["📊 Dashboard", "💰 Controle de C.C", "🏎️ Competição F1"])
        # Criar variáveis dummy para evitar erros em blocos condicionais
        tab_resumo = tab_emissao = tab_ia = tab_ia_cc = tab_rdc_digital = None
    else:
        tab_dashboard, tab_resumo, tab_emissao, tab_cc, tab_f1, tab_ia, tab_ia_cc, tab_rdc_digital = st.tabs(["📊 Dashboard", "📅 Resumo Diário", "📝 Emissão de RDC", "💰 Controle de C.C", "🏎️ Competição F1", "🤖 Leitor de RDC (IA)", "🤖 IA - Atualizador de C.C", "📱 RDC Digital"])

    with tab_dashboard:
        exibir_apresentacao(st.session_state.get("nome_completo", "Usuário"))
        
        # === RELÓGIO DIGITAL ===
        import streamlit.components.v1 as components
        html_relogio = """
        <div id="clock_container" style="font-family: 'Courier New', Courier, monospace; font-size: 28px; color: #0ea5e9; font-weight: bold; text-shadow: 0 0 10px rgba(14, 165, 233, 0.8); text-align: center; background: linear-gradient(145deg, rgba(15, 23, 42, 0.9), rgba(30, 41, 59, 0.9)); padding: 10px 20px; border-radius: 12px; border: 1px solid rgba(14, 165, 233, 0.4); width: fit-content; margin: 0 auto 20px auto; box-shadow: 0 4px 15px rgba(0,0,0,0.5), inset 0 0 10px rgba(14,165,233,0.1);">
            <div id="time" style="letter-spacing: 2px;">--:--:--</div>
            <div id="date" style="font-size: 14px; color: #94a3b8; font-weight: normal; text-shadow: none; font-family: 'Inter', sans-serif; margin-top: 5px; text-transform: uppercase; letter-spacing: 1px;">Carregando...</div>
        </div>
        <script>
            function updateClock() {
                const now = new Date();
                const timeStr = now.toLocaleTimeString('pt-BR');
                const dateOptions = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
                let dateStr = now.toLocaleDateString('pt-BR', dateOptions);
                document.getElementById('time').innerText = timeStr;
                document.getElementById('date').innerText = dateStr;
            }
            setInterval(updateClock, 1000);
            updateClock();
        </script>
        """
        components.html(html_relogio, height=110)
        
        # === MÓDULO CLIMA-OBRA ===
        st.markdown("### 🌦️ Clima e Impacto Operacional")
        col_cidade, col_clima = st.columns([1, 2])
        
        if "cidade_obra" not in st.session_state:
            st.session_state.cidade_obra = "Lençóis Paulista"
            
        with col_cidade:
            nova_cidade = st.text_input("📍 Local da Obra (Cidade):", value=st.session_state.cidade_obra)
            if nova_cidade != st.session_state.cidade_obra:
                st.session_state.cidade_obra = nova_cidade
                st.rerun()
                
        with col_clima:
            try:
                import urllib.request, json
                from urllib.parse import quote
                # Pegar Coordenadas
                city_encoded = quote(st.session_state.cidade_obra)
                url_geo = f"https://geocoding-api.open-meteo.com/v1/search?name={city_encoded}&count=1&language=pt&format=json"
                req_geo = urllib.request.Request(url_geo, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req_geo, timeout=5) as response:
                    geo_data = json.loads(response.read().decode())
                
                if geo_data.get("results"):
                    lat = geo_data["results"][0]["latitude"]
                    lon = geo_data["results"][0]["longitude"]
                    cidade_nome = geo_data["results"][0]["name"]
                    
                    # Pegar Clima
                    url_weather = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&current=temperature_2m,weather_code&hourly=precipitation_probability,temperature_2m&timezone=America%2FSao_Paulo"
                    req_w = urllib.request.Request(url_weather, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req_w, timeout=5) as response:
                        weather_data = json.loads(response.read().decode())
                        
                    temp_atual = weather_data["current"]["temperature_2m"]
                    code = weather_data["current"]["weather_code"]
                    
                    # Interpretar Weather Code (WMO)
                    icon = "☀️"
                    desc = "Ensolarado"
                    if code in [1, 2, 3]: icon, desc = "⛅", "Parcialmente Nublado"
                    elif code in [45, 48]: icon, desc = "🌫️", "Neblina"
                    elif code in [51, 53, 55, 61, 63, 65, 80, 81, 82]: icon, desc = "🌧️", "Chuva"
                    elif code in [71, 73, 75, 77, 85, 86]: icon, desc = "❄️", "Frio extremo"
                    elif code in [95, 96, 99]: icon, desc = "⛈️", "Tempestade"
                    
                    # Verificar risco de chuva nas próximas 12 horas
                    import datetime
                    agora = datetime.datetime.now().hour
                    chuva_probs = weather_data["hourly"]["precipitation_probability"][agora:agora+12]
                    risco_chuva = max(chuva_probs) if chuva_probs else 0
                    
                    # Layout
                    st.markdown(f"""
                    <div style="background: linear-gradient(90deg, #1e293b, #0f172a); padding: 15px; border-radius: 12px; border-left: 5px solid #0ea5e9; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 6px rgba(0,0,0,0.2);">
                        <div>
                            <div style="font-size: 14px; color: #94a3b8;">{cidade_nome}</div>
                            <div style="font-size: 28px; font-weight: bold; color: white;">{icon} {temp_atual}°C</div>
                            <div style="font-size: 14px; color: #cbd5e1;">{desc}</div>
                        </div>
                        <div style="text-align: right;">
                            <div style="font-size: 12px; color: #94a3b8; margin-bottom: 5px;">PROBABILIDADE DE CHUVA (12h)</div>
                            <div style="font-size: 22px; font-weight: bold; color: {'#ef4444' if risco_chuva > 50 else '#f59e0b' if risco_chuva > 20 else '#10b981'};">{risco_chuva}%</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if risco_chuva > 50:
                        st.error(f"⚠️ **ALERTA CLIMÁTICO:** Alta probabilidade de chuva ({risco_chuva}%) nas próximas horas. Considere replanejar frentes de trabalho externas (Caldeiras / Pátio).")
                    elif temp_atual > 34:
                        st.warning(f"⚠️ **ALERTA CALOR:** Temperatura muito elevada ({temp_atual}°C). Reforce no DDS a necessidade de pausas para hidratação das equipes!")
                else:
                    st.info("🌦️ Cidade não encontrada. Tente digitar sem acentos.")
            except Exception as e:
                st.caption(f"⚠️ Não foi possível carregar a API do clima no momento. (Verifique conexão)")
                
        st.divider()
        
        st.markdown("### 🎛️ Centro de Comando (Overview)")
        
        # Filtro de MOI / MOD e Local
        col_filtros1, col_filtros2 = st.columns(2)
        with col_filtros1:
            filtro_dash_mo = st.segmented_control(
                "Filtrar Visão por Tipo de Mão de Obra:", 
                ["Ambas", "MOD", "MOI"], 
                default="Ambas"
            )
            if not filtro_dash_mo:
                filtro_dash_mo = "Ambas"
                
        with col_filtros2:
            filtro_dash_local = st.segmented_control(
                "Filtrar Dados por Local:", 
                ["Ambas", "PB", "RB", "ESP"], 
                default="Ambas",
                key="filtro_dash_local_key"
            )
            if not filtro_dash_local:
                filtro_dash_local = "Ambas"
            
        df_dash = df_atual.copy()
        
        # Aplicar filtro MOI/MOD
        if filtro_dash_mo == "MOD":
            df_dash = df_dash[df_dash["MÃO DE OBRA"].astype(str).str.strip().str.upper() == "MOD"]
        elif filtro_dash_mo == "MOI":
            df_dash = df_dash[df_dash["MÃO DE OBRA"].astype(str).str.strip().str.upper() == "MOI"]
            
        # Aplicar filtro Local
        df_dash = df_dash[df_dash["C.C"].str.strip() != ""]
        if filtro_dash_local == "PB":
            df_dash = df_dash[df_dash["C.C"].apply(lambda x: "125.02" in str(x) and ".005" not in str(x))]
        elif filtro_dash_local == "RB":
            df_dash = df_dash[df_dash["C.C"].apply(lambda x: "125.01" in str(x) and ".005" not in str(x))]
        elif filtro_dash_local == "ESP":
            df_dash = df_dash[df_dash["C.C"].apply(lambda x: ".005" in str(x))]
        
        # Linha 1: Cartões de KPI Customizados (Premium)
        qtd_encarregados_dash = len([e for e in df_dash["ENCARREGADO"].unique() if str(e).strip() != "" and str(e) in lista_completa_encarregados])
        qtd_mod_g = len(df_atual[df_atual["MÃO DE OBRA"].str.strip().str.upper() == "MOD"])
        qtd_moi_g = len(df_atual[df_atual["MÃO DE OBRA"].str.strip().str.upper() == "MOI"])
        total_mo_g = qtd_mod_g + qtd_moi_g
        pct_mod_g = round((qtd_mod_g / total_mo_g * 100), 1) if total_mo_g > 0 else 0
        span_control = round(len(df_dash) / qtd_encarregados_dash, 1) if qtd_encarregados_dash > 0 else 0
        
        def card_kpi(titulo, valor, icone, cor):
            return f"""
            <div style="background: rgba(30, 41, 59, 0.45); backdrop-filter: blur(10px); border-radius: 16px; border: 1px solid rgba(255,255,255,0.05); padding: 18px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); position: relative; overflow: hidden; height: 110px; transition: transform 0.3s ease;" onmouseover="this.style.transform='translateY(-5px)'" onmouseout="this.style.transform='translateY(0px)'">
                <p style="margin: 0; font-size: 13px; color: #94a3b8; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;">{titulo}</p>
                <h2 style="margin: 5px 0 0 0; font-size: 34px; font-weight: 700; color: #f8fafc; text-shadow: 0 0 15px {cor}60;">{valor}</h2>
                <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(90deg, {cor}, transparent); box-shadow: 0 -2px 10px {cor}80;"></div>
            </div>
            """
            
        st.markdown("<br>", unsafe_allow_html=True)
        m1, m2, m3, m4, m5 = st.columns(5)
        with m1: st.markdown(card_kpi(f"Efetivo ({filtro_dash_mo})", len(df_dash), "engineering", "#3b82f6"), unsafe_allow_html=True)
        with m2: st.markdown(card_kpi("Encarregados", qtd_encarregados_dash, "shield_person", "#10b981"), unsafe_allow_html=True)
        with m3: st.markdown(card_kpi("% MOD Global", f"{pct_mod_g}%", "pie_chart", "#0ea5e9"), unsafe_allow_html=True)
        with m4: st.markdown(card_kpi("Funções", df_dash["FUNÇÃO"].nunique(), "build", "#f59e0b"), unsafe_allow_html=True)
        with m5: st.markdown(card_kpi("Span Control", span_control, "groups", "#8b5cf6"), unsafe_allow_html=True)
        
        st.markdown("---")
        
        col_dash1, col_dash2, col_dash3 = st.columns([3, 3, 4])
        
        with col_dash1:
            st.markdown("**Status Operacional (Global)**")
            if total_mo_g > 0:
                df_mo_global = pd.DataFrame({"Tipo": ["MOD", "MOI"], "Quantidade": [qtd_mod_g, qtd_moi_g]})
                fig_mo_g = px.pie(df_mo_global, values="Quantidade", names="Tipo", hole=0.6, color_discrete_sequence=["#10b981", "#ef4444"])
                fig_mo_g.update_layout(margin=dict(l=20, r=20, t=10, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), height=280, showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
                st.plotly_chart(fig_mo_g, use_container_width=True)
            else:
                st.info("Classificação de Mão de Obra não encontrada.")
                
        with col_dash2:
            st.markdown("**Efetivo por Área**")
            df_area = df_dash.copy()
            df_area['ÁREA_RESUMO'] = df_area['C.C'].apply(lambda x: 'PB' if '125.02' in str(x) and '.005' not in str(x) else ('RB' if '125.01' in str(x) and '.005' not in str(x) else ('ESP' if '.005' in str(x) else 'OUTROS')))
            df_area_count = df_area[df_area['ÁREA_RESUMO'] != 'OUTROS'].groupby('ÁREA_RESUMO').size().reset_index(name='Quantidade')
            
            if not df_area_count.empty and df_area_count['Quantidade'].sum() > 0:
                cores_areas = {'PB': '#3498db', 'RB': '#e67e22', 'ESP': '#9b59b6'}
                fig_area = px.pie(df_area_count, values="Quantidade", names="ÁREA_RESUMO", hole=0.6, color="ÁREA_RESUMO", color_discrete_map=cores_areas)
                fig_area.update_layout(margin=dict(l=20, r=20, t=10, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), height=280, showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
                st.plotly_chart(fig_area, use_container_width=True)
            else:
                st.info("Áreas não identificadas.")
                
        with col_dash3:
            st.markdown(f"**Top 10 Maiores Equipes ({filtro_dash_mo})**")
            df_enc_dash = df_dash[(df_dash["ENCARREGADO"].str.strip() != "") & (df_dash["ENCARREGADO"].isin(lista_completa_encarregados))]
            if not df_enc_dash.empty:
                top_enc = df_enc_dash["ENCARREGADO"].value_counts().head(10).reset_index()
                top_enc.columns = ["Encarregado", "Efetivo"]
                fig_top_enc = px.bar(top_enc, x="Efetivo", y="Encarregado", orientation="h", color="Efetivo", color_continuous_scale=[(0, "#0f172a"), (1, "#0ea5e9")], text="Efetivo")
                fig_top_enc.update_layout(showlegend=False, xaxis_title="", yaxis_title="", margin=dict(l=0, r=40, t=10, b=0), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), height=280)
                fig_top_enc.update_yaxes(categoryorder="total ascending")
                fig_top_enc.update_xaxes(visible=False)
                fig_top_enc.update_coloraxes(showscale=False)
                fig_top_enc.update_traces(textposition='outside', cliponaxis=False)
                st.plotly_chart(fig_top_enc, use_container_width=True)
                
        st.markdown("---")
        
        col_evolucao, col_gauge = st.columns([6, 4])
        
        with col_evolucao:
            st.markdown("**📈 Evolução Diária de Entregas de RDC (Mês Atual)**")
            if "df_historico_f1" in st.session_state and not st.session_state.df_historico_f1.empty:
                df_hist_dash = st.session_state.df_historico_f1.copy()
                df_hist_dash["DATA"] = pd.to_datetime(df_hist_dash["DATA"], errors='coerce')
                mes_atual = datetime.date.today().strftime("%Y-%m")
                df_hist_dash = df_hist_dash[df_hist_dash["DATA"].dt.strftime("%Y-%m") == mes_atual]
                
                if not df_hist_dash.empty:
                    entregas_por_dia = df_hist_dash.groupby(df_hist_dash["DATA"].dt.strftime("%Y-%m-%d")).size().reset_index(name="Entregas")
                    entregas_por_dia.columns = ["Data", "Qtd Entregue"]
                    
                    fig_evolucao = px.line(entregas_por_dia, x="Data", y="Qtd Entregue", markers=True, 
                                           title="", line_shape="spline", color_discrete_sequence=["#0ea5e9"])
                    fig_evolucao.update_layout(
                        xaxis_title="Dia", yaxis_title="RDCs Entregues",
                        margin=dict(l=0, r=20, t=10, b=0),
                        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                        font=dict(color="#e0e4ea"), height=250
                    )
                    fig_evolucao.update_traces(line=dict(width=3), marker=dict(size=8))
                    st.plotly_chart(fig_evolucao, use_container_width=True)
                else:
                    st.info("Ainda não há entregas neste mês.")
            else:
                st.info("Sem histórico de F1.")
                
        with col_gauge:
            st.markdown("**🌡️ Termômetro de Engajamento**")
            if "df_historico_f1" in st.session_state and not st.session_state.df_historico_f1.empty:
                dias_unicos = df_hist_dash["DATA"].nunique()
                dias_unicos = dias_unicos if dias_unicos > 0 else 1
                rdcs_esperados = dias_unicos * len(lista_completa_encarregados)
                rdcs_entregues = len(df_hist_dash)
                
                pct_engajamento = round((rdcs_entregues / rdcs_esperados) * 100, 1) if rdcs_esperados > 0 else 0
                
                import plotly.graph_objects as go
                fig_gauge = go.Figure(go.Indicator(
                    mode = "gauge+number",
                    value = pct_engajamento,
                    number = {'suffix': "%", 'font': {'size': 30, 'color': '#e0e4ea'}},
                    domain = {'x': [0, 1], 'y': [0, 1]},
                    gauge = {
                        'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "white"},
                        'bar': {'color': "#2c3e50"},
                        'bgcolor': "rgba(0,0,0,0)",
                        'borderwidth': 2,
                        'bordercolor': "gray",
                        'steps': [
                            {'range': [0, 60], 'color': '#ef4444'},
                            {'range': [60, 85], 'color': '#f59e0b'},
                            {'range': [85, 100], 'color': '#10b981'}],
                        'threshold': {
                            'line': {'color': "white", 'width': 4},
                            'thickness': 0.75,
                            'value': pct_engajamento}
                    }
                ))
                fig_gauge.update_layout(height=250, margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"))
                st.plotly_chart(fig_gauge, use_container_width=True)
            else:
                st.info("Sem dados suficientes.")
                
        st.markdown("---")
        st.markdown("**📦 Raio-X da Mão de Obra Indireta (MOI)**")
        col_moi1, col_moi2 = st.columns([5, 5])
        with col_moi1:
            df_moi = df_atual[df_atual["MÃO DE OBRA"].astype(str).str.strip().str.upper() == "MOI"].copy()
            if not df_moi.empty:
                moi_count = df_moi.groupby("DISCIPLINA").size().reset_index(name="Quantidade")
                moi_count = moi_count.sort_values(by="Quantidade", ascending=False).head(8)
                fig_moi = px.pie(moi_count, values="Quantidade", names="DISCIPLINA", hole=0.5, color_discrete_sequence=px.colors.sequential.YlOrRd[::-1])
                fig_moi.update_layout(margin=dict(l=20, r=20, t=10, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), height=280)
                st.plotly_chart(fig_moi, use_container_width=True)
            else:
                st.info("Nenhuma MOI na base atual.")

        st.markdown("---")
        
        st.markdown(f"**👥 Liderança: Resumo Geral de Encarregados ({filtro_dash_mo})**")
        df_enc_full = df_dash[(df_dash["ENCARREGADO"].str.strip() != "") & (df_dash["ENCARREGADO"].isin(lista_completa_encarregados))]
        if not df_enc_full.empty:
            resumo_enc = df_enc_full["ENCARREGADO"].value_counts().reset_index()
            resumo_enc.columns = ["Encarregado", "Tamanho da Equipe"]
            
            termo_enc = st.text_input("🔍 Procurar Encarregado específico:")
            if termo_enc:
                resumo_enc = resumo_enc[resumo_enc["Encarregado"].astype(str).str.contains(termo_enc, case=False, na=False)]
                
            st.dataframe(resumo_enc, hide_index=True, use_container_width=True)
            
        st.markdown("---")
        st.markdown(f"**🔍 Base Completa ({filtro_dash_mo})**")
        termo_busca = st.text_input("Buscar funcionário (Nome, Matrícula ou Função):")
        df_exibicao = df_dash[["MATRICULA", "NOME", "FUNÇÃO", "ENCARREGADO", "C.C"]].copy()
        if termo_busca:
            mask = (
                df_exibicao["NOME"].astype(str).str.contains(termo_busca, case=False, na=False) |
                df_exibicao["MATRICULA"].astype(str).str.contains(termo_busca, case=False, na=False) |
                df_exibicao["FUNÇÃO"].astype(str).str.contains(termo_busca, case=False, na=False)
            )
            df_exibicao = df_exibicao[mask]
        st.dataframe(df_exibicao, hide_index=True, use_container_width=True)
        
        # === RELATÓRIO SEMANAL EM PDF ===
        st.markdown("---")
        st.markdown("### 📄 Relatório Semanal")
        if st.button("📄 Gerar Relatório PDF da Semana", use_container_width=True, key="btn_gerar_pdf_semanal"):
            try:
                from fpdf import FPDF
                
                class RelatorioPDF(FPDF):
                    def header(self):
                        self.set_font('Helvetica', 'B', 18)
                        self.set_text_color(14, 165, 233)
                        self.cell(0, 12, 'ENESA - Relatorio Semanal', align='C', new_x="LMARGIN", new_y="NEXT")
                        self.set_font('Helvetica', '', 10)
                        self.set_text_color(100, 116, 139)
                        self.cell(0, 6, f'Gerado em: {datetime.datetime.now().strftime("%d/%m/%Y as %H:%M")}', align='C', new_x="LMARGIN", new_y="NEXT")
                        self.ln(5)
                        self.set_draw_color(14, 165, 233)
                        self.line(10, self.get_y(), 200, self.get_y())
                        self.ln(8)
                    
                    def footer(self):
                        self.set_y(-15)
                        self.set_font('Helvetica', 'I', 8)
                        self.set_text_color(128, 128, 128)
                        self.cell(0, 10, f'Pagina {self.page_no()}/{{nb}} | Sistema RDC & PDE - ENESA', align='C')
                
                pdf = RelatorioPDF()
                pdf.alias_nb_pages()
                pdf.add_page()
                pdf.set_auto_page_break(auto=True, margin=20)
                
                # KPIs Gerais
                pdf.set_font('Helvetica', 'B', 14)
                pdf.set_text_color(248, 250, 252)
                pdf.set_fill_color(30, 41, 59)
                pdf.cell(0, 10, '  INDICADORES GERAIS (KPIs)', fill=True, new_x="LMARGIN", new_y="NEXT")
                pdf.ln(5)
                
                total_geral = len(df_atual)
                total_pb = len(df_atual[df_atual["C.C"].astype(str).str.contains("125.02", na=False)]) if "C.C" in df_atual.columns else 0
                total_rb = len(df_atual[df_atual["C.C"].astype(str).str.contains("125.01", na=False)]) if "C.C" in df_atual.columns else 0
                total_mod = len(df_atual[df_atual["MÃO DE OBRA"].astype(str).str.strip().str.upper() == "MOD"]) if "MÃO DE OBRA" in df_atual.columns else 0
                total_moi = len(df_atual[df_atual["MÃO DE OBRA"].astype(str).str.strip().str.upper() == "MOI"]) if "MÃO DE OBRA" in df_atual.columns else 0
                
                pdf.set_font('Helvetica', '', 11)
                pdf.set_text_color(50, 50, 50)
                kpis = [
                    ("Efetivo Total", str(total_geral)),
                    ("PB (Caldeira de Potencia)", str(total_pb)),
                    ("RB (Caldeira de Recuperacao)", str(total_rb)),
                    ("MOD (Mao de Obra Direta)", str(total_mod)),
                    ("MOI (Mao de Obra Indireta)", str(total_moi)),
                ]
                for label, valor in kpis:
                    pdf.set_font('Helvetica', '', 11)
                    pdf.cell(120, 8, f'  {label}:', border=0)
                    pdf.set_font('Helvetica', 'B', 11)
                    pdf.set_text_color(14, 165, 233)
                    pdf.cell(0, 8, valor, new_x="LMARGIN", new_y="NEXT")
                    pdf.set_text_color(50, 50, 50)
                
                # Disciplinas
                pdf.ln(8)
                pdf.set_font('Helvetica', 'B', 14)
                pdf.set_text_color(248, 250, 252)
                pdf.set_fill_color(30, 41, 59)
                pdf.cell(0, 10, '  DISTRIBUICAO POR DISCIPLINA', fill=True, new_x="LMARGIN", new_y="NEXT")
                pdf.ln(5)
                
                if "DISCIPLINA" in df_atual.columns:
                    disc_count = df_atual.groupby("DISCIPLINA").size().reset_index(name="Qtd").sort_values("Qtd", ascending=False).head(10)
                    pdf.set_font('Helvetica', 'B', 10)
                    pdf.set_text_color(255, 255, 255)
                    pdf.set_fill_color(14, 165, 233)
                    pdf.cell(130, 8, '  Disciplina', fill=True, border=1)
                    pdf.cell(50, 8, '  Quantidade', fill=True, border=1, new_x="LMARGIN", new_y="NEXT")
                    
                    pdf.set_font('Helvetica', '', 10)
                    pdf.set_text_color(50, 50, 50)
                    for _, row in disc_count.iterrows():
                        pdf.set_fill_color(241, 245, 249)
                        pdf.cell(130, 7, f'  {str(row["DISCIPLINA"])[:40]}', border=1, fill=True)
                        pdf.cell(50, 7, f'  {row["Qtd"]}', border=1, fill=True, new_x="LMARGIN", new_y="NEXT")
                
                # F1 Ranking
                pdf.ln(8)
                pdf.set_font('Helvetica', 'B', 14)
                pdf.set_text_color(248, 250, 252)
                pdf.set_fill_color(30, 41, 59)
                pdf.cell(0, 10, '  RANKING F1 - COMPETICAO DE ENTREGAS', fill=True, new_x="LMARGIN", new_y="NEXT")
                pdf.ln(5)
                
                try:
                    hist_f1 = st.session_state.get("df_historico_f1", pd.DataFrame())
                    if not hist_f1.empty and "DATA" in hist_f1.columns and "ENCARREGADO" in hist_f1.columns:
                        mes_atual = datetime.datetime.now().strftime("%Y-%m")
                        hist_mes = hist_f1[hist_f1["DATA"].astype(str).str.startswith(mes_atual)]
                        if not hist_mes.empty:
                            rank = hist_mes.groupby("ENCARREGADO").size().reset_index(name="Entregas").sort_values("Entregas", ascending=False).head(10)
                            
                            pdf.set_font('Helvetica', 'B', 10)
                            pdf.set_text_color(255, 255, 255)
                            pdf.set_fill_color(139, 92, 246)
                            pdf.cell(20, 8, '  #', fill=True, border=1)
                            pdf.cell(110, 8, '  Encarregado', fill=True, border=1)
                            pdf.cell(50, 8, '  Entregas', fill=True, border=1, new_x="LMARGIN", new_y="NEXT")
                            
                            pdf.set_font('Helvetica', '', 10)
                            pdf.set_text_color(50, 50, 50)
                            medalhas = {0: '[OURO]', 1: '[PRATA]', 2: '[BRONZE]'}
                            for idx, (_, row) in enumerate(rank.iterrows()):
                                medal = medalhas.get(idx, f'  {idx+1}')
                                pdf.set_fill_color(241, 245, 249)
                                pdf.cell(20, 7, f'  {medal}', border=1, fill=True)
                                pdf.cell(110, 7, f'  {row["ENCARREGADO"][:35]}', border=1, fill=True)
                                pdf.cell(50, 7, f'  {row["Entregas"]}', border=1, fill=True, new_x="LMARGIN", new_y="NEXT")
                except Exception:
                    pass
                
                pdf_output = pdf.output()
                nome_pdf = f"Relatorio_Semanal_ENESA_{datetime.datetime.now().strftime('%d_%m_%Y')}.pdf"
                
                st.download_button(
                    label="⬇️ Baixar Relatório PDF",
                    data=pdf_output,
                    file_name=nome_pdf,
                    mime="application/pdf",
                    use_container_width=True
                )
                st.success(f"✅ Relatório gerado com sucesso! Clique acima para baixar.")
            except ImportError:
                st.error("⚠️ Biblioteca `fpdf2` não encontrada. Rode: `pip install fpdf2`")
            except Exception as e:
                st.error(f"❌ Erro ao gerar PDF: {e}")

    if tab_resumo is not None:
      with tab_resumo:
        st.markdown("### 📅 Resumo Diário")
        
        # --- FIX: Filtro de data para ver o resumo de qualquer dia ---
        data_resumo = st.date_input("Selecione a Data do Resumo:", datetime.date.today())
        data_filtro_str = data_resumo.strftime("%Y-%m-%d")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- MÓDULO 1: Controle de Chuva / Paradas ---
        st.markdown("### 🌤️ Clima e Interferências Diárias")
        
        if "df_clima" not in st.session_state:
            if os.path.exists("clima.json"):
                st.session_state.df_clima = pd.read_json("clima.json")
            else:
                st.session_state.df_clima = pd.DataFrame(columns=["Data", "Condição", "Horas_Paradas", "Observação"])
                
        clima_hoje = st.session_state.df_clima[st.session_state.df_clima["Data"] == data_filtro_str]
        cond_atual = "Ensolarado"
        horas_atual = 0.0
        obs_atual = ""
        if not clima_hoje.empty:
            cond_atual = str(clima_hoje.iloc[0].get("Condição", "Ensolarado"))
            try:
                horas_atual = float(clima_hoje.iloc[0].get("Horas_Paradas", 0.0))
            except:
                pass
            obs_atual = str(clima_hoje.iloc[0].get("Observação", ""))
        
        with st.form("form_clima"):
            col_clima1, col_clima2, col_clima3 = st.columns([2, 2, 4])
            with col_clima1:
                opcoes_cond = ["Ensolarado", "Nublado", "Chuva Leve", "Chuva Forte (Impeditiva)"]
                try:
                    idx_cond = opcoes_cond.index(cond_atual)
                except:
                    idx_cond = 0
                condicao = st.selectbox("Condição Climática", opcoes_cond, index=idx_cond)
            with col_clima2:
                horas_paradas = st.number_input("Horas Perdidas", min_value=0.0, max_value=24.0, step=0.5, value=horas_atual)
            with col_clima3:
                obs_clima = st.text_input("Observação / Justificativa", value=obs_atual)
                
            if st.form_submit_button("💾 Salvar Registro de Clima"):
                novo_clima = pd.DataFrame([{"Data": data_filtro_str, "Condição": condicao, "Horas_Paradas": horas_paradas, "Observação": obs_clima}])
                st.session_state.df_clima = st.session_state.df_clima[st.session_state.df_clima["Data"] != data_filtro_str]
                st.session_state.df_clima = pd.concat([st.session_state.df_clima, novo_clima], ignore_index=True)
                st.session_state.df_clima.to_json("clima.json", orient="records")
                st.success(f"Clima para o dia {data_filtro_str} salvo com sucesso!")
                
        if not st.session_state.df_clima.empty:
            with st.expander("Ver Histórico de Clima do Mês (Dossiê de Pleito)"):
                df_clima_mes = st.session_state.df_clima[st.session_state.df_clima["Data"].str.startswith(datetime.date.today().strftime("%Y-%m"))]
                st.dataframe(df_clima_mes, use_container_width=True, hide_index=True)
                
        st.markdown("---")

        if st.toggle("➕ Lançar RDC Manualmente (Para papéis ilegíveis ou atrasados)", key="toggle_manual_resumo"):
            with st.form("form_resumo_manual"):
                st.info("💡 Você pode colar a lista inteira de encarregados aqui (um por linha ou separados por vírgula). O robô vai verificar: se a IA já tiver lido, ele ignora. Se faltou, ele adiciona!")
                col_m1, col_m2 = st.columns([1, 2])
                with col_m1:
                    data_manual_resumo = st.date_input("Data do RDC a lançar:", value=data_resumo, key="data_manual_resumo")
                with col_m2:
                    nomes_colados_resumo = st.text_area("Cole os nomes dos Encarregados", height=120, key="nomes_colados_resumo")
                
                btn_manual_resumo = st.form_submit_button("Processar Lista e Lançar no Sistema")
                if btn_manual_resumo and nomes_colados_resumo.strip():
                    import re
                    import difflib
                    data_str_resumo = data_manual_resumo.strftime("%Y-%m-%d")
                    
                    lista_suja = [n.strip().upper() for n in re.split(r'[\n,;]', nomes_colados_resumo) if n.strip()]
                    
                    novos_registros = []
                    nomes_ja_existentes = 0
                    nomes_nao_encontrados = []
                    
                    for nome_sujo in lista_suja:
                        match = difflib.get_close_matches(nome_sujo, lista_completa_encarregados, n=1, cutoff=0.55)
                        if match:
                            nome_oficial = match[0]
                            ja_existe = ((st.session_state.df_historico_f1["DATA"] == data_str_resumo) & (st.session_state.df_historico_f1["ENCARREGADO"] == nome_oficial)).any()
                            if ja_existe:
                                nomes_ja_existentes += 1
                            else:
                                novos_registros.append({"DATA": data_str_resumo, "ENCARREGADO": nome_oficial})
                        else:
                            nomes_nao_encontrados.append(nome_sujo)
                            
                    if novos_registros:
                        df_novos = pd.DataFrame(novos_registros)
                        
                        if conn and not st.session_state.get('force_use_local', False):
                            try:
                                df_fresco = conn.read(worksheet="Historico_F1", ttl=0)
                                if not df_fresco.empty:
                                    df_fresco = df_fresco.dropna(how='all')
                                    df_final = pd.concat([df_fresco, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                                else:
                                    df_final = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                                
                                conn.update(worksheet="Historico_F1", data=df_final)
                                st.session_state.df_historico_f1 = df_final
                                st.cache_data.clear()
                                st.success(f"✅ {len(novos_registros)} novos RDCs adicionados e sincronizados com a nuvem! ({nomes_ja_existentes} já constavam).")
                            except Exception as e:
                                st.error(f"Erro ao salvar na nuvem: {e}")
                                st.session_state.df_historico_f1 = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                        else:
                            st.session_state.df_historico_f1 = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                            st.success(f"✅ {len(novos_registros)} novos RDCs adicionados localmente! ({nomes_ja_existentes} já constavam).")
                    elif nomes_ja_existentes > 0:
                        st.warning(f"⚠️ Todos os nomes reconhecidos ({nomes_ja_existentes}) já estavam devidamente lançados neste dia!")
                        
                    if nomes_nao_encontrados:
                        st.error(f"❌ Não encontrei na lista oficial (verifique a escrita): {', '.join(nomes_nao_encontrados)}")
                        
                    time.sleep(3)
                    st.rerun()
                    
        st.markdown("<hr style='margin-top:0px; margin-bottom:20px'>", unsafe_allow_html=True)

        
        df_hoje = pd.DataFrame()
        if "df_historico_f1" in st.session_state and not st.session_state.df_historico_f1.empty:
            df_hist = st.session_state.df_historico_f1.copy()
            df_hist["DATA_STR"] = pd.to_datetime(df_hist["DATA"], errors="coerce").dt.strftime("%Y-%m-%d")
            df_hoje = df_hist[df_hist["DATA_STR"] == data_filtro_str]
            
        encarregados_esperados = len(lista_completa_encarregados)
        entregues_hoje_lista = [e for e in df_hoje["ENCARREGADO"].unique() if e in lista_completa_encarregados] if not df_hoje.empty else []
        encarregados_entregues = len(entregues_hoje_lista)
        encarregados_pendentes = encarregados_esperados - encarregados_entregues
        
        c1, c2, c3 = st.columns(3)
        c1.metric("🎯 Esperados", encarregados_esperados)
        c2.metric("✅ Entregues", encarregados_entregues)
        c3.metric("⏳ Pendentes", encarregados_pendentes)
        
        st.markdown("---")
        if encarregados_pendentes > 0:
            st.error(f"**Atenção:** {encarregados_pendentes} encarregados ainda não entregaram o RDC nesta data ({data_filtro_str}).")
            entregues_list = df_hoje["ENCARREGADO"].unique() if not df_hoje.empty else []
            pendentes_list = [e for e in lista_completa_encarregados if e not in entregues_list]
            
            # Botão de Gerar PDF
            col_pdf, col_gap = st.columns([2, 3])
            with col_pdf:
                if st.button("📄 Gerar PDF de Cobrança", use_container_width=True):
                    from fpdf import FPDF
                    import tempfile
                    
                    class PDF(FPDF):
                        def header(self):
                            self.set_font('Helvetica', 'B', 15)
                            self.set_text_color(0, 0, 0)
                            self.cell(0, 10, 'Relatorio de Pendencias - RDC', 0, 1, 'C')
                            self.set_font('Helvetica', 'I', 10)
                            self.cell(0, 10, f'Data Referencia: {data_filtro_str} (Gerado em: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M")})', 0, 1, 'C')
                            self.ln(5)
                    
                    pdf = PDF()
                    pdf.add_page()
                    pdf.set_font('Helvetica', 'B', 12)
                    pdf.set_text_color(200, 0, 0)
                    pdf.cell(0, 10, f'{len(pendentes_list)} Encarregados nao entregaram o RDC nesta data ({data_filtro_str}):', 0, 1, 'L')
                    pdf.ln(2)
                    
                    pdf.set_font('Helvetica', '', 10)
                    pdf.set_text_color(0, 0, 0)
                    for pendente in sorted(pendentes_list):
                        # Evitar problemas de encoding no PDF básico
                        nome_p = str(pendente).encode('latin-1', 'replace').decode('latin-1')
                        pdf.cell(0, 8, f'- {nome_p}', 0, 1, 'L')
                        
                    nome_pdf = f"Cobranca_RDC_{data_filtro_str}.pdf"
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        pdf.output(tmp.name)
                        with open(tmp.name, "rb") as f:
                            pdf_bytes = f.read()
                            
                        # Backup Drive
                        success, msg = backup_google_drive(tmp.name, "application/pdf", nome_pdf)
                        if success:
                            st.toast("☁️ Relatório PDF salvo no Drive!")
                    
                    st.download_button(
                        label="⬇️ Baixar PDF",
                        data=pdf_bytes,
                        file_name=nome_pdf,
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                    
            st.markdown("<br>", unsafe_allow_html=True)
            df_pend = pd.DataFrame({f"Encarregados Pendentes ({data_filtro_str})": sorted(pendentes_list)})
            st.dataframe(df_pend, hide_index=True, use_container_width=True)
        else:
            st.success(f"🎉 Todos os RDCs desta data ({data_filtro_str}) já foram entregues!")

    if tab_emissao is not None:
      with tab_emissao:
        st.markdown("### Emissão de RDC")
        if not lista_encarregados_base:
            st.warning("Nenhum encarregado encontrado na base.")
        else:
            encarregado_sel = st.selectbox("Escolha o Encarregado:", lista_encarregados_base)
            equipe = df_atual[df_atual["ENCARREGADO"] == encarregado_sel]
            st.markdown("")
            st.markdown(f"""<div style="background: {cor_card}; border-radius: 10px; padding: 20px; border: 1px solid {cor_borda}; margin-bottom: 16px;"><div style="text-align: center; border-bottom: 2px solid {cor_azul}; padding-bottom: 12px; margin-bottom: 12px;"><h3 style="margin: 0; font-size: 1.2rem; color: {cor_texto} !important;">RDC - Relatório Diário de Campo</h3><p style="color: {cor_texto_sub}; margin: 4px 0 0 0; font-size: 0.85rem;">{nome_site}</p></div><table style="width: 100%; color: {cor_texto}; font-size: 0.9rem;"><tr><td style="padding: 4px 0;"><strong>Encarregado:</strong></td><td>{encarregado_sel}</td></tr><tr><td style="padding: 4px 0;"><strong>Data:</strong></td><td>{datetime.datetime.now().strftime("%d/%m/%Y")}</td></tr><tr><td style="padding: 4px 0;"><strong>Efetivo:</strong></td><td>{len(equipe)} colaborador(es)</td></tr></table></div>""", unsafe_allow_html=True)
            st.dataframe(equipe[["MATRICULA", "NOME", "FUNÇÃO"]].reset_index(drop=True), hide_index=True, use_container_width=True)
            st.markdown("")
            

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("🟢 GERAR EXCEL", type="primary", use_container_width=True):
                    wb = preencher_excel(equipe, encarregado_sel)
                    if wb:
                        nome_limpo = encarregado_sel.replace(" ", "_")
                        nome_arquivo = f"RDC_{nome_limpo}.xlsx"
                        buffer = io.BytesIO()
                        wb.save(buffer)
                        wb.close()
                        try:
                            hoje = datetime.datetime.now()
                            pasta_hist = os.path.join(pasta_base, "Historico_RDC", str(hoje.year), f"{hoje.month:02d}_{hoje.strftime('%B')}")
                            os.makedirs(pasta_hist, exist_ok=True)
                            
                            caminho_local = os.path.join(pasta_hist, f"{hoje.strftime('%d_%H%M')}_{nome_arquivo}")
                            with open(caminho_local, "wb") as f:
                                f.write(buffer.getvalue())
                                
                            success, msg = backup_google_drive(caminho_local, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{hoje.strftime('%d_%H%M')}_{nome_arquivo}")
                            if success:
                                st.toast("☁️ Backup salvo no Google Drive!")
                        except Exception as e:
                            pass
                        buffer.seek(0)
                        st.download_button("⬇️ Baixar Planilha", data=buffer, file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        st.success("✅ Gerado!")
                    else:
                        st.error("Modelo não encontrado. Faça upload do MODELO.xlsx.")
            with col_btn2:
                if st.button("🚀 GERAR TODOS (.ZIP)", use_container_width=True):
                    with st.spinner("Gerando..."):
                        try:
                            zip_buffer = io.BytesIO()
                            qtd = 0
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                for enc in lista_encarregados_base:
                                    eq = df_atual[df_atual["ENCARREGADO"] == enc]
                                    if len(eq) > 0:
                                        wb_e = preencher_excel(eq, enc)
                                        if wb_e:
                                            buf = io.BytesIO()
                                            wb_e.save(buf)
                                            wb_e.close()
                                            try:
                                                hoje = datetime.datetime.now()
                                                pasta_hist = os.path.join(pasta_base, "Historico_RDC", str(hoje.year), f"{hoje.month:02d}_{hoje.strftime('%B')}")
                                                os.makedirs(pasta_hist, exist_ok=True)
                                                n = enc.replace(" ", "_")
                                                with open(os.path.join(pasta_hist, f"{hoje.strftime('%d_%H%M')}_RDC_{n}.xlsx"), "wb") as f:
                                                    f.write(buf.getvalue())
                                            except Exception:
                                                pass
                                            buf.seek(0)
                                            zf.writestr(f"RDC_{enc.replace(' ', '_')}.xlsx", buf.read())
                                            qtd += 1
                            zip_buffer.seek(0)
                            nome_zip = f"LOTE_RDC_{datetime.datetime.now().strftime('%d_%m_%Y')}.zip"
                            st.download_button(f"⬇️ Baixar Todos ({qtd} arquivos)", data=zip_buffer, file_name=nome_zip, mime="application/zip", use_container_width=True)
                            
                            try:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_zip:
                                    tmp_zip.write(zip_buffer.getvalue())
                                    tmp_zip_path = tmp_zip.name
                                success, msg = backup_google_drive(tmp_zip_path, "application/zip", nome_zip)
                                if success:
                                    st.toast("☁️ Lote salvo no Google Drive!")
                                os.remove(tmp_zip_path)
                            except:
                                pass
                                
                            st.success(f"✅ {qtd} planilhas geradas!")
                        except Exception as e:
                            st.error(f"Erro: {e}")

    with tab_f1:
        st.markdown("### 🏎️ Competição F1 - Entrega de RDC")
        st.markdown("Acompanhamento mensal da entrega dos Relatórios Diários de Campo (RDC).")
        
        # A lista completa foi movida para cima para ser compartilhada com a aba de Resumo Diário
        
        # === PAINEL: GERENCIAR LISTA DE ENCARREGADOS ===
        if st.toggle("👥 Gerenciar Lista de Encarregados do F1", key="toggle_gerenciar_lista_f1"):
            st.markdown("""
            <div style="background: rgba(14, 165, 233, 0.1); border: 1px solid rgba(14, 165, 233, 0.3); border-radius: 12px; padding: 15px; margin-bottom: 15px;">
                <p style="margin: 0; color: #94a3b8; font-size: 14px;">⚙️ Aqui você pode <b style="color: #0ea5e9;">adicionar</b> ou <b style="color: #ef4444;">remover</b> encarregados do controle F1. As alterações são salvas automaticamente.</p>
            </div>
            """, unsafe_allow_html=True)
            
            col_add, col_rem = st.columns(2)
            
            with col_add:
                with st.form("form_add_enc_f1"):
                    st.markdown("**➕ Adicionar Encarregado**")
                    novo_nome = st.text_input("Nome completo do Encarregado:", placeholder="Ex: JOÃO DA SILVA SOUZA")
                    btn_add = st.form_submit_button("Adicionar à Lista", type="primary", use_container_width=True)
                    if btn_add and novo_nome.strip():
                        nome_upper = novo_nome.strip().upper()
                        if nome_upper in lista_completa_encarregados:
                            st.warning(f"⚠️ '{nome_upper}' já está na lista!")
                        else:
                            encarregados_f1_oficial.append(nome_upper)
                            with open(caminho_f1_json, "w", encoding="utf-8") as f:
                                json.dump(encarregados_f1_oficial, f, ensure_ascii=False, indent=2)
                            st.success(f"✅ '{nome_upper}' adicionado com sucesso!")
                            time.sleep(2)
                            st.rerun()
            
            with col_rem:
                with st.form("form_rem_enc_f1"):
                    st.markdown("**🗑️ Remover Encarregado**")
                    enc_remover = st.multiselect("Selecione quem remover:", lista_completa_encarregados)
                    btn_rem = st.form_submit_button("Remover da Lista", type="primary", use_container_width=True)
                    if btn_rem and enc_remover:
                        lista_atualizada = [e for e in encarregados_f1_oficial if e.upper() not in enc_remover]
                        with open(caminho_f1_json, "w", encoding="utf-8") as f:
                            json.dump(lista_atualizada, f, ensure_ascii=False, indent=2)
                        st.success(f"✅ {len(enc_remover)} encarregado(s) removido(s)!")
                        time.sleep(2)
                        st.rerun()
            
            st.caption(f"📋 Total atual na lista: **{len(lista_completa_encarregados)}** encarregados")
        
        # === PAINEL: ABONAR FALTAS ===
        if st.toggle("⏸️ Abonar Faltas (Folga / Atestado / Feriado)", key="toggle_abono_f1"):
            st.markdown("""
            <div style="background: rgba(245, 158, 11, 0.1); border: 1px solid rgba(245, 158, 11, 0.3); border-radius: 12px; padding: 15px; margin-bottom: 15px;">
                <p style="margin: 0; color: #94a3b8; font-size: 14px;">📝 Marque os dias em que o encarregado <b style="color: #f59e0b;">não precisava</b> entregar o RDC. Esses dias aparecerão como <b style="color: #f59e0b;">⏸️</b> na tabela em vez de ❌.</p>
            </div>
            """, unsafe_allow_html=True)
            
            with st.form("form_abono_f1"):
                col_ab1, col_ab2 = st.columns(2)
                with col_ab1:
                    datas_abono = st.date_input("📅 Data(s) do Abono:", value=datetime.date.today(), key="datas_abono_input")
                with col_ab2:
                    motivo_abono = st.selectbox("Motivo:", ["FOLGA", "ATESTADO MÉDICO", "FERIADO", "CHUVA / INTEMPÉRIE", "FALTA JUSTIFICADA", "OUTRO"])
                
                encs_abono = st.multiselect("Selecione os Encarregados para abonar:", lista_completa_encarregados, key="encs_abono_multi")
                
                btn_abono = st.form_submit_button("✅ Registrar Abono", type="primary", use_container_width=True)
                if btn_abono and encs_abono:
                    novos_abonos = []
                    # datas_abono pode ser uma data única ou uma tupla de datas
                    if isinstance(datas_abono, (list, tuple)):
                        lista_datas = [d.strftime("%Y-%m-%d") for d in datas_abono]
                    else:
                        lista_datas = [datas_abono.strftime("%Y-%m-%d")]
                    
                    for data_ab in lista_datas:
                        for enc_ab in encs_abono:
                            ja_existe = False
                            if not st.session_state.df_f1_excecoes.empty:
                                ja_existe = ((st.session_state.df_f1_excecoes["DATA"] == data_ab) & (st.session_state.df_f1_excecoes["ENCARREGADO"] == enc_ab)).any()
                            if not ja_existe:
                                novos_abonos.append({"DATA": data_ab, "ENCARREGADO": enc_ab, "MOTIVO": motivo_abono})
                    
                    if novos_abonos:
                        df_novos_ab = pd.DataFrame(novos_abonos)
                        st.session_state.df_f1_excecoes = pd.concat([st.session_state.df_f1_excecoes, df_novos_ab], ignore_index=True)
                        st.session_state.df_f1_excecoes.to_csv(caminho_f1_excecoes, index=False)
                        st.success(f"✅ {len(novos_abonos)} abono(s) registrado(s) com sucesso!")
                        time.sleep(2)
                        st.rerun()
                    else:
                        st.info("ℹ️ Todos os abonos selecionados já estavam cadastrados.")
            
            # Mostrar abonos existentes do mês atual
            if not st.session_state.df_f1_excecoes.empty:
                st.markdown("**Abonos registrados:**")
                df_exc_show = st.session_state.df_f1_excecoes.copy()
                df_exc_show = df_exc_show.sort_values("DATA", ascending=False).head(20)
                st.dataframe(df_exc_show, hide_index=True, use_container_width=True)
                
                if st.button("🗑️ Limpar Todos os Abonos", key="btn_limpar_abonos"):
                    st.session_state.df_f1_excecoes = pd.DataFrame(columns=["DATA", "ENCARREGADO", "MOTIVO"])
                    if os.path.exists(caminho_f1_excecoes):
                        os.remove(caminho_f1_excecoes)
                    st.success("✅ Todos os abonos foram removidos!")
                    time.sleep(2)
                    st.rerun()
        
        # --- LANÇAMENTO MANUAL ---
        if st.toggle("➕ Lançar RDC Manualmente (Para papéis ilegíveis ou atrasados)"):
            with st.form("form_f1_manual"):
                st.info("💡 Você pode colar a lista inteira de encarregados aqui (um por linha ou separados por vírgula). O robô vai verificar: se a IA já tiver lido, ele ignora. Se faltou, ele adiciona!")
                col_m1, col_m2 = st.columns([1, 2])
                with col_m1:
                    data_manual = st.date_input("Data do RDC")
                with col_m2:
                    nomes_colados = st.text_area("Cole os nomes dos Encarregados", height=120)
                
                btn_manual = st.form_submit_button("Processar Lista e Lançar no F1")
                if btn_manual and nomes_colados.strip():
                    import re
                    import difflib
                    data_str = data_manual.strftime("%Y-%m-%d")
                    
                    lista_suja = [n.strip().upper() for n in re.split(r'[\n,;]', nomes_colados) if n.strip()]
                    
                    novos_registros = []
                    nomes_ja_existentes = 0
                    nomes_nao_encontrados = []
                    
                    for nome_sujo in lista_suja:
                        match = difflib.get_close_matches(nome_sujo, lista_completa_encarregados, n=1, cutoff=0.55)
                        if match:
                            nome_oficial = match[0]
                            ja_existe = ((st.session_state.df_historico_f1["DATA"] == data_str) & (st.session_state.df_historico_f1["ENCARREGADO"] == nome_oficial)).any()
                            if ja_existe:
                                nomes_ja_existentes += 1
                            else:
                                novos_registros.append({"DATA": data_str, "ENCARREGADO": nome_oficial})
                        else:
                            nomes_nao_encontrados.append(nome_sujo)
                            
                    if novos_registros:
                        df_novos = pd.DataFrame(novos_registros)
                        
                        if conn and not st.session_state.get('force_use_local', False):
                            try:
                                # Puxar a versão mais fresca da nuvem para evitar sobrescrever a IA rodando em outra aba
                                df_fresco = conn.read(worksheet="Historico_F1", ttl=0)
                                if not df_fresco.empty:
                                    df_fresco = df_fresco.dropna(how='all')
                                    df_final = pd.concat([df_fresco, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                                else:
                                    df_final = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                                
                                conn.update(worksheet="Historico_F1", data=df_final)
                                st.session_state.df_historico_f1 = df_final
                                st.cache_data.clear()
                                st.success(f"✅ {len(novos_registros)} novos RDCs adicionados e sincronizados com a nuvem! ({nomes_ja_existentes} já constavam).")
                            except Exception as e:
                                st.error(f"Erro ao salvar na nuvem: {e}")
                                st.session_state.df_historico_f1 = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                        else:
                            st.session_state.df_historico_f1 = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                            st.success(f"✅ {len(novos_registros)} novos RDCs adicionados localmente! ({nomes_ja_existentes} já constavam).")
                    elif nomes_ja_existentes > 0:
                        st.warning(f"⚠️ Todos os nomes reconhecidos ({nomes_ja_existentes}) já estavam devidamente lançados neste dia!")
                        
                    if nomes_nao_encontrados:
                        st.error(f"❌ Não encontrei na lista oficial (verifique a escrita): {', '.join(nomes_nao_encontrados)}")
                        
                    time.sleep(4)
                    st.rerun()
        # -------------------------

        # Preparar dados de data do histórico
        df_hist = st.session_state.df_historico_f1.copy()
        if not df_hist.empty:
            df_hist["DATA"] = pd.to_datetime(df_hist["DATA"], format="%Y-%m-%d", errors="coerce")
            df_hist = df_hist.dropna(subset=["DATA"])
            df_hist["MES_ANO"] = df_hist["DATA"].dt.strftime("%Y-%m")
            meses_disponiveis = sorted(df_hist["MES_ANO"].unique(), reverse=True)
        else:
            meses_disponiveis = [datetime.date.today().strftime("%Y-%m")]
            
        mes_selecionado = st.selectbox("📅 Selecione o Mês para Análise:", meses_disponiveis)
        
        if not df_hist.empty:
            df_mes = df_hist[df_hist["MES_ANO"] == mes_selecionado]
        else:
            df_mes = pd.DataFrame(columns=["DATA", "ENCARREGADO"])
            
        import calendar
        ano, mes = map(int, mes_selecionado.split('-'))
        num_dias = calendar.monthrange(ano, mes)[1]
        
        # Montar a Matriz com a lista oficial + qualquer outro nome que já tenha entregue no mês
        nomes_no_mes = df_mes["ENCARREGADO"].dropna().unique().tolist() if not df_mes.empty else []
        todos_encarregados_matriz = sorted(list(set(lista_completa_encarregados + nomes_no_mes)))
        # Remover nomes inválidos da matriz
        todos_encarregados_matriz = [e for e in todos_encarregados_matriz if e.strip() != "" and e.upper() != "AJUSTAR NOME"]
        
        dias_str = [str(d) for d in range(1, num_dias + 1)]
        
        # Identificar sábados e domingos
        dias_fim_de_semana = set()
        for d in range(1, num_dias + 1):
            data_check = datetime.date(ano, mes, d)
            if data_check.weekday() >= 5:  # 5=Sábado, 6=Domingo
                dias_fim_de_semana.add(str(d))
        
        dias_uteis = [d for d in dias_str if d not in dias_fim_de_semana]
        
        matriz = pd.DataFrame(index=todos_encarregados_matriz, columns=dias_str)
        # Preencher dias úteis com ❌ e fins de semana com ➖
        for col in dias_str:
            if col in dias_fim_de_semana:
                matriz[col] = "➖"
            else:
                matriz[col] = "❌"
        
        for _, row in df_mes.iterrows():
            dia = str(row["DATA"].day)
            enc = row["ENCARREGADO"]
            if enc in matriz.index and dia not in dias_fim_de_semana:
                matriz.loc[enc, dia] = "✅"
        
        # Aplicar Abonos (substituir ❌ por ⏸️ para dias com exceção cadastrada)
        if not st.session_state.df_f1_excecoes.empty:
            df_exc_mes = st.session_state.df_f1_excecoes.copy()
            df_exc_mes["DATA"] = pd.to_datetime(df_exc_mes["DATA"], errors='coerce')
            df_exc_mes = df_exc_mes.dropna(subset=["DATA"])
            df_exc_mes = df_exc_mes[df_exc_mes["DATA"].dt.strftime("%Y-%m") == mes_selecionado]
            
            for _, row_exc in df_exc_mes.iterrows():
                dia_exc = str(row_exc["DATA"].day)
                enc_exc = row_exc["ENCARREGADO"]
                if enc_exc in matriz.index and dia_exc not in dias_fim_de_semana:
                    if matriz.loc[enc_exc, dia_exc] == "❌":
                        matriz.loc[enc_exc, dia_exc] = "⏸️"
                
        # Total conta apenas dias úteis (ignora fins de semana e abonos)
        matriz["Total"] = (matriz[dias_uteis] == "✅").sum(axis=1)
        
        # Adicionar o total do dia no próprio cabeçalho da coluna (em cima dos dias)
        total_por_dia = (matriz[dias_str] == "✅").sum(axis=0)
        novas_colunas = {}
        for dia in dias_str:
            if dia in dias_fim_de_semana:
                data_check = datetime.date(ano, mes, int(dia))
                nome_dia = "SAB" if data_check.weekday() == 5 else "DOM"
                novas_colunas[dia] = f"{dia}\n({nome_dia})"
            else:
                novas_colunas[dia] = f"{dia}\n({total_por_dia[dia]})"
        matriz.rename(columns=novas_colunas, inplace=True)
        
        total_entregue = matriz["Total"].sum()
        col_tit, col_met = st.columns([3, 1])
        with col_tit:
            st.markdown(f"#### 📊 Matriz de Entregas - {mes_selecionado}")
        with col_met:
            st.metric("📄 Total de RDCs Entregues", total_entregue)
            
        # Alerta de Devedores (3 dias úteis)
        if mes_selecionado == datetime.date.today().strftime("%Y-%m"):
            hoje_int = datetime.date.today().day
            dias_passados = [d for d in dias_uteis if int(d) <= hoje_int]
            devedores = []
            if len(dias_passados) >= 3:
                ultimos_3 = dias_passados[-3:]
                for enc in matriz.index:
                    if all(matriz.loc[enc, novas_colunas[dia]] == "❌" for dia in ultimos_3):
                        devedores.append(enc)
            if devedores:
                st.error(f"🚨 **ALERTA CRÍTICO:** {len(devedores)} encarregados não entregaram RDC nos últimos 3 dias úteis.")
                if st.toggle("👀 Mostrar lista de encarregados com pendência crítica"):
                    dados_dev = []
                    for enc in devedores:
                        entregues_ate_hoje = sum(1 for d in dias_passados if matriz.loc[enc, novas_colunas[d]] == "✅")
                        pendentes_ate_hoje = len(dias_passados) - entregues_ate_hoje
                        dados_dev.append({"Encarregados": enc, "Faltas Totais no Mês": pendentes_ate_hoje})
                    
                    df_dev = pd.DataFrame(dados_dev).sort_values(by="Faltas Totais no Mês", ascending=False)
                    st.dataframe(df_dev, hide_index=True, use_container_width=True)
        
        def cor_fundo(valor):
            if valor == "✅":
                return "background-color: rgba(74, 222, 128, 0.2); color: #4ade80;"
            elif valor == "❌":
                return "background-color: rgba(255, 75, 75, 0.2); color: #ff4b4b;"
            elif valor == "⏸️":
                return "background-color: rgba(245, 158, 11, 0.2); color: #f59e0b;"
            elif valor == "➖":
                return "background-color: rgba(128, 128, 128, 0.2); color: #888;"
            return ""
            
        try:
            matriz_estilizada = matriz.style.map(cor_fundo)
        except AttributeError:
            matriz_estilizada = matriz.style.applymap(cor_fundo)
            
        st.dataframe(matriz_estilizada, use_container_width=True)
        
        # === MARCAR / DESMARCAR ENTREGA MANUALMENTE ===
        if st.toggle("✏️ Marcar ou Desmarcar Entrega de um Dia", key="toggle_marcar_dia_f1"):
            st.markdown("""
            <div style="background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.3); border-radius: 12px; padding: 15px; margin-bottom: 15px;">
                <p style="margin: 0; color: #94a3b8; font-size: 14px;">Selecione os encarregados e o dia para colocar <b style="color: #10b981;">✅</b> ou tirar (voltar para <b style="color: #ef4444;">❌</b>).</p>
            </div>
            """, unsafe_allow_html=True)
            
            col_mk1, col_mk2, col_mk3 = st.columns([3, 1, 1])
            with col_mk1:
                encs_marcar = st.multiselect("Encarregado(s):", todos_encarregados_matriz, key="encs_marcar_dia")
            with col_mk2:
                dia_marcar = st.selectbox("Dia:", [int(d) for d in dias_uteis], key="dia_marcar_sel")
            with col_mk3:
                acao_marcar = st.selectbox("Ação:", ["✅ Marcar Entregue", "❌ Desmarcar"], key="acao_marcar_sel")
            
            if st.button("Aplicar", type="primary", use_container_width=True, key="btn_aplicar_marcar"):
                if encs_marcar:
                    data_str = f"{ano}-{str(mes).zfill(2)}-{str(dia_marcar).zfill(2)}"
                    
                    if "✅" in acao_marcar:
                        # Adicionar ao histórico F1
                        novos = []
                        for enc_mk in encs_marcar:
                            ja_existe = ((st.session_state.df_historico_f1["DATA"] == data_str) & (st.session_state.df_historico_f1["ENCARREGADO"] == enc_mk)).any()
                            if not ja_existe:
                                novos.append({"DATA": data_str, "ENCARREGADO": enc_mk})
                        if novos:
                            df_novos_mk = pd.DataFrame(novos)
                            st.session_state.df_historico_f1 = pd.concat([st.session_state.df_historico_f1, df_novos_mk], ignore_index=True)
                            
                            if conn and not st.session_state.get('force_use_local', False):
                                try:
                                    conn.update(worksheet="Historico_F1", data=st.session_state.df_historico_f1)
                                    st.cache_data.clear()
                                except Exception:
                                    pass
                            
                            st.success(f"✅ {len(novos)} entrega(s) marcada(s) no dia {dia_marcar}!")
                        else:
                            st.info("ℹ️ Todos já estavam marcados nesse dia.")
                    else:
                        # Remover do histórico F1
                        removidos = 0
                        for enc_mk in encs_marcar:
                            mask = (st.session_state.df_historico_f1["DATA"] == data_str) & (st.session_state.df_historico_f1["ENCARREGADO"] == enc_mk)
                            if mask.any():
                                st.session_state.df_historico_f1 = st.session_state.df_historico_f1[~mask]
                                removidos += 1
                        
                        if removidos > 0:
                            if conn and not st.session_state.get('force_use_local', False):
                                try:
                                    conn.update(worksheet="Historico_F1", data=st.session_state.df_historico_f1)
                                    st.cache_data.clear()
                                except Exception:
                                    pass
                            st.success(f"❌ {removidos} entrega(s) desmarcada(s) no dia {dia_marcar}!")
                        else:
                            st.info("ℹ️ Nenhum deles estava marcado nesse dia.")
                    
                    time.sleep(2)
                    st.rerun()
                else:
                    st.warning("⚠️ Selecione pelo menos um encarregado.")
        
        # --- EXPORTAR PARA RH ---
        buffer_rh = io.BytesIO()
        matriz_export = matriz.reset_index().rename(columns={"index": "ENCARREGADO"})
        matriz_export.to_excel(buffer_rh, index=False, engine='openpyxl')
        buffer_rh.seek(0)
        
        st.download_button(
            label="📥 Baixar Planilha do Mês para o RH (.xlsx)",
            data=buffer_rh,
            file_name=f"Relatorio_RH_F1_{mes_selecionado}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        # ------------------------
        st.markdown("---")
        st.markdown("#### 🏆 Pódio do Mês (Top Melhores Entregas)")
        
        ranking = matriz[["Total"]].sort_values(by="Total", ascending=False).reset_index()
        ranking.columns = ["ENCARREGADO", "ENTREGAS"]
        
        st.success("🥇 Os 3 que MAIS entregaram RDCs")
        top3 = ranking.head(3)
        
        if len(top3) >= 3:
            n1 = top3.iloc[0]["ENCARREGADO"].split()[0] + " " + (top3.iloc[0]["ENCARREGADO"].split()[-1] if len(top3.iloc[0]["ENCARREGADO"].split())>1 else "")
            t1 = top3.iloc[0]["ENTREGAS"]
            n2 = top3.iloc[1]["ENCARREGADO"].split()[0] + " " + (top3.iloc[1]["ENCARREGADO"].split()[-1] if len(top3.iloc[1]["ENCARREGADO"].split())>1 else "")
            t2 = top3.iloc[1]["ENTREGAS"]
            n3 = top3.iloc[2]["ENCARREGADO"].split()[0] + " " + (top3.iloc[2]["ENCARREGADO"].split()[-1] if len(top3.iloc[2]["ENCARREGADO"].split())>1 else "")
            t3 = top3.iloc[2]["ENTREGAS"]
            
            html_podio = f"""
            <div style="display: flex; justify-content: center; align-items: flex-end; height: 190px; gap: 15px; margin-top: 30px; margin-bottom: 20px;">
                <!-- 2 Lugar -->
                <div style="display: flex; flex-direction: column; align-items: center; width: 130px; transition: transform 0.3s;" onmouseover="this.style.transform='translateY(-5px)'" onmouseout="this.style.transform='translateY(0)'">
                    <div style="font-size: 13px; color: #cbd5e1; font-weight: bold; text-align: center; margin-bottom: 5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 100%;">{n2}</div>
                    <div style="font-size: 28px; margin-bottom: -5px;">🥈</div>
                    <div style="background: linear-gradient(180deg, rgba(148,163,184,0.8), rgba(71,85,105,0.8)); backdrop-filter: blur(5px); width: 100%; height: 90px; border-radius: 12px 12px 0 0; display: flex; justify-content: center; align-items: flex-start; padding-top: 15px; color: white; font-weight: 900; font-size: 22px; box-shadow: 0 -5px 20px rgba(148,163,184,0.3); border: 1px solid rgba(255,255,255,0.3); border-bottom: none;">{t2}</div>
                </div>
                <!-- 1 Lugar -->
                <div style="display: flex; flex-direction: column; align-items: center; width: 140px; transform: translateY(-15px); transition: transform 0.3s;" onmouseover="this.style.transform='translateY(-20px)'" onmouseout="this.style.transform='translateY(-15px)'">
                    <div style="font-size: 15px; color: #fbbf24; font-weight: bold; text-align: center; margin-bottom: 5px; text-shadow: 0 0 10px rgba(251, 191, 36, 0.6); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 100%;">{n1}</div>
                    <div style="font-size: 38px; margin-bottom: -5px;">👑</div>
                    <div style="background: linear-gradient(180deg, rgba(251,191,36,0.9), rgba(180,83,9,0.9)); backdrop-filter: blur(5px); width: 100%; height: 130px; border-radius: 12px 12px 0 0; display: flex; justify-content: center; align-items: flex-start; padding-top: 15px; color: white; font-weight: 900; font-size: 26px; box-shadow: 0 -5px 25px rgba(251,191,36,0.5); border: 1px solid rgba(255,255,255,0.5); border-bottom: none;">{t1}</div>
                </div>
                <!-- 3 Lugar -->
                <div style="display: flex; flex-direction: column; align-items: center; width: 130px; transition: transform 0.3s;" onmouseover="this.style.transform='translateY(-5px)'" onmouseout="this.style.transform='translateY(0)'">
                    <div style="font-size: 13px; color: #d97706; font-weight: bold; text-align: center; margin-bottom: 5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 100%;">{n3}</div>
                    <div style="font-size: 28px; margin-bottom: -5px;">🥉</div>
                    <div style="background: linear-gradient(180deg, rgba(217,119,6,0.8), rgba(120,53,15,0.8)); backdrop-filter: blur(5px); width: 100%; height: 70px; border-radius: 12px 12px 0 0; display: flex; justify-content: center; align-items: flex-start; padding-top: 15px; color: white; font-weight: 900; font-size: 20px; box-shadow: 0 -5px 20px rgba(217,119,6,0.3); border: 1px solid rgba(255,255,255,0.2); border-bottom: none;">{t3}</div>
                </div>
            </div>
            """
            st.markdown(html_podio, unsafe_allow_html=True)
        else:
            for i, row in top3.iterrows():
                medalha = "🥇" if i == 0 else ("🥈" if i == 1 else "🥉")
                st.markdown(f"**{medalha} {row['ENCARREGADO']}** ({row['ENTREGAS']} RDCs)")
                
        st.markdown("---")
        st.markdown("#### 📈 Evolução Mensal")
        if not df_hist.empty and "MES_ANO" in df_hist.columns:
            df_evolucao = df_hist.groupby("MES_ANO").size().reset_index(name="RDCs Entregues")
        else:
            df_evolucao = pd.DataFrame()
        if not df_evolucao.empty:
            fig_ev = px.line(df_evolucao, x="MES_ANO", y="RDCs Entregues", text="RDCs Entregues", markers=True)
            fig_ev.update_traces(textposition="top center", line_color="#4a9eed", marker=dict(size=8))
            fig_ev.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), xaxis_title="Mês", yaxis_title="Total de RDCs")
            st.plotly_chart(fig_ev, use_container_width=True)
            
        st.markdown("---")
        if st.button("📄 Gerar Relatório Mensal em PDF", type="primary", use_container_width=True):
            try:
                from fpdf import FPDF
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(200, 10, txt=f"Relatorio Mensal F1 - {mes_selecionado}", ln=True, align='C')
                pdf.ln(10)
                
                pdf.set_font("Arial", 'B', 12)
                pdf.cell(200, 10, txt=f"Total de RDCs Entregues no Mes: {total_entregue}", ln=True)
                pdf.ln(10)
                
                pdf.set_font("Arial", 'B', 14)
                pdf.cell(200, 10, txt="Os 3 Melhores do Mes:", ln=True)
                pdf.set_font("Arial", '', 12)
                for i, row in top3.iterrows():
                    pdf.cell(200, 10, txt=f"{i+1} Lugar: {row['ENCARREGADO']} - {row['ENTREGAS']} RDCs", ln=True)
                pdf.ln(5)
                # Devedores Críticos
                if devedores:
                    pdf.set_font("Arial", 'B', 14)
                    pdf.set_text_color(255, 0, 0)
                    pdf.cell(200, 10, txt="Alerta Critico - Sem RDC a mais de 3 dias:", ln=True)
                    pdf.set_font("Arial", '', 12)
                    for d in devedores:
                        pdf.cell(200, 10, txt=f"- {d}", ln=True)
                
                pdf_output = bytes(pdf.output())
                st.download_button("📥 Clique aqui para baixar o PDF", data=pdf_output, file_name=f"Relatorio_{mes_selecionado}.pdf", mime="application/pdf", type="primary")
            except ImportError:
                st.error("Biblioteca FPDF não encontrada. Avise o desenvolvedor para instalar `fpdf2`.")
        
        st.markdown("<br><br>", unsafe_allow_html=True)

    if tab_ia is not None:
      with tab_ia:
        st.markdown("### 🤖 Robô de Extração Inteligente (Google Gemini)")
        st.markdown("<p style='margin-top: -15px; font-size: 14px; color: #888;'>Uma ideia original por <b>Caio Farisco</b></p>", unsafe_allow_html=True)
        st.markdown("Arraste os formulários RDC físicos escaneados abaixo. A inteligência artificial irá extrair as informações e padronizar com a sua base de Encarregados.")
        
        try:
            from google import genai
            HAS_GENAI = True
        except ImportError:
            HAS_GENAI = False
            st.error("A biblioteca `google-genai` não está instalada no servidor. Instale usando `pip install google-genai`.")

        if HAS_GENAI:
            # Tentar ler a chave do cofre secreto (.streamlit/secrets.toml)
            chave_padrao = ""
            try:
                chave_padrao = st.secrets.get("GEMINI_API_KEY", "")
            except Exception:
                pass
            
            st.markdown("#### Configuração e Upload")
            if not chave_padrao:
                chave_padrao = st.text_input("🔑 Cole suas Chaves da API Gemini (separadas por vírgula):", type="password", help="Se usar múltiplas chaves, o robô alterna em caso de limite.")
            
            if not chave_padrao:
                st.info("☝️ Cole a(s) sua(s) chave(s) de API do Gemini acima para ativar o robô de leitura.")
            
            arquivos_scan = st.file_uploader("Upload de RDCs Escaneados (PDF, JPG, PNG)", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
                
            btn_processar = st.button("🚀 Processar Arquivos com IA", type="primary", use_container_width=True)
            
            if btn_processar and arquivos_scan and chave_padrao:
                # --- FIX: Evitar que o Gemini tente usar o Service Account do Google Sheets ---
                old_cred = os.environ.pop("GOOGLE_APPLICATION_CREDENTIALS", None)
                
                lista_chaves = [c.strip() for c in chave_padrao.split(",") if c.strip()]
                idx_chave_atual = 0
                client = genai.Client(api_key=lista_chaves[idx_chave_atual])
                nomes_para_prompt = ", ".join(lista_encarregados_base)
                
                prompt_ia = f"""
                Analise este documento (que pode ter várias páginas). Para CADA formulário de obra (RDC) encontrado no arquivo, extraia os dados.
                REGRA IMPORTANTÍSSIMA: Retorne APENAS UM objeto JSON por formulário/página. NÃO separe as atividades em linhas diferentes. JUNTE TODAS as atividades do mesmo formulário em UM ÚNICO campo ATIVIDADE.
                DICA DE OURO: Todos os RDCs dentro deste arquivo PDF pertencem EXATAMENTE ao mesmo dia. Portanto, a DATA extraída deve ser idêntica para todos os formulários.
                Retorne APENAS um array (lista) em formato JSON válido. Exemplo do formato exato esperado:
                [
                  {{
                    "DATA": "YYYY-MM-DD",
                    "DISCIPLINA": "...",
                    "ENCARREGADO": "...",
                    "TURNO": "...",
                    "DDS": "...",
                    "ATIVIDADE": "...",
                    "PROBLEMAS": "...",
                    "LOCAL": "...",
                    "AREA": "..."
                  }}
                ]

                Regras de negócio:
                - DATA: Extraia a data em que o RDC foi preenchido. Retorne RIGOROSAMENTE no formato YYYY-MM-DD (Ano-Mês-Dia).
                - DISCIPLINA: Extraia a disciplina ou função do topo, mas RETORNE APENAS A PRIMEIRA PALAVRA OU A PALAVRA PRINCIPAL (ex: MECÂNICA, SOLDA, TOPOGRAFIA, CALDEIRARIA). Se for montador de andaime escreva ANDAIME. Sempre apenas 1 palavra.
                - ENCARREGADO: FAÇA O MÁXIMO ESFORÇO POSSÍVEL para descobrir quem é o encarregado. Compare o que está escrito à mão com esta lista oficial: [{nomes_para_prompt}]. Se a caligrafia estiver ruim, com erros de ortografia, ou se houver apenas o primeiro e segundo nome (ex: "Jailson Gois"), use dedução lógica e similaridade para encontrar a correspondência exata na lista. Retorne EXATAMENTE o nome completo que consta na lista fornecida. Somente se for 100% impossível deduzir quem é, retorne o texto 'AJUSTAR NOME'.
                - TURNO: Analise os horários. De dia (ex: 07:00 as 17:00) = 'DIURNO'. De noite = 'NOTURNO'.
                - DDS: Extraia o tema principal de Segurança mencionado no relatório (DDS, Diálogo de Segurança). (ex: Trabalho a quente, Bloqueio, etc). Se não tiver, retorne 'Não Informado'.
                - ATIVIDADE: OBRIGATÓRIO: Crie um ÚNICO RESUMO SUPER CURTO E DIRETO de NO MÁXIMO 20 PALAVRAS sobre o que foi feito na seção 'ATIVIDADES'. Se você usar mais de 20 palavras, será considerado um erro gravíssimo! Extraia a ação principal, corrija a ortografia e escreva TUDO EM MAIÚSCULAS. NUNCA SEPARE EM LINHAS.
                - CALDEIRA: Se mencionar 'caldeira de recuperação' = 'RB'. Se 'caldeira de potência' = 'PB'. Se a descrição da atividade mencionar 'PRECIPITADOR' ou 'ESP' = 'ESP'. Se nenhum = ''.
                - LOCAL: Analise a imagem CUIDADOSAMENTE. Procure as opções 'PB ( )' e 'RB ( )'. Verifique se há um 'X', um rabisco, um visto ou qualquer marcação (mesmo que mal desenhada) dentro, em cima ou do lado dos parênteses. Retorne APENAS 'PB' ou 'RB' correspondente ao que estiver marcado. Se nenhum, retorne ''.
                - AREA: Analise as caixinhas de área na imagem com LUPA. Procure por qualquer marcação (X, visto, círculo, rabisco) dentro ou sobre os parênteses. As opções são exatamente: DUTO, EQUIPAMENTO, TUBULAÇÃO, ESTRUTURA MET, PRECIPITADOR, PRESSAO - MEC, PRESSAO - TUBULACAO, PRESSAO - FORNALHA, PINTURA, SOPRAGEM, ANDAIME. Retorne EXATAMENTE o nome da área que estiver marcada. Se nenhuma estiver marcada, retorne ''.

                Não inclua crases, formatação markdown ou texto adicional, apenas o JSON puro começando com [ e terminando com ].
                """

                # === ANIMAÇÃO PREMIUM DE LOADING ===
                animacao_html = """
                <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 25px; background: rgba(15,23,42,0.9); border: 1px solid #0ea5e9; border-radius: 15px; box-shadow: 0 0 30px rgba(14, 165, 233, 0.3); margin-bottom: 20px;">
                    <div class="radar" style="position: relative; width: 120px; height: 120px; border-radius: 50%; border: 2px solid rgba(14,165,233,0.5); overflow: hidden; background: radial-gradient(circle, rgba(14,165,233,0.15) 0%, rgba(15,23,42,0) 100%);">
                        <div style="position: absolute; width: 50%; height: 50%; top: 0; left: 50%; transform-origin: bottom left; background: linear-gradient(45deg, rgba(14,165,233,0.9) 0%, transparent 50%); animation: radar-spin 1.5s linear infinite;"></div>
                        <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); color: #fff; font-weight: bold; font-size: 14px; letter-spacing: 2px; text-shadow: 0 0 10px #0ea5e9; background: #0f172a; padding: 5px; border-radius: 5px;">ENESA</div>
                        <div style="position: absolute; top: 0; bottom: 0; left: 50%; width: 1px; background: rgba(14,165,233,0.4);"></div>
                        <div style="position: absolute; left: 0; right: 0; top: 50%; height: 1px; background: rgba(14,165,233,0.4);"></div>
                        <div style="position: absolute; top: 20%; left: 20%; width: 6px; height: 6px; background: #4ade80; border-radius: 50%; box-shadow: 0 0 10px #4ade80; animation: blip 1.5s infinite;"></div>
                        <div style="position: absolute; top: 70%; left: 60%; width: 4px; height: 4px; background: #4ade80; border-radius: 50%; box-shadow: 0 0 10px #4ade80; animation: blip 1.5s infinite 0.7s;"></div>
                    </div>
                    <p style="color: #0ea5e9; margin-top: 20px; font-weight: bold; font-size: 16px; animation: pulse 1s infinite; margin-bottom: 0;">🤖 IA Processando Documentos...</p>
                    <style>
                        @keyframes radar-spin { 100% { transform: rotate(360deg); } }
                        @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.6; } }
                        @keyframes blip { 0%, 100% { opacity: 0; } 10% { opacity: 1; } }
                    </style>
                </div>
                """
                animacao_placeholder = st.empty()
                animacao_placeholder.markdown(animacao_html, unsafe_allow_html=True)
                
                with st.status("🤖 Robô iniciando análise...", expanded=True) as status:
                    progresso = st.progress(0)
                    total_arquivos = len(arquivos_scan)
                    
                    for i, arquivo_scan in enumerate(arquivos_scan):
                        status.update(label=f"Processando arquivo {i+1} de {total_arquivos}: {arquivo_scan.name}...", state="running")
                    
                        try:
                            # Salvar temporariamente para enviar pro Gemini
                            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{arquivo_scan.name.split('.')[-1]}") as tmp:
                                tmp.write(arquivo_scan.getvalue())
                                tmp_path = tmp.name
                            
                            arquivo_up = client.files.upload(file=tmp_path)
                        
                            max_tentativas = 3
                            sucesso_arquivo = False
                            for tentativa in range(max_tentativas):
                                try:
                                    resposta = client.models.generate_content(
                                        model='gemini-3.5-flash',
                                        contents=[arquivo_up, prompt_ia],
                                        config=genai.types.GenerateContentConfig(
                                            response_mime_type="application/json",
                                            temperature=0.0
                                        )
                                    )
                                
                                    # --- FIX: Restaurar as credenciais do Sheets caso necessário ---
                                    if old_cred:
                                        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = old_cred

                                    texto_resposta = resposta.text.strip()
                                    if texto_resposta.startswith("```json"):
                                        texto_resposta = texto_resposta[7:-3].strip()
                                    elif texto_resposta.startswith("```"):
                                        texto_resposta = texto_resposta[3:-3].strip()

                                    dados_extraidos_lista = json.loads(texto_resposta)

                                    if isinstance(dados_extraidos_lista, dict):
                                        dados_extraidos_lista = [dados_extraidos_lista]

                                    # --- FIX: Consenso de Data do Lote ---
                                    # Como todos os RDCs escaneados juntos pertencem ao mesmo dia,
                                    # pegamos a data mais frequente encontrada e forçamos para todos.
                                    datas_encontradas = [str(d.get("DATA")).strip() for d in dados_extraidos_lista if d.get("DATA") and str(d.get("DATA")).strip() != ""]
                                    if datas_encontradas:
                                        data_consenso = max(set(datas_encontradas), key=datas_encontradas.count)
                                        for d in dados_extraidos_lista:
                                            d["DATA"] = data_consenso

                                    for dados in dados_extraidos_lista:
                                        ultimo_item = st.session_state.df_ia['ITEM'].max() if not st.session_state.df_ia.empty and pd.notna(st.session_state.df_ia['ITEM'].max()) else 0
                                        dados['ITEM'] = int(ultimo_item) + 1
                                        if 'LOCAL' not in dados:
                                            dados['LOCAL'] = ''
                                        if 'AREA' not in dados:
                                            dados['AREA'] = ''
                                        st.session_state.df_ia = pd.concat([st.session_state.df_ia, pd.DataFrame([dados])], ignore_index=True)
                                    
                                        # Apenas extrai os dados, aguardando validação do usuário.

                                    sucesso_arquivo = True
                                    break 

                                except Exception as inner_e:
                                    erro_str = str(inner_e)
                                    if '429' in erro_str or 'RESOURCE_EXHAUSTED' in erro_str:
                                        if tentativa < max_tentativas - 1:
                                            if idx_chave_atual < len(lista_chaves) - 1:
                                                idx_chave_atual += 1
                                                client = genai.Client(api_key=lista_chaves[idx_chave_atual])
                                                st.warning(f"🔄 Limite atingido na chave atual. Trocando para a chave reserva {idx_chave_atual + 1}/{len(lista_chaves)}...")
                                                time.sleep(2)
                                                continue
                                            else:
                                                st.warning(f"⏳ Cota do Google atingida em todas as chaves. Aguardando 60 segundos... (Tentativa {tentativa+1}/{max_tentativas})")
                                                time.sleep(60)
                                                continue
                                            
                                    msg_erro = f"Erro detalhado na IA: {inner_e}"
                                    try:
                                        # Tentar buscar a lista de modelos para debug
                                        modelos = [m.name for m in client.models.list()]
                                        msg_erro += f" | Modelos liberados na sua chave: {modelos}"
                                    except:
                                        pass
                                    st.error(msg_erro)
                                    if '503' in str(inner_e):
                                        time.sleep(10)
                                    else:
                                        break
                                    
                            os.remove(tmp_path)
                        
                            if sucesso_arquivo:
                                st.toast(f"✅ {arquivo_scan.name} processado com sucesso!")
                            else:
                                st.toast(f"❌ Falha ao processar {arquivo_scan.name}.")
                                # Se falhou pelo menos um, mantem expandido
                                st.session_state.teve_falha_ia = True
                            
                        except Exception as e:
                            st.error(f"Erro no envio do arquivo {arquivo_scan.name}: {e}")
                            st.session_state.teve_falha_ia = True
                        
                        progresso.progress((i + 1) / total_arquivos)

                    expandir_status = st.session_state.get('teve_falha_ia', False)
                    status.update(label="🎉 Leitura concluída!" if not expandir_status else "⚠️ Leitura finalizada com erros", state="complete", expanded=expandir_status)
                    st.session_state.teve_falha_ia = False
                    
                    pass
                animacao_placeholder.empty()
                st.session_state.force_use_local = True
                
            if not st.session_state.df_ia.empty:
                st.markdown("#### Dados Extraídos")
                
                lista_com_alerta = lista_encarregados_base + ["AJUSTAR NOME"]
                df_filtrado = st.session_state.df_ia[st.session_state.df_ia['ENCARREGADO'].isin(lista_com_alerta)]
                
                # --- NOVO FILTRO DE DATA ---
                datas_disponiveis = df_filtrado['DATA'].dropna().unique().tolist()
                if datas_disponiveis:
                    datas_sel = st.multiselect("📅 Filtrar Tabela por Data (Deixe em branco para ver todos):", sorted(datas_disponiveis, reverse=True), default=None)
                    if datas_sel:
                        df_filtrado = df_filtrado[df_filtrado['DATA'].isin(datas_sel)]
                # ---------------------------
                
                col_dw1, col_dw2 = st.columns([1, 1])
                with col_dw1:
                    buffer_df = io.BytesIO()
                    df_filtrado.to_excel(buffer_df, index=False, engine='openpyxl')
                    buffer_df.seek(0)
                    st.download_button(
                        label="⬇️ Baixar Planilha RDC Lida (.xlsx)",
                        data=buffer_df,
                        file_name=f"RDCs_Extraidos_{datetime.datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                with col_dw2:
                    if st.button("🗑️ Limpar Dados Lidos", use_container_width=True):
                        st.session_state.df_ia = pd.DataFrame(columns=['ITEM', 'DATA', 'DISCIPLINA', 'ENCARREGADO', 'TURNO', 'DDS', 'ATIVIDADE', 'PROBLEMAS', 'LOCAL', 'AREA'])
                        st.rerun()
                
                st.info("✏️ **Dica:** Você pode editar os dados na tabela abaixo antes de confirmar. Dê dois cliques em qualquer célula para corrigir nomes errados, datas ou locais.")
                df_editado = st.data_editor(df_filtrado, hide_index=True, use_container_width=True, key="editor_ia_df")
                
                if st.button("✅ Confirmar e Salvar no Sistema", type="primary", use_container_width=True):
                    # Salva no df_ia
                    st.session_state.df_ia = df_editado
                    
                    # Salva no F1
                    novos_registros = []
                    for _, row in df_editado.iterrows():
                        enc_lido = str(row.get('ENCARREGADO', '')).strip()
                        if enc_lido and enc_lido in lista_encarregados_base:
                            data_extraida = str(row.get('DATA', '')).strip()
                            try:
                                data_registro = pd.to_datetime(data_extraida).strftime('%Y-%m-%d')
                            except:
                                data_registro = datetime.date.today().strftime('%Y-%m-%d')
                            
                            ja_existe = ((st.session_state.df_historico_f1["DATA"] == data_registro) & (st.session_state.df_historico_f1["ENCARREGADO"] == enc_lido)).any()
                            if not ja_existe:
                                novos_registros.append({"DATA": data_registro, "ENCARREGADO": enc_lido})
                                
                    if novos_registros:
                        df_novos = pd.DataFrame(novos_registros)
                        if conn and not st.session_state.get('force_use_local', False):
                            try:
                                df_fresco = conn.read(worksheet="Historico_F1", ttl=0)
                                if not df_fresco.empty:
                                    df_fresco = df_fresco.dropna(how='all')
                                    df_final = pd.concat([df_fresco, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                                else:
                                    df_final = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                                
                                conn.update(worksheet="Historico_F1", data=df_final)
                                st.session_state.df_historico_f1 = df_final
                                st.cache_data.clear()
                                st.success(f"✅ {len(novos_registros)} RDCs registrados no Resumo Diário e sincronizados com a nuvem!")
                            except Exception as e:
                                st.error(f"Erro ao salvar na nuvem: {e}")
                                st.session_state.df_historico_f1 = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                        else:
                            st.session_state.df_historico_f1 = pd.concat([st.session_state.df_historico_f1, df_novos], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                            st.success(f"✅ {len(novos_registros)} RDCs registrados localmente no Resumo Diário!")
                    else:
                        st.info("ℹ️ Os dados foram processados, mas os Encarregados dessa lista já haviam sido contabilizados.")

    if tab_ia_cc is not None:
      with tab_ia_cc:
        st.markdown("### Robô Atualizador de C.C (Google Gemini)")
        st.markdown("Faça o upload dos PDFs aqui para o robô identificar o Local (PB/RB) e a Área (Estrutura, Tubulação, etc) e atualizar automaticamente o C.C. das equipes na base global do Google Sheets.")
        
        if HAS_GENAI:
            chave_padrao = ""
            try:
                chave_padrao = st.secrets.get("GEMINI_API_KEY", "")
            except Exception:
                pass
            
            st.markdown("#### Configuração e Upload")
            if not chave_padrao:
                chave_padrao = st.text_input("🔑 Cole suas Chaves da API Gemini (separadas por vírgula):", type="password", help="Chave oculta e protegida.", key="chave_cc")
            
            arquivos_scan_cc = st.file_uploader("Upload de RDCs para atualização de C.C (PDF, JPG, PNG)", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True, key="uploader_cc")
                
            btn_processar_cc = st.button("🚀 Atualizar C.C das Equipes com IA", type="primary", use_container_width=True)
            
            if btn_processar_cc and arquivos_scan_cc and chave_padrao:
                old_cred = os.environ.pop("GOOGLE_APPLICATION_CREDENTIALS", None)
                lista_chaves = [c.strip() for c in chave_padrao.split(",") if c.strip()]
                idx_chave_atual = 0
                client = genai.Client(api_key=lista_chaves[idx_chave_atual])
                nomes_para_prompt = ", ".join(lista_encarregados_base)
                
                prompt_ia_cc = f"""
                Analise este documento. Para CADA formulário de obra (RDC) encontrado no arquivo, extraia os dados.
                REGRA IMPORTANTÍSSIMA: Retorne APENAS UM objeto JSON por formulário/página.
                Retorne APENAS um array (lista) em formato JSON válido.
                [
                  {{
                    "DISCIPLINA": "...",
                    "ENCARREGADO": "...",
                    "PROBLEMAS": "...",
                    "LOCAL": "...",
                    "AREA": "..."
                  }}
                ]

                Regras de negócio:
                - DISCIPLINA: Extraia a disciplina ou função do topo, mas RETORNE APENAS A PRIMEIRA PALAVRA OU A PALAVRA PRINCIPAL.
                - ENCARREGADO: Extraia o nome do Encarregado escrito no papel. Compare com: [{nomes_para_prompt}]. Retorne EXATAMENTE o nome correspondente. Se ilegível, retorne 'AJUSTAR NOME'.
                - CALDEIRA: Se mencionar 'caldeira de recuperação' = 'RB'. Se 'caldeira de potência' = 'PB'. Se 'PRECIPITADOR' ou 'ESP' = 'ESP'. Se nenhum = ''.
                - LOCAL: Analise a imagem CUIDADOSAMENTE. Procure as opções 'PB ( )' e 'RB ( )'. Verifique se há um 'X', rabisco, visto ou marcação (mesmo que mal desenhada) dentro, em cima ou do lado dos parênteses. Retorne APENAS 'PB' ou 'RB'. Se nenhum, retorne ''.
                - AREA: Analise as caixinhas de área na imagem com LUPA. Procure por qualquer marcação (X, visto, círculo, rabisco) dentro ou sobre os parênteses. Opções: DUTO, EQUIPAMENTO, TUBULAÇÃO, ESTRUTURA MET, PRECIPITADOR, PRESSAO - MEC, PRESSAO - TUBULACAO, PRESSAO - FORNALHA, PINTURA, SOPRAGEM, ANDAIME. Retorne EXATAMENTE a área marcada. Se nenhuma, retorne ''.

                Apenas o JSON puro começando com [ e terminando com ].
                """

                # === ANIMAÇÃO PREMIUM DE LOADING ===
                animacao_html = """
                <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 25px; background: rgba(15,23,42,0.9); border: 1px solid #0ea5e9; border-radius: 15px; box-shadow: 0 0 30px rgba(14, 165, 233, 0.3); margin-bottom: 20px;">
                    <div class="radar" style="position: relative; width: 120px; height: 120px; border-radius: 50%; border: 2px solid rgba(14,165,233,0.5); overflow: hidden; background: radial-gradient(circle, rgba(14,165,233,0.15) 0%, rgba(15,23,42,0) 100%);">
                        <div style="position: absolute; width: 50%; height: 50%; top: 0; left: 50%; transform-origin: bottom left; background: linear-gradient(45deg, rgba(14,165,233,0.9) 0%, transparent 50%); animation: radar-spin 1.5s linear infinite;"></div>
                        <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); color: #fff; font-weight: bold; font-size: 14px; letter-spacing: 2px; text-shadow: 0 0 10px #0ea5e9; background: #0f172a; padding: 5px; border-radius: 5px;">ENESA</div>
                        <div style="position: absolute; top: 0; bottom: 0; left: 50%; width: 1px; background: rgba(14,165,233,0.4);"></div>
                        <div style="position: absolute; left: 0; right: 0; top: 50%; height: 1px; background: rgba(14,165,233,0.4);"></div>
                        <div style="position: absolute; top: 20%; left: 20%; width: 6px; height: 6px; background: #4ade80; border-radius: 50%; box-shadow: 0 0 10px #4ade80; animation: blip 1.5s infinite;"></div>
                        <div style="position: absolute; top: 70%; left: 60%; width: 4px; height: 4px; background: #4ade80; border-radius: 50%; box-shadow: 0 0 10px #4ade80; animation: blip 1.5s infinite 0.7s;"></div>
                    </div>
                    <p style="color: #0ea5e9; margin-top: 20px; font-weight: bold; font-size: 16px; animation: pulse 1s infinite; margin-bottom: 0;">🤖 IA Atualizando Centros de Custo...</p>
                    <style>
                        @keyframes radar-spin { 100% { transform: rotate(360deg); } }
                        @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.6; } }
                        @keyframes blip { 0%, 100% { opacity: 0; } 10% { opacity: 1; } }
                    </style>
                </div>
                """
                animacao_placeholder_cc = st.empty()
                animacao_placeholder_cc.markdown(animacao_html, unsafe_allow_html=True)

                with st.status("🤖 Atualizando C.C das equipes...", expanded=True) as status_cc:
                    progresso = st.progress(0)
                    total_arquivos = len(arquivos_scan_cc)
                    houve_atualizacao_global = False
                    
                    for i, arquivo_scan in enumerate(arquivos_scan_cc):
                        status_cc.update(label=f"Processando arquivo {i+1} de {total_arquivos}: {arquivo_scan.name}...", state="running")
                    
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{arquivo_scan.name.split('.')[-1]}") as tmp:
                            tmp.write(arquivo_scan.getvalue())
                            tmp_path = tmp.name
                            
                        arquivo_up = client.files.upload(file=tmp_path)
                        
                        max_tentativas = 3
                        sucesso_arquivo = False
                        for tentativa in range(max_tentativas):
                            try:
                                resposta = client.models.generate_content(
                                    model='gemini-3.5-flash',
                                    contents=[arquivo_up, prompt_ia_cc]
                                )
                                
                                if old_cred:
                                    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = old_cred

                                texto_json = resposta.text.strip()
                                if texto_json.startswith("```json"):
                                    texto_json = texto_json[7:-3].strip()
                                elif texto_json.startswith("```"):
                                    texto_json = texto_json[3:-3].strip()
                                
                                dados_extraidos_lista = json.loads(texto_json)
                                if isinstance(dados_extraidos_lista, dict):
                                    dados_extraidos_lista = [dados_extraidos_lista]
                                    
                                for dados in dados_extraidos_lista:
                                    # === ATUALIZAR C.C. COMPLETO NA BASE ===
                                    local_bruto = str(dados.get('LOCAL', '')).strip().upper()
                                    area_bruta = str(dados.get('AREA', '')).strip().upper()
                                    disciplina_lida = str(dados.get('DISCIPLINA', '')).strip().upper()
                                    enc_lido = str(dados.get('ENCARREGADO', '')).strip().upper()
                                    
                                    local_lido = ''
                                    if 'PB' in local_bruto: local_lido = 'PB'
                                    elif 'RB' in local_bruto: local_lido = 'RB'
                                    if not local_lido:
                                        cald = str(dados.get('CALDEIRA', '')).strip().upper()
                                        if 'PB' in cald: local_lido = 'PB'
                                        elif 'RB' in cald: local_lido = 'RB'
                                    
                                    area_lida = ''
                                    # 1. Tenta achar na área bruta (exato ou contendo)
                                    chaves_ordenadas = sorted(mapa_area_sufixo.keys(), key=len, reverse=True)
                                    for k in chaves_ordenadas:
                                        if k in area_bruta:
                                            area_lida = k
                                            break
                                            
                                    # 2. Se não achar, procura na disciplina (cuidado com falsos positivos de 'ESP')
                                    if not area_lida:
                                        import re
                                        for k in chaves_ordenadas:
                                            if k == 'ESP':
                                                if re.search(r'\bESP\b', disciplina_lida):
                                                    area_lida = k
                                                    break
                                            elif k in disciplina_lida:
                                                area_lida = k
                                                break
                                            
                                    if enc_lido and enc_lido != 'AJUSTAR NOME' and 'C.C' in df_atual.columns:
                                        encarregados_unicos = df_atual['ENCARREGADO'].dropna().unique()
                                        enc_encontrado = None
                                        
                                        for e in encarregados_unicos:
                                            if str(e).strip().upper() == enc_lido:
                                                enc_encontrado = e
                                                break
                                        if not enc_encontrado:
                                            for e in encarregados_unicos:
                                                if enc_lido in str(e).upper():
                                                    enc_encontrado = e
                                                    break
                                        if not enc_encontrado:
                                            import difflib
                                            matches = difflib.get_close_matches(enc_lido, [str(e).upper() for e in encarregados_unicos], n=1, cutoff=0.6)
                                            if matches:
                                                for e in encarregados_unicos:
                                                    if str(e).upper() == matches[0]:
                                                        enc_encontrado = e
                                                        break
                                                        
                                        if enc_encontrado:
                                            mask_enc = df_atual['ENCARREGADO'] == enc_encontrado
                                            atualizado = False
                                            
                                            if local_lido in ['PB', 'RB']:
                                                prefixo_novo = '125.02' if local_lido == 'PB' else '125.01'
                                                sufixo = mapa_area_sufixo.get(area_lida, '')
                                                
                                                if sufixo:
                                                    cc_novo = f"{prefixo_novo}.{sufixo}"
                                                    df_atual.loc[mask_enc, 'C.C'] = cc_novo
                                                    atualizado = True
                                                    st.toast(f"✅ C.C. de TODA A EQUIPE de {enc_encontrado} → {cc_novo}")
                                                else:
                                                    if local_lido == 'PB':
                                                        df_atual.loc[mask_enc, 'C.C'] = df_atual.loc[mask_enc, 'C.C'].str.replace('125.01.', '125.02.', regex=False)
                                                    else:
                                                        df_atual.loc[mask_enc, 'C.C'] = df_atual.loc[mask_enc, 'C.C'].str.replace('125.02.', '125.01.', regex=False)
                                                    atualizado = True
                                                    st.toast(f"⚠️ C.C. de TODA A EQUIPE de {enc_encontrado} atualizado parcialmente → {local_lido} (manteve sufixo)")
                                            else:
                                                st.warning(f"❌ C.C não atualizado para a equipe de {enc_encontrado}: O robô não conseguiu identificar se o local era PB ou RB.")
                                            
                                            if atualizado:
                                                st.session_state.df = df_atual.copy()
                                                houve_atualizacao_global = True
                                        else:
                                            st.error(f"❌ Encarregado '{enc_lido}' não encontrado na base. Equipe não atualizada.")

                                sucesso_arquivo = True
                                break 

                            except Exception as inner_e:
                                erro_str = str(inner_e)
                                if '429' in erro_str or 'RESOURCE_EXHAUSTED' in erro_str:
                                    if tentativa < max_tentativas - 1:
                                        if idx_chave_atual < len(lista_chaves) - 1:
                                            idx_chave_atual += 1
                                            client = genai.Client(api_key=lista_chaves[idx_chave_atual])
                                            st.warning(f"🔄 Limite atingido na chave atual. Trocando para a chave reserva {idx_chave_atual + 1}/{len(lista_chaves)}...")
                                            time.sleep(2)
                                            continue
                                        else:
                                            st.warning(f"⏳ Cota do Google atingida em todas as chaves. Aguardando 60 segundos... (Tentativa {tentativa+1}/{max_tentativas})")
                                            time.sleep(60)
                                            continue
                                        
                                msg_erro = f"Erro detalhado na IA: {inner_e}"
                                try:
                                    modelos = [m.name for m in client.models.list()]
                                    msg_erro += f" | Modelos liberados na sua chave: {modelos}"
                                except:
                                    pass
                                st.error(msg_erro)
                                if '503' in str(inner_e):
                                    time.sleep(10)
                                else:
                                    break
                                    
                        os.remove(tmp_path)
                        
                        if sucesso_arquivo:
                            st.toast(f"✅ {arquivo_scan.name} processado com sucesso!")
                        else:
                            st.toast(f"❌ Falha ao processar {arquivo_scan.name}.")
                            st.session_state.teve_falha_ia_cc = True
                            
                    except Exception as e:
                        st.error(f"Erro no envio do arquivo {arquivo_scan.name}: {e}")
                        st.session_state.teve_falha_ia_cc = True
                        
                    progresso.progress((i + 1) / total_arquivos)

                expandir_status = st.session_state.get('teve_falha_ia_cc', False)
                status_cc.update(label="✅ Atualização de C.Cs concluída!" if not expandir_status else "⚠️ Leitura finalizada com erros", state="complete", expanded=expandir_status)
                animacao_placeholder_cc.empty()
                st.session_state.teve_falha_ia_cc = False
                
                if houve_atualizacao_global:
                    try:
                        df_atual = preparar_dataframe(df_atual)
                        st.session_state.df = df_atual.copy()
                        
                        status_cc.update(label="Sincronizando C.Cs atualizados com a nuvem...", state="running")
                        conn_update = st.connection("gsheets", type=GSheetsConnection)
                        conn_update.update(worksheet="Página1", data=df_atual)
                        st.cache_data.clear()
                    except Exception as e:
                        st.error(f"Erro ao salvar na nuvem: {e}")

                status_cc.update(label="🎉 Atualização de C.Cs concluída!", state="complete", expanded=False)
                time.sleep(2)
                st.session_state.force_use_local = True
                st.rerun()
                
                st.dataframe(df_filtrado, use_container_width=True)

    with tab_cc:
        st.markdown("### 💰 Controle de Centro de Custo (C.C)")
        
        # === ÚLTIMA ATUALIZAÇÃO ===
        ultima_base = ""
        ultima_cc = ""
        try:
            if os.path.exists(caminho_base_salva_csv):
                ts_base = os.path.getmtime(caminho_base_salva_csv)
                ultima_base = datetime.datetime.fromtimestamp(ts_base).strftime("%d/%m/%Y às %H:%M")
            elif os.path.exists(caminho_base_salva_xlsx):
                ts_base = os.path.getmtime(caminho_base_salva_xlsx)
                ultima_base = datetime.datetime.fromtimestamp(ts_base).strftime("%d/%m/%Y às %H:%M")
        except Exception:
            pass
        try:
            if os.path.exists(caminho_hist_cc):
                ts_cc = os.path.getmtime(caminho_hist_cc)
                ultima_cc = datetime.datetime.fromtimestamp(ts_cc).strftime("%d/%m/%Y às %H:%M")
        except Exception:
            pass
        
        html_update = f"""
        <div style="display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap;">
            <div style="background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.3); border-radius: 10px; padding: 10px 18px; display: flex; align-items: center; gap: 10px;">
                <div style="width: 8px; height: 8px; border-radius: 50%; background: #10b981; box-shadow: 0 0 8px #10b981; animation: pulse_dot 2s infinite;"></div>
                <span style="font-size: 13px; color: #94a3b8;">Base PDE atualizada em: <b style="color: #10b981;">{ultima_base if ultima_base else 'N/A'}</b></span>
            </div>
            <div style="background: rgba(139, 92, 246, 0.1); border: 1px solid rgba(139, 92, 246, 0.3); border-radius: 10px; padding: 10px 18px; display: flex; align-items: center; gap: 10px;">
                <div style="width: 8px; height: 8px; border-radius: 50%; background: #8b5cf6; box-shadow: 0 0 8px #8b5cf6; animation: pulse_dot 2s infinite;"></div>
                <span style="font-size: 13px; color: #94a3b8;">Histórico C.C salvo em: <b style="color: #8b5cf6;">{ultima_cc if ultima_cc else 'N/A'}</b></span>
            </div>
        </div>
        <style>
            @keyframes pulse_dot {{
                0%, 100% {{ opacity: 1; }}
                50% {{ opacity: 0.3; }}
            }}
        </style>
        """
        st.markdown(html_update, unsafe_allow_html=True)
        
        if "C.C" not in df_atual.columns or df_atual["C.C"].str.strip().eq("").all():
            st.warning("⚠️ A coluna de Centro de Custo (C.C) não foi encontrada na base de dados atual. Verifique se a planilha possui essa coluna.")
        else:
            # === ALERTA DE C.C INVÁLIDO ===
            df_em_branco = df_atual[df_atual["C.C"].isna() | df_atual["C.C"].str.strip().eq("")]
            if not df_em_branco.empty:
                st.error(f"⚠️ **ALERTA DE SISTEMA:** Existem **{len(df_em_branco)} colaboradores** na base atual **sem Centro de Custo** (C.C em branco). Eles não aparecerão nos cálculos de custo!")
            
            valid_prefixes = ["125.01.", "125.02."]
            valid_suffixes = ['001', '002', '003', '004', '005', '006', '007', '008', '009', '010', '011', '012', '013', '014', '015', '016', '101', '102', '103', '104', '105', '106', '107', '108', '109', '110', '111', '112', '113']
            
            invalid_cc_list = []
            df_preenchido = df_atual[~df_atual["C.C"].isna() & (df_atual["C.C"].str.strip() != "")]
            for _, row in df_preenchido.iterrows():
                cc_val = str(row["C.C"]).strip()
                is_valid = False
                for prefix in valid_prefixes:
                    if cc_val.startswith(prefix):
                        suf = cc_val.replace(prefix, "")
                        if suf in valid_suffixes:
                            is_valid = True
                            break
                if not is_valid:
                    invalid_cc_list.append(cc_val)
                    
            if invalid_cc_list:
                invalid_cc_count = len(invalid_cc_list)
                unique_invalids = list(set(invalid_cc_list))
                st.warning(f"⚠️ **ATENÇÃO:** Foram encontrados **{invalid_cc_count} colaboradores** com C.C **inválido** (não existe no mapa oficial). Exemplos: {', '.join(unique_invalids[:5])}")

            # Filtro PB/RB/ESP Global para a aba C.C
            filtro_local = st.segmented_control(
                "Filtrar Dados por Local:", 
                ["Ambas", "PB", "RB", "ESP"], 
                default="Ambas",
                key="filtro_cc_local_key"
            )
            if not filtro_local:
                filtro_local = "Ambas"
                
            df_cc_aba = df_atual[df_atual["C.C"].str.strip() != ""]
            if filtro_local == "PB":
                df_cc_aba = df_cc_aba[df_cc_aba["C.C"].apply(lambda x: "125.02" in str(x) and ".005" not in str(x))]
            elif filtro_local == "RB":
                df_cc_aba = df_cc_aba[df_cc_aba["C.C"].apply(lambda x: "125.01" in str(x) and ".005" not in str(x))]
            elif filtro_local == "ESP":
                df_cc_aba = df_cc_aba[df_cc_aba["C.C"].apply(lambda x: ".005" in str(x))]

            lista_cc = sorted([str(cc) for cc in df_cc_aba["C.C"].unique()])
            
            def format_cc(cc_code):
                if cc_code == "TODOS": return "TODOS"
                local = "PB" if "125.02" in cc_code else ("RB" if "125.01" in cc_code else "")
                sufixo = str(cc_code).split('.')[-1] if '.' in str(cc_code) else str(cc_code)
                mapa_sufixo_nome = {
                    '001': 'Equipamentos', '002': 'Dutos', '003': 'Tubulação', 
                    '004': 'Estrutura Metálica', '005': 'Precipitador', '006': 'Pressão - Mecânica', 
                    '007': 'Pressão - Tubulação', '008': 'Pressão - Fornalha', '009': 'Pintura', 
                    '010': 'Comissionamento', '011': 'Op. Assistida', '012': 'Lavagem Química', 
                    '013': 'Sopragem', '014': 'Andaime', '015': 'Operadores', '016': 'Fora de Escopo',
                    '101': 'Gerência', '102': 'Produção', '103': 'Garantia da Qualidade',
                    '104': 'Planejamento', '105': 'Administração', '106': 'Segurança e Medicina',
                    '107': 'Infraestrutura', '108': 'Almoxarifado ENESA', '109': 'Almoxarifado Materiais',
                    '110': 'Manut. Elétrica Provisória', '111': 'Topografia', '112': 'Movimentação de Cargas',
                    '113': 'Medição/Contratos'
                }
                nome = mapa_sufixo_nome.get(sufixo, '')
                
                if nome and local: return f"{cc_code} - {nome} ({local})"
                elif nome: return f"{cc_code} - {nome}"
                elif local: return f"{cc_code} ({local})"
                else: return str(cc_code)
            
            # Métricas gerais Customizadas
            def card_kpi_cc(titulo, valor, cor):
                return f"""
                <div style="background: rgba(30, 41, 59, 0.45); backdrop-filter: blur(10px); border-radius: 16px; border: 1px solid rgba(255,255,255,0.05); padding: 18px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); position: relative; overflow: hidden; height: 110px; transition: transform 0.3s ease;" onmouseover="this.style.transform='translateY(-5px)'" onmouseout="this.style.transform='translateY(0px)'">
                    <p style="margin: 0; font-size: 13px; color: #94a3b8; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;">{titulo}</p>
                    <h2 style="margin: 5px 0 0 0; font-size: 34px; font-weight: 700; color: #f8fafc; text-shadow: 0 0 15px {cor}60;">{valor}</h2>
                    <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(90deg, {cor}, transparent); box-shadow: 0 -2px 10px {cor}80;"></div>
                </div>
                """
                
            mc1, mc2, mc3, mc4, mc5 = st.columns(5)
            with mc1: st.markdown(card_kpi_cc("Centros de Custo", len(lista_cc), "#8b5cf6"), unsafe_allow_html=True)
            with mc2: st.markdown(card_kpi_cc("Total Alocados", len(df_cc_aba), "#3b82f6"), unsafe_allow_html=True)
            with mc3: st.markdown(card_kpi_cc("Funções Distintas", df_cc_aba["FUNÇÃO"].nunique(), "#f59e0b"), unsafe_allow_html=True)
            
            qtd_encarregados = len([e for e in df_cc_aba["ENCARREGADO"].unique() if str(e).strip() != "" and str(e) in lista_completa_encarregados])
            with mc4: st.markdown(card_kpi_cc("Encarregados", qtd_encarregados, "#10b981"), unsafe_allow_html=True)
            
            span_of_control = round(len(df_cc_aba) / qtd_encarregados, 1) if qtd_encarregados > 0 else 0
            with mc5: st.markdown(card_kpi_cc("Span of Control", span_of_control, "#0ea5e9"), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

            # Gráficos lado a lado
            col_graf1, col_graf2 = st.columns([6, 4])
            
            with col_graf1:
                # Gráfico de distribuição por C.C.
                st.markdown("**Distribuição de Efetivo por Centro de Custo**")
                
                cc_contagem = df_cc_aba["C.C"].value_counts().reset_index()
                cc_contagem.columns = ["Centro de Custo", "Quantidade"]
                cc_contagem["Nome C.C"] = cc_contagem["Centro de Custo"].apply(format_cc)
                
                if len(cc_contagem) > 0:
                    fig_cc = px.bar(cc_contagem, x="Quantidade", y="Nome C.C", orientation="h", color="Quantidade", color_continuous_scale=[(0, "#0f172a"), (1, "#8b5cf6")], text="Quantidade")
                    fig_cc.update_layout(showlegend=False, xaxis_title="", yaxis_title="", margin=dict(l=0, r=40, t=10, b=0), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), height=max(300, len(cc_contagem) * 35))
                    fig_cc.update_yaxes(categoryorder="total ascending")
                    fig_cc.update_xaxes(visible=False)
                    fig_cc.update_coloraxes(showscale=False)
                    fig_cc.update_traces(textposition='outside', cliponaxis=False)
                    
                    st.plotly_chart(fig_cc, use_container_width=True)
                else:
                    st.info("Nenhum dado encontrado para gerar gráfico de C.C.")
                    
            with col_graf2:
                # Gráfico: MOD vs MOI
                st.markdown("**Proporção MOD vs MOI**")
                df_mod = df_cc_aba[df_cc_aba["MÃO DE OBRA"].str.strip() != ""]
                if not df_mod.empty:
                    mo_contagem = df_mod["MÃO DE OBRA"].value_counts().reset_index()
                    mo_contagem.columns = ["Tipo", "Quantidade"]
                    fig_mo = px.pie(mo_contagem, values="Quantidade", names="Tipo", hole=0.65, color_discrete_sequence=["#4a9eed", "#f39c12", "#e74c3c"])
                    fig_mo.update_layout(margin=dict(l=20, r=20, t=40, b=20), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), height=350, showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
                    
                    st.plotly_chart(fig_mo, use_container_width=True)
                else:
                    st.info("Dados de Mão de Obra não disponíveis.")
            
            st.markdown("---")
            
            # --- Seção de Histórico (Máquina do Tempo) ---
            if os.path.exists(caminho_hist_cc):
                try:
                    df_hist = pd.read_csv(caminho_hist_cc)
                    if not df_hist.empty and "DATA" in df_hist.columns:
                        st.markdown("**📈 Máquina do Tempo: Evolução do Efetivo**")
                        # Filtro para escolher o CC ou Área
                        opcoes_historico = ["Geral (Todos)", "Resumo: PB (Caldeira)", "Resumo: RB (Retorta)", "Resumo: ESP (Precipitador)"] + lista_cc
                        cc_selecionado = st.selectbox("Selecione a equipe ou área para analisar o crescimento:", opcoes_historico)
                        
                        if cc_selecionado == "Geral (Todos)":
                            df_plot = df_hist.groupby("DATA")["Efetivo"].sum().reset_index()
                            titulo_graf = "Crescimento Geral da Obra"
                        elif cc_selecionado == "Resumo: PB (Caldeira)":
                            df_plot = df_hist[df_hist["C.C"].apply(lambda x: "125.02" in str(x) and ".005" not in str(x))].groupby("DATA")["Efetivo"].sum().reset_index()
                            titulo_graf = "Crescimento - Área PB (Caldeira)"
                        elif cc_selecionado == "Resumo: RB (Retorta)":
                            df_plot = df_hist[df_hist["C.C"].apply(lambda x: "125.01" in str(x) and ".005" not in str(x))].groupby("DATA")["Efetivo"].sum().reset_index()
                            titulo_graf = "Crescimento - Área RB (Retorta)"
                        elif cc_selecionado == "Resumo: ESP (Precipitador)":
                            df_plot = df_hist[df_hist["C.C"].apply(lambda x: ".005" in str(x))].groupby("DATA")["Efetivo"].sum().reset_index()
                            titulo_graf = "Crescimento - Área ESP (Precipitador)"
                        else:
                            df_plot = df_hist[df_hist["C.C"] == cc_selecionado].copy()
                            titulo_graf = f"Evolução - C.C {cc_selecionado}"
                            
                        if not df_plot.empty:
                            df_plot["DATA_DT"] = pd.to_datetime(df_plot["DATA"], errors='coerce')
                            df_plot = df_plot.sort_values("DATA_DT")
                            
                            fig_hist = px.line(df_plot, x="DATA", y="Efetivo", markers=True, title=titulo_graf, line_shape="spline")
                            fig_hist.update_layout(xaxis_title="", yaxis_title="Quantidade de Colaboradores", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), height=300)
                            fig_hist.update_xaxes(type='category')
                            fig_hist.update_yaxes(tickformat="d")
                            fig_hist.update_traces(line=dict(width=3, color="#0ea5e9"), marker=dict(size=8, color="#10b981"))
                            st.plotly_chart(fig_hist, use_container_width=True)
                        else:
                            st.info("Aguardando acumular mais dias de dados para gerar a curva.")
                except Exception:
                    pass
            # ---------------------------------------------
            st.markdown("**Liderança: Efetivo de Encarregados por C.C**")
            # Tabela sumarizando quantos encarregados tem em cada CC (Apenas encarregados reais da lista oficial)
            df_lideres = df_cc_aba[(df_cc_aba["ENCARREGADO"].str.strip() != "") & (df_cc_aba["ENCARREGADO"].isin(lista_completa_encarregados))].copy()
            
            if len(df_lideres) > 0:
                df_agrupado = df_lideres.groupby(["C.C", "ENCARREGADO"]).size().reset_index(name="QTD. COLABORADORES")
                df_agrupado["LOCAL"] = df_agrupado["C.C"].apply(lambda x: "PB" if "125.02" in str(x) else ("RB" if "125.01" in str(x) else "OUTROS"))
                df_agrupado["NOME C.C"] = df_agrupado["C.C"].apply(format_cc)
                
                # Reorganizar colunas
                df_agrupado = df_agrupado[["LOCAL", "NOME C.C", "ENCARREGADO", "QTD. COLABORADORES"]].sort_values(by=["LOCAL", "NOME C.C", "QTD. COLABORADORES"], ascending=[True, True, False])
                
                # Exibe a tabela agrupada
                st.dataframe(df_agrupado, hide_index=True, use_container_width=True)
            else:
                st.info("Nenhum encarregado da lista oficial vinculado a um Centro de Custo para o filtro selecionado.")
                
            st.markdown("---")
            
            # Filtros por C.C. e Equipe
            st.markdown("**Consulta Detalhada**")
            
            # Filtramos a lista de encarregados para exibir APENAS quem realmente é encarregado da lista oficial
            lista_encarregados_detalhada = sorted([str(e) for e in df_cc_aba["ENCARREGADO"].unique() if str(e).strip() != "" and str(e) in lista_completa_encarregados])
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                cc_selecionado = st.selectbox("Selecione o Centro de Custo:", ["TODOS"] + lista_cc, format_func=format_cc)
            with col_f2:
                enc_selecionado = st.selectbox("Selecione a Equipe (Encarregado):", ["TODAS AS EQUIPES"] + lista_encarregados_detalhada)
            
            df_cc_filtrado = df_cc_aba.copy()
            
            if cc_selecionado != "TODOS":
                df_cc_filtrado = df_cc_filtrado[df_cc_filtrado["C.C"] == cc_selecionado]
                
            if enc_selecionado != "TODAS AS EQUIPES":
                df_cc_filtrado = df_cc_filtrado[df_cc_filtrado["ENCARREGADO"] == enc_selecionado]
            
            if len(df_cc_filtrado) > 0:
                # Resumo de funções no C.C. selecionado
                st.markdown(f"**Funções no C.C. selecionado** ({len(df_cc_filtrado)} colaboradores)")
                func_cc = df_cc_filtrado["FUNÇÃO"].value_counts().reset_index()
                func_cc.columns = ["Função", "Quantidade"]
                
                fig_func = px.bar(func_cc, x="Quantidade", y="Função", orientation="h", color="Quantidade", color_continuous_scale="Oranges", text="Quantidade")
                fig_func.update_layout(showlegend=False, xaxis_title="", yaxis_title="", margin=dict(l=0, r=0, t=10, b=0), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"), height=max(200, len(func_cc) * 35))
                fig_func.update_yaxes(categoryorder="total ascending")
                fig_func.update_xaxes(visible=False)
                fig_func.update_coloraxes(showscale=False)
                fig_func.update_traces(textposition='outside')
                if st.toggle("📊 Visualizar Gráfico de Funções"):
                    st.plotly_chart(fig_func, use_container_width=True)
                
                # Tabela detalhada
                colunas_exibir = ["MATRICULA", "NOME", "FUNÇÃO", "C.C", "ENCARREGADO"]
                if "DISCIPLINA" in df_cc_filtrado.columns:
                    colunas_exibir.append("DISCIPLINA")
                if "MÃO DE OBRA" in df_cc_filtrado.columns:
                    colunas_exibir.append("MÃO DE OBRA")
                colunas_exibir = [c for c in colunas_exibir if c in df_cc_filtrado.columns]
                
                # Botão de download
                buf_cc = io.BytesIO()
                df_cc_filtrado[colunas_exibir].to_excel(buf_cc, index=False)
                buf_cc.seek(0)
                nome_cc_arq = cc_selecionado.replace(".", "_") if cc_selecionado != "TODOS" else "TODOS"
                st.download_button("⬇️ Baixar Relatório C.C (.xlsx)", data=buf_cc, file_name=f"CC_{nome_cc_arq}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
                
                st.dataframe(df_cc_filtrado[colunas_exibir].reset_index(drop=True), hide_index=True, use_container_width=True)
            else:
                st.info("Nenhum colaborador encontrado para este Centro de Custo.")

    if tab_rdc_digital is not None:
      with tab_rdc_digital:
        st.markdown("### <span class='material-symbols-rounded' style='vertical-align: middle; color: #0ea5e9; font-size: 32px;'>edit_document</span> Lançamento de RDC Digital", unsafe_allow_html=True)
        st.caption("Preencha as informações do seu dia de trabalho seguindo as 3 etapas abaixo. Os dados serão salvos na nuvem.")
        
        with st.form("form_rdc_digital"):
            tab_id, tab_local, tab_ativ = st.tabs(["1️⃣ Identificação", "2️⃣ Localização", "3️⃣ Atividades e Envio"])
            
            with tab_id:
                st.markdown("<p style='color: #94a3b8; font-size: 14px;'>Quem é você e qual seu turno?</p>", unsafe_allow_html=True)
                rdc_encarregado = st.selectbox("Selecione seu Nome (Encarregado):", [""] + lista_completa_encarregados)
                rdc_turno = st.selectbox("Turno de Trabalho:", ["DIURNO", "NOTURNO", "MISTO"])
                
            with tab_local:
                import datetime
                st.markdown("<p style='color: #94a3b8; font-size: 14px;'>Onde você trabalhou hoje?</p>", unsafe_allow_html=True)
                
                rdc_data = st.date_input("Data do Relatório:", datetime.date.today())
                
                area_options = ["PB", "RB", "ESP", "LAYDOWN 1", "LAYDOWN 2", "OUTRO (DIGITAR)"]
                area_sel = st.selectbox("Área / Local de Trabalho:", area_options)
                rdc_area = area_sel
                if area_sel == "OUTRO (DIGITAR)":
                    rdc_area = st.text_input("Qual Área/Local?", placeholder="Ex: Escritório, Almoxarifado...")
                
                disc_options = [
                    "EQUIPAMENTOS", "DUTOS", "TUBULACAO", "ESTRUTURA METALICA", "PRECIPITADOR", 
                    "PRESSAO - MECANICA", "PRESSAO - TUBULACAO", "PRESSAO - FORNALHA", "PINTURA", 
                    "COMISSIONAMENTO", "OP. ASSISTIDA", "LAVAGEM QUIMICA", "SOPRAGEM", "ANDAIME", 
                    "OPERADORES", "FORA DE ESCOPO", "GERENCIA", "PRODUCAO", "GARANTIA DA QUALIDADE", 
                    "PLANEJAMENTO", "ADMINISTRACAO", "SEGURANCA E MEDICINA DO TRABALHO", "INFRAESTRUTURA", 
                    "ALMOXARIFADO ENESA", "ALMOXARIFADO MATERIAIS", "MANUT. ELETRICA PROVISORIA", 
                    "TOPOGRAFIA", "MOVIMENTACAO DE CARGAS", "MEDICAO/CUSTO/CONTRATOS", "CIVIL", "MECÂNICA", "ELÉTRICA", "INSTRUMENTAÇÃO", "ISOLAMENTO", "OUTRA (DIGITAR)"
                ]
                disc_sel = st.selectbox("Disciplina Principal:", disc_options)
                
                rdc_disciplina = disc_sel
                if disc_sel == "OUTRA (DIGITAR)":
                    rdc_disciplina = st.text_input("Qual Disciplina?", placeholder="Ex: Tubulação, Solda...")
                    
            with tab_ativ:
                st.markdown("<p style='color: #94a3b8; font-size: 14px;'>O que foi executado?</p>", unsafe_allow_html=True)
                rdc_dds = st.text_input("Tópico do DDS do dia:")
                rdc_atividades = st.text_area("Atividades Executadas (Detalhe os serviços feitos pela equipe):", height=150)
                rdc_problemas = st.text_area("Problemas / Interrupções / Ocorrências (Opcional):", height=68)
                
                st.markdown("<br>", unsafe_allow_html=True)
                submit_rdc = st.form_submit_button("🚀 Salvar e Enviar RDC na Nuvem", use_container_width=True, type="primary")
            
            if submit_rdc:
                if not rdc_encarregado:
                    st.error("⚠️ Por favor, selecione o nome do Encarregado.")
                elif not rdc_atividades.strip():
                    st.error("⚠️ Por favor, preencha as Atividades Executadas.")
                elif disc_sel == "OUTRA (DIGITAR)" and not rdc_disciplina.strip():
                    st.error("⚠️ Digite a disciplina na caixa 'Qual Disciplina?'.")
                else:
                    rdc_json = [{
                        "ENCARREGADO": rdc_encarregado,
                        "DATA": rdc_data.strftime("%Y/%m/%d"),
                        "TURNO": rdc_turno,
                        "AREA": rdc_area.strip().upper(),
                        "DISCIPLINA": rdc_disciplina.strip().upper(),
                        "DDS": rdc_dds.strip(),
                        "ATIVIDADE": rdc_atividades.strip(),
                        "CALDEIRA": rdc_problemas.strip(),
                        "PROBLEMAS": rdc_problemas.strip()
                    }]
                    
                    import json
                    import requests
                    
                    WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxfE96gE7ckdmapBLBHJuoX2bvAt-2d76OUJNiSRsLgFCOiySeQhFOopp3DoC5Fn95D/exec"
                    
                    try:
                        with st.spinner("Enviando dados para a nuvem..."):
                            res = requests.post(WEBHOOK_URL, json=rdc_json, allow_redirects=True)
                        if res.status_code == 200:
                            st.success(f"✅ RDC Digital de {rdc_encarregado} salvo com sucesso na Nuvem!")
                            st.info("Para visualizar na tabela da IA, clique em 'Puxar Dados Automáticos' abaixo.")
                        else:
                            st.error(f"❌ Erro ao enviar. Servidor retornou: {res.text}")
                    except Exception as e:
                        st.error(f"❌ Falha de conexão: {e}")
        
        st.markdown("---")
        st.markdown("### 📲 Gerar Acesso Rápido (QR Code)")
        st.caption("Gere um QR Code para os encarregados escanearem e entrarem direto no formulário sem digitar senha.")
        url_site = st.text_input("Qual o link do seu site?", placeholder="Ex: https://rdo-enesa.streamlit.app")
        if st.button("Gerar QR Code", type="primary"):
            if not url_site.strip():
                st.warning("⚠️ Cole o link do site primeiro!")
            else:
                link_final = url_site.strip()
                if "?" in link_final:
                    link_final += "&pwd=Campo@2026"
                else:
                    if not link_final.endswith("/"):
                        link_final += "/"
                    link_final += "?pwd=Campo@2026"
                
                import urllib.parse
                link_encoded = urllib.parse.quote(link_final)
                qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=300x300&data={link_encoded}"
                
                col_qr1, col_qr2 = st.columns([1, 2])
                with col_qr1:
                    st.image(qr_url, caption="Escaneie para logar", use_container_width=True)
                with col_qr2:
                    st.success("QR Code gerado com sucesso!")
                    st.markdown(f"**Link de acesso rápido:**\n[{link_final}]({link_final})")
                    st.markdown(f"**[📥 Clique aqui para abrir a imagem do QR Code e Salvar]({qr_url})**")
        
        st.markdown("---")
        st.markdown("### 📥 Sincronização de RDCs (Nuvem)")
        st.caption("Clique no botão abaixo para puxar todos os RDCs lançados pelos encarregados no sistema.")
        
        WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxfE96gE7ckdmapBLBHJuoX2bvAt-2d76OUJNiSRsLgFCOiySeQhFOopp3DoC5Fn95D/exec"
        
        if st.button("🔄 Puxar Dados Automáticos (Google Sheets)", type="primary", use_container_width=True):
            with st.spinner("Conectando ao Banco de Dados na Nuvem..."):
                try:
                    import requests
                    response = requests.get(WEBHOOK_URL, timeout=15)
                    
                    if response.status_code == 200:
                        dados_offline = response.json()
                        
                        if isinstance(dados_offline, list) and len(dados_offline) > 0:
                            if 'df_ia' not in st.session_state:
                                st.session_state.df_ia = pd.DataFrame(columns=['ITEM', 'DATA', 'DISCIPLINA', 'ENCARREGADO', 'TURNO', 'DDS', 'ATIVIDADE', 'PROBLEMAS', 'LOCAL', 'AREA'])
                                
                            ultimo_item = st.session_state.df_ia['ITEM'].max() if not st.session_state.df_ia.empty and pd.notna(st.session_state.df_ia['ITEM'].max()) else 0
                            
                            novos_registros = []
                            for r in dados_offline:
                                ultimo_item += 1
                                novo_reg = {
                                    'ITEM': ultimo_item,
                                    'DATA': r.get('DATA', ''),
                                    'DISCIPLINA': str(r.get('DISCIPLINA', '')).strip().upper(),
                                    'ENCARREGADO': r.get('ENCARREGADO', ''),
                                    'TURNO': r.get('TURNO', ''),
                                    'DDS': r.get('TOPICO_DDS', r.get('DDS', '')),
                                    'ATIVIDADE': r.get('ATIVIDADES', r.get('ATIVIDADE', '')),
                                    'PROBLEMAS': r.get('PROBLEMAS', r.get('CALDEIRA', '')),
                                    'LOCAL': str(r.get('AREA', '')).strip().upper(),
                                    'AREA': str(r.get('AREA', '')).strip().upper()
                                }
                                novos_registros.append(novo_reg)
                                
                            st.session_state.df_ia = pd.concat([st.session_state.df_ia, pd.DataFrame(novos_registros)], ignore_index=True)
                            st.success(f"📦 Sincronização Automática concluída! {len(novos_registros)} RDCs puxados do Google Sheets com sucesso.")
                            st.balloons()
                        else:
                            st.info("👍 Nenhum RDC novo pendente no Google Sheets no momento.")
                    else:
                        st.error(f"❌ Erro de conexão. Código HTTP: {response.status_code}")
                except Exception as e:
                    st.error(f"❌ Falha de rede ao tentar conectar com a nuvem: {e}")


else:
    st.markdown(f"""<div style="background-color: {cor_card}; padding: 50px; border-radius: 12px; text-align: center; border: 2px dashed {cor_borda}; margin-top: 50px;"><h2 style="color: {cor_texto} !important; margin-bottom: 10px;">Aguardando base de dados</h2><p style="font-size: 1rem; color: {cor_texto_sub};">Arraste o arquivo de efetivo (.csv ou .xlsx) na barra lateral para começar.</p></div>""", unsafe_allow_html=True)