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

# Paleta suave — escura mas não agressiva
cor_fundo = "#1a1d23"
cor_card = "#22262e"
cor_borda = "#2d323b"
cor_texto = "#e0e4ea"
cor_texto_sub = "#8b919e"
cor_azul = "#4a9eed"
cor_azul_hover = "#3a85d6"
cor_verde = "#4ade80"
cor_laranja = "#f59e42"
cor_destaque = "#6366f1"

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    @keyframes fadeIn {{
        from {{ opacity: 0; transform: translateY(8px); }}
        to {{ opacity: 1; transform: translateY(0); }}
    }}
    
    /* === BASE === */
    .stApp {{
        font-family: 'Inter', sans-serif !important;
        background-color: {cor_fundo};
        color: {cor_texto};
    }}
    
    .block-container {{
        animation: fadeIn 0.6s ease-out;
        max-width: 1200px;
        padding-top: 1rem;
    }}

    /* === CABEÇALHO === */
    .enesa-header {{
        text-align: center;
        margin-top: -30px;
        margin-bottom: 28px;
        padding: 28px 20px;
        background: {cor_card};
        border-radius: 14px;
        border-left: 4px solid {cor_azul};
        box-shadow: 0 2px 12px rgba(0,0,0,0.2);
    }}
    
    /* === TIPOGRAFIA === */
    h1, h2, h3 {{
        color: {cor_texto} !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 600 !important;
        letter-spacing: -0.3px;
    }}
    p, span, label, div {{
        font-family: 'Inter', sans-serif !important;
    }}
    
    /* === BOTÃO PRINCIPAL (verde suave) === */
    div.stButton > button[data-baseweb="button"] {{
        background: {cor_azul};
        color: white;
        border: none;
        font-weight: 500;
        font-size: 0.9rem;
        border-radius: 8px;
        padding: 10px 24px;
        transition: all 0.25s ease;
        box-shadow: 0 2px 8px rgba(74, 158, 237, 0.2);
    }}
    div.stButton > button[data-baseweb="button"]:hover {{
        background: {cor_azul_hover};
        transform: translateY(-1px);
        box-shadow: 0 4px 14px rgba(74, 158, 237, 0.35);
    }}
    
    /* === SIDEBAR === */
    [data-testid="stSidebar"] {{
        background-color: {cor_card};
        border-right: 1px solid {cor_borda};
    }}
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {{
        color: {cor_texto} !important;
    }}
    
    /* === INPUTS === */
    .stTextInput input, .stSelectbox > div > div {{
        border-radius: 8px;
        border: 1px solid {cor_borda};
        background-color: {cor_fundo};
        color: {cor_texto};
        transition: border-color 0.2s;
    }}
    .stTextInput input:focus {{
        border-color: {cor_azul};
        box-shadow: 0 0 0 2px rgba(74, 158, 237, 0.15);
    }}
    
    /* === ABAS === */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px;
        background-color: {cor_card};
        border-radius: 10px;
        padding: 4px;
        border: 1px solid {cor_borda};
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 8px;
        font-weight: 500;
        font-size: 0.9rem;
        font-family: 'Inter', sans-serif !important;
        padding: 10px 18px;
        color: {cor_texto_sub};
        transition: all 0.2s ease;
    }}
    .stTabs [data-baseweb="tab"]:hover {{
        color: {cor_texto};
        background-color: rgba(74, 158, 237, 0.08);
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {cor_azul} !important;
        color: white !important;
        box-shadow: 0 2px 8px rgba(74, 158, 237, 0.25);
    }}
    
    /* === MÉTRICAS === */
    [data-testid="stMetric"] {{
        background-color: {cor_card};
        border: 1px solid {cor_borda};
        border-radius: 10px;
        padding: 16px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.1);
    }}
    [data-testid="stMetricLabel"] {{
        color: {cor_texto_sub} !important;
    }}
    [data-testid="stMetricValue"] {{
        color: {cor_texto} !important;
        font-weight: 600 !important;
    }}
    
    /* === DATAFRAMES === */
    .stDataFrame {{
        border-radius: 10px;
        overflow: hidden;
    }}
    
    /* === FILE UPLOADER FIX PARA TRADUTOR DO CHROME === */
    [data-testid="stFileUploader"] {{
        border-radius: 10px;
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
        font-family: 'Inter', sans-serif !important;
        font-size: 14px;
        white-space: nowrap;
    }}
    
    /* === ALERTAS === */
    .stAlert {{
        border-radius: 10px;
    }}
    
    /* === SCROLLBAR SUAVE === */
    ::-webkit-scrollbar {{
        width: 6px;
    }}
    ::-webkit-scrollbar-track {{
        background: {cor_fundo};
    }}
    ::-webkit-scrollbar-thumb {{
        background: {cor_borda};
        border-radius: 3px;
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
        df[c] = df[c].replace(["nan", "NaN", "None", "0.0", "0"], "")

    for c in ["C.C", "DISCIPLINA", "MÃO DE OBRA"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
        df[c] = df[c].replace(["nan", "NaN", "None", "0.0", "0"], "")

    df["MATRICULA"] = df["MATRICULA"].str.replace(".0", "", regex=False)
    
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
# SESSION STATE
# =================================================================
if 'df' not in st.session_state:
    st.session_state.df = None
if 'df_ia' not in st.session_state:
    st.session_state.df_ia = pd.DataFrame(columns=['ITEM', 'DISCIPLINA', 'ENCARREGADO', 'TURNO', 'ATIVIDADE', 'CALDEIRA', 'LOCAL', 'AREA'])
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
            
        with st.container(border=True):
            st.markdown(f"<h3 style='text-align: center; color: {cor_azul}; margin-bottom: 0px; font-weight: 600;'>{nome_site}</h3>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; color: #8b919e; margin-bottom: 25px; font-size: 14px;'>Acesso Restrito ao Sistema</p>", unsafe_allow_html=True)
            
            user_input = st.text_input("Usuário (Login):")
            pass_input = st.text_input("Senha:", type="password")
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
# CABEÇALHO GLOBAL (Mostrado apenas se logado)
# =================================================================
st.markdown(f"""
    <div class="enesa-header">
        <h1 style="margin: 0; font-size: 1.8rem; font-weight: 700;">Sistema de Gestão RDC & PDE</h1>
        <p style="color: {cor_texto_sub}; font-size: 0.95rem; margin: 6px 0 0 0;">{nome_site} - Controle Operacional de Efetivo</p>
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
        
        with st.expander("🔑 Ver Usuários e Senhas"):
            usuarios_carregados = carregar_usuarios()
            dados_usuarios = []
            for u_nome, u_dados in usuarios_carregados.items():
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
            nova_role = st.selectbox("Nível de Acesso:", ["user", "admin"])
            submit_user = st.form_submit_button("Salvar Usuário")
            if submit_user and novo_user and nova_senha:
                usuarios_db[novo_user] = {"senha": nova_senha, "nome": novo_nome, "role": nova_role}
                salvar_usuarios(usuarios_db)
                st.success(f"Usuário '{novo_user}' salvo!")
                time.sleep(1)
                st.rerun()
        
        st.markdown("**Usuários Cadastrados:**")
        for u, dados in usuarios_db.items():
            col_u, col_del = st.columns([3, 1])
            col_u.text(f"👤 {u} ({dados.get('role', 'user')})")
            if u != "admin":
                if col_del.button("❌", key=f"del_{u}"):
                    del usuarios_db[u]
                    salvar_usuarios(usuarios_db)
                    st.rerun()

    st.markdown("---")
    st.markdown(
        f"""
        <div style='text-align: center; margin-top: 20px; font-size: 12px; color: #888;'>
            <p>Desenvolvido por</p>
            <p style='font-size: 16px; font-weight: bold; color: #ff4b4b; margin-top: -10px;'>Edson Garcia</p>
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
                    mapa_cc = dict(zip(df_com_cc["MATRICULA"], df_com_cc["C.C"]))
                    
                    df_carregado["C.C"] = df_carregado.apply(
                        lambda row: mapa_cc.get(row["MATRICULA"], row["C.C"]) if mapa_cc.get(row["MATRICULA"]) else row["C.C"], axis=1
                    )
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

    encarregados_f1_oficial = [
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
    lista_completa_encarregados = sorted([e.upper() for e in encarregados_f1_oficial])

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

    tab_dashboard, tab_resumo, tab_emissao, tab_cc, tab_f1, tab_ia, tab_ia_cc = st.tabs(["📊 Dashboard", "📅 Resumo Diário", "📝 Emissão de RDC", "💰 Controle de C.C", "🏎️ Competição F1", "🤖 Leitor de RDC (IA)", "🤖 IA - Atualizador de C.C"])

    with tab_dashboard:
        st.markdown("### Painel de Gestão")
        
        termo_busca = st.text_input("🔍 Buscar funcionário (Nome, Matrícula ou Função):")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("👷 Funcionários", len(df_atual))
        m2.metric("👔 Encarregados (Base)", len(lista_encarregados_base))
        m3.metric("🔧 Funções", df_atual["FUNÇÃO"].nunique())
        st.markdown("")
        
        mod_direta = ["MONTADOR DE ANDAIME", "SOLDADOR TIG/ER", "LIXADOR", "MECANICO MONTADOR", "ENCANADOR", "AJUDANTE", "SOLDADOR RX", "SOLDADOR TIG", "CALDEREIRO", "SOLDADOR MIG/ER"]
        df_mod = df_atual[df_atual["FUNÇÃO"].str.upper().isin(mod_direta)]
        
        if len(df_mod) > 0:
            st.markdown("**Mão de Obra Direta**")
            func_mod = df_mod["FUNÇÃO"].value_counts().reset_index()
            func_mod.columns = ["Função", "Quantidade"]
            
            fig_mod = px.bar(func_mod, x="Quantidade", y="Função", orientation="h", color="Quantidade", color_continuous_scale="Oranges", text="Quantidade")
            fig_mod.update_layout(showlegend=False, xaxis_title="", yaxis_title="", margin=dict(l=0, r=0, t=10, b=0), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#e0e4ea"))
            fig_mod.update_yaxes(categoryorder="total ascending")
            fig_mod.update_xaxes(visible=False)
            fig_mod.update_coloraxes(showscale=False)
            fig_mod.update_traces(textposition='outside')
            if st.toggle("📊 Visualizar Gráfico de Mão de Obra Direta"):
                st.plotly_chart(fig_mod, use_container_width=True)
            
        st.markdown("")
        st.markdown("**Base Completa**")
        df_exibicao = df_atual[["MATRICULA", "NOME", "FUNÇÃO", "ENCARREGADO"]].copy()
        if termo_busca:
            mask = (
                df_exibicao["NOME"].astype(str).str.contains(termo_busca, case=False, na=False) |
                df_exibicao["MATRICULA"].astype(str).str.contains(termo_busca, case=False, na=False) |
                df_exibicao["FUNÇÃO"].astype(str).str.contains(termo_busca, case=False, na=False)
            )
            df_exibicao = df_exibicao[mask]
        st.dataframe(df_exibicao, hide_index=True, use_container_width=True)

    with tab_resumo:
        st.markdown("### 📅 Resumo Diário")
        hoje = datetime.date.today().strftime("%Y-%m-%d")
        
        df_hoje = pd.DataFrame()
        if "df_historico_f1" in st.session_state and not st.session_state.df_historico_f1.empty:
            df_hist = st.session_state.df_historico_f1.copy()
            df_hist["DATA_STR"] = pd.to_datetime(df_hist["DATA"], errors="coerce").dt.strftime("%Y-%m-%d")
            df_hoje = df_hist[df_hist["DATA_STR"] == hoje]
            
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
            st.error(f"**Atenção:** {encarregados_pendentes} encarregados ainda não entregaram o RDC hoje.")
            entregues_list = df_hoje["ENCARREGADO"].unique() if not df_hoje.empty else []
            pendentes_list = [e for e in lista_completa_encarregados if e not in entregues_list]
            df_pend = pd.DataFrame({"Encarregados Pendentes (Hoje)": pendentes_list})
            st.dataframe(df_pend, hide_index=True, use_container_width=True)
        else:
            st.success("🎉 Todos os RDCs de hoje já foram entregues!")

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
                            with open(os.path.join(pasta_hist, f"{hoje.strftime('%d_%H%M')}_{nome_arquivo}"), "wb") as f:
                                f.write(buffer.getvalue())
                        except Exception:
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
                            st.download_button(f"⬇️ Baixar Todos ({qtd} arquivos)", data=zip_buffer, file_name=f"LOTE_RDC_{datetime.datetime.now().strftime('%d_%m_%Y')}.zip", mime="application/zip", use_container_width=True)
                            st.success(f"✅ {qtd} planilhas geradas!")
                        except Exception as e:
                            st.error(f"Erro: {e}")

    with tab_f1:
        st.markdown("### 🏎️ Competição F1 - Entrega de RDC")
        st.markdown("Acompanhamento mensal da entrega dos Relatórios Diários de Campo (RDC).")
        
        # A lista completa foi movida para cima para ser compartilhada com a aba de Resumo Diário
        
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
            df_mes = df_mes[df_mes["ENCARREGADO"].isin(lista_completa_encarregados)]
        else:
            df_mes = pd.DataFrame(columns=["DATA", "ENCARREGADO"])
            
        import calendar
        ano, mes = map(int, mes_selecionado.split('-'))
        num_dias = calendar.monthrange(ano, mes)[1]
        
        # Montar a Matriz
        dias_str = [str(d) for d in range(1, num_dias + 1)]
        
        # Identificar sábados e domingos
        dias_fim_de_semana = set()
        for d in range(1, num_dias + 1):
            data_check = datetime.date(ano, mes, d)
            if data_check.weekday() >= 5:  # 5=Sábado, 6=Domingo
                dias_fim_de_semana.add(str(d))
        
        dias_uteis = [d for d in dias_str if d not in dias_fim_de_semana]
        
        matriz = pd.DataFrame(index=lista_completa_encarregados, columns=dias_str)
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
                
        # Total conta apenas dias úteis (ignora fins de semana)
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
            elif valor == "➖":
                return "background-color: rgba(128, 128, 128, 0.2); color: #888;"
            return ""
            
        try:
            matriz_estilizada = matriz.style.map(cor_fundo)
        except AttributeError:
            matriz_estilizada = matriz.style.applymap(cor_fundo)
            
        st.dataframe(matriz_estilizada, use_container_width=True)
        
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
        
        st.success("🥇 Os 3 que MAIS entregaram")
        top3 = ranking.head(3)
        for i, row in top3.iterrows():
            medalha = "🥇" if i == 0 else ("🥈" if i == 1 else "🥉")
            st.markdown(f"**{medalha} {row['ENCARREGADO']}** ({row['ENTREGAS']} RDCs)")
                
        st.markdown("---")
        st.markdown("#### 📈 Evolução Mensal")
        df_evolucao = df_hist.groupby("MES_ANO").size().reset_index(name="RDCs Entregues")
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
                chave_padrao = st.text_input("🔑 Cole sua Chave da API Gemini:", type="password", help="Acesse https://aistudio.google.com/apikey para gerar sua chave. Ela fica oculta e protegida.")
            
            if not chave_padrao:
                st.info("☝️ Cole a sua chave de API do Gemini acima para ativar o robô de leitura.")
            
            arquivos_scan = st.file_uploader("Upload de RDCs Escaneados (PDF, JPG, PNG)", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
                
            btn_processar = st.button("🚀 Processar Arquivos com IA", type="primary", use_container_width=True)
            
            if btn_processar and arquivos_scan and chave_padrao:
                # --- FIX: Evitar que o Gemini tente usar o Service Account do Google Sheets ---
                old_cred = os.environ.pop("GOOGLE_APPLICATION_CREDENTIALS", None)
                
                client = genai.Client(api_key=chave_padrao)
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
                    "ATIVIDADE": "...",
                    "CALDEIRA": "...",
                    "LOCAL": "...",
                    "AREA": "..."
                  }}
                ]

                Regras de negócio:
                - DATA: Extraia a data em que o RDC foi preenchido. Retorne RIGOROSAMENTE no formato YYYY-MM-DD (Ano-Mês-Dia).
                - DISCIPLINA: Extraia a disciplina ou função do topo, mas RETORNE APENAS A PRIMEIRA PALAVRA OU A PALAVRA PRINCIPAL (ex: MECÂNICA, SOLDA, TOPOGRAFIA, CALDEIRARIA). Se for montador de andaime escreva ANDAIME. Sempre apenas 1 palavra.
                - ENCARREGADO: FAÇA O MÁXIMO ESFORÇO POSSÍVEL para descobrir quem é o encarregado. Compare o que está escrito à mão com esta lista oficial: [{nomes_para_prompt}]. Se a caligrafia estiver ruim, com erros de ortografia, ou se houver apenas o primeiro e segundo nome (ex: "Jailson Gois"), use dedução lógica e similaridade para encontrar a correspondência exata na lista. Retorne EXATAMENTE o nome completo que consta na lista fornecida. Somente se for 100% impossível deduzir quem é, retorne o texto 'AJUSTAR NOME'.
                - TURNO: Analise os horários. De dia (ex: 07:00 as 17:00) = 'DIURNO'. De noite = 'NOTURNO'.
                - ATIVIDADE: OBRIGATÓRIO: Crie um ÚNICO RESUMO SUPER CURTO E DIRETO de NO MÁXIMO 20 PALAVRAS sobre o que foi feito na seção 'ATIVIDADES'. Se você usar mais de 20 palavras, será considerado um erro gravíssimo! Extraia a ação principal, corrija a ortografia e escreva TUDO EM MAIÚSCULAS. NUNCA SEPARE EM LINHAS.
                - CALDEIRA: Se mencionar 'caldeira de recuperação' = 'RB'. Se 'caldeira de potência' = 'PB'. Se a descrição da atividade mencionar 'PRECIPITADOR' ou 'ESP' = 'ESP'. Se nenhum = ''.
                - LOCAL: Analise a imagem CUIDADOSAMENTE. Procure as opções 'PB ( )' e 'RB ( )'. Verifique se há um 'X', um rabisco, um visto ou qualquer marcação (mesmo que mal desenhada) dentro, em cima ou do lado dos parênteses. Retorne APENAS 'PB' ou 'RB' correspondente ao que estiver marcado. Se nenhum, retorne ''.
                - AREA: Analise as caixinhas de área na imagem com LUPA. Procure por qualquer marcação (X, visto, círculo, rabisco) dentro ou sobre os parênteses. As opções são exatamente: DUTO, EQUIPAMENTO, TUBULAÇÃO, ESTRUTURA MET, PRECIPITADOR, PRESSAO - MEC, PRESSAO - TUBULACAO, PRESSAO - FORNALHA, PINTURA, SOPRAGEM, ANDAIME. Retorne EXATAMENTE o nome da área que estiver marcada. Se nenhuma estiver marcada, retorne ''.

                Não inclua crases, formatação markdown ou texto adicional, apenas o JSON puro começando com [ e terminando com ].
                """

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
                                        model='gemini-2.5-flash',
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
                                    
                                        # Registro F1
                                        enc_lido = str(dados.get('ENCARREGADO', '')).strip()
                                        if enc_lido:
                                            # Usa a data do formulário RDC extraída pela IA. Se falhar, cai pra data de hoje.
                                            data_extraida = dados.get('DATA', '').strip()
                                            try:
                                                # tenta converter pra garantir que está em YYYY-MM-DD
                                                data_registro = pd.to_datetime(data_extraida).strftime('%Y-%m-%d')
                                            except:
                                                data_registro = datetime.date.today().strftime('%Y-%m-%d')
                                            
                                            ja_existe = ((st.session_state.df_historico_f1["DATA"] == data_registro) & (st.session_state.df_historico_f1["ENCARREGADO"] == enc_lido)).any()
                                            if not ja_existe:
                                                novo_registro = pd.DataFrame([{"DATA": data_registro, "ENCARREGADO": enc_lido}])
                                                st.session_state.df_historico_f1 = pd.concat([st.session_state.df_historico_f1, novo_registro], ignore_index=True)
                                                st.session_state.f1_modificado = True
                                    
                                        # Apenas extrai os dados, sem atualizar a base principal (a pedido do usuário)

                                    sucesso_arquivo = True
                                    break 

                                except Exception as inner_e:
                                    if '503' in str(inner_e):
                                        time.sleep(10)
                                    else:
                                        break
                                    
                            os.remove(tmp_path)
                        
                            if sucesso_arquivo:
                                st.toast(f"✅ {arquivo_scan.name} processado com sucesso!")
                            else:
                                st.toast(f"❌ Falha ao processar {arquivo_scan.name}.")
                            
                        except Exception as e:
                            st.error(f"Erro no envio do arquivo {arquivo_scan.name}: {e}")
                        
                        progresso.progress((i + 1) / total_arquivos)

                    status.update(label="🎉 Leitura concluída!", state="complete", expanded=False)
                    
                    # --- FIX: Salvar F1 na Nuvem em Lote ---
                    if st.session_state.get('f1_modificado', False):
                        if conn and not st.session_state.get('force_use_local', False):
                            try:
                                df_fresco = conn.read(worksheet="Historico_F1", ttl=0)
                                if not df_fresco.empty:
                                    df_fresco = df_fresco.dropna(how='all')
                                    df_final = pd.concat([df_fresco, st.session_state.df_historico_f1], ignore_index=True).drop_duplicates(subset=["DATA", "ENCARREGADO"])
                                else:
                                    df_final = st.session_state.df_historico_f1
                                
                                conn.update(worksheet="Historico_F1", data=df_final)
                                st.session_state.df_historico_f1 = df_final
                                st.cache_data.clear()
                            except Exception as e:
                                st.error(f"⚠️ Erro ao salvar histórico da F1 na nuvem: {e}")
                        st.session_state.f1_modificado = False
                time.sleep(2)
                st.session_state.force_use_local = True
                st.rerun()
                
            if not st.session_state.df_ia.empty:
                st.markdown("#### Dados Extraídos")
                
                lista_com_alerta = lista_encarregados_base + ["AJUSTAR NOME"]
                df_filtrado = st.session_state.df_ia[st.session_state.df_ia['ENCARREGADO'].isin(lista_com_alerta)]
                
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
                        st.session_state.df_ia = pd.DataFrame(columns=['ITEM', 'DISCIPLINA', 'ENCARREGADO', 'TURNO', 'ATIVIDADE', 'CALDEIRA', 'LOCAL'])
                        st.rerun()
                
                st.dataframe(df_filtrado, hide_index=True, use_container_width=True)

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
                chave_padrao = st.text_input("🔑 Cole sua Chave da API Gemini:", type="password", help="Chave oculta e protegida.", key="chave_cc")
            
            arquivos_scan_cc = st.file_uploader("Upload de RDCs para atualização de C.C (PDF, JPG, PNG)", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True, key="uploader_cc")
                
            btn_processar_cc = st.button("🚀 Atualizar C.C das Equipes com IA", type="primary", use_container_width=True)
            
            if btn_processar_cc and arquivos_scan_cc and chave_padrao:
                old_cred = os.environ.pop("GOOGLE_APPLICATION_CREDENTIALS", None)
                client = genai.Client(api_key=chave_padrao)
                nomes_para_prompt = ", ".join(lista_encarregados_base)
                
                prompt_ia_cc = f"""
                Analise este documento. Para CADA formulário de obra (RDC) encontrado no arquivo, extraia os dados.
                REGRA IMPORTANTÍSSIMA: Retorne APENAS UM objeto JSON por formulário/página.
                Retorne APENAS um array (lista) em formato JSON válido.
                [
                  {{
                    "DISCIPLINA": "...",
                    "ENCARREGADO": "...",
                    "CALDEIRA": "...",
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
                                    model='gemini-2.5-flash',
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
                                if '503' in str(inner_e):
                                    time.sleep(10)
                                else:
                                    break
                                    
                        os.remove(tmp_path)
                        
                        if sucesso_arquivo:
                            st.toast(f"✅ {arquivo_scan.name} processado com sucesso!")
                        else:
                            st.toast(f"❌ Falha ao processar {arquivo_scan.name}.")
                            
                    except Exception as e:
                        st.error(f"Erro no envio do arquivo {arquivo_scan.name}: {e}")
                        
                    progresso.progress((i + 1) / total_arquivos)

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
        
        if "C.C" not in df_atual.columns or df_atual["C.C"].str.strip().eq("").all():
            st.warning("⚠️ A coluna de Centro de Custo (C.C) não foi encontrada na base de dados atual. Verifique se a planilha possui essa coluna.")
        else:
            # Filtro PB/RB Global para a aba C.C
            filtro_local = st.segmented_control(
                "Filtrar Dados por Local:", 
                ["Ambas", "PB", "RB"], 
                default="Ambas"
            )
            if not filtro_local:
                filtro_local = "Ambas"
                
            df_cc_aba = df_atual[df_atual["C.C"].str.strip() != ""]
            if filtro_local == "PB":
                df_cc_aba = df_cc_aba[df_cc_aba["C.C"].apply(lambda x: "125.02" in str(x))]
            elif filtro_local == "RB":
                df_cc_aba = df_cc_aba[df_cc_aba["C.C"].apply(lambda x: "125.01" in str(x))]

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
            
            # Métricas gerais
            mc1, mc2, mc3, mc4, mc5 = st.columns(5)
            mc1.metric("🏢 Centros de Custo", len(lista_cc))
            mc2.metric("👷 Total Alocados", len(df_cc_aba))
            mc3.metric("🔧 Funções Distintas", df_cc_aba["FUNÇÃO"].nunique())
            
            qtd_encarregados = len([e for e in df_cc_aba["ENCARREGADO"].unique() if str(e).strip() != "" and str(e) in lista_completa_encarregados])
            mc4.metric("👨‍💼 Encarregados", qtd_encarregados)
            
            span_of_control = round(len(df_cc_aba) / qtd_encarregados, 1) if qtd_encarregados > 0 else 0
            mc5.metric("👥 Span of Control", span_of_control)
            st.markdown("")

            # Gráficos lado a lado
            col_graf1, col_graf2 = st.columns([6, 4])
            
            with col_graf1:
                # Gráfico de distribuição por C.C.
                st.markdown("**Distribuição de Efetivo por Centro de Custo**")
                
                cc_contagem = df_cc_aba["C.C"].value_counts().reset_index()
                cc_contagem.columns = ["Centro de Custo", "Quantidade"]
                cc_contagem["Nome C.C"] = cc_contagem["Centro de Custo"].apply(format_cc)
                
                if len(cc_contagem) > 0:
                    fig_cc = px.bar(cc_contagem, x="Quantidade", y="Nome C.C", orientation="h", color="Quantidade", color_continuous_scale="Blues", text="Quantidade")
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

else:
    st.markdown(f"""<div style="background-color: {cor_card}; padding: 50px; border-radius: 12px; text-align: center; border: 2px dashed {cor_borda}; margin-top: 50px;"><h2 style="color: {cor_texto} !important; margin-bottom: 10px;">Aguardando base de dados</h2><p style="font-size: 1rem; color: {cor_texto_sub};">Arraste o arquivo de efetivo (.csv ou .xlsx) na barra lateral para começar.</p></div>""", unsafe_allow_html=True)