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
try:
    from google import genai
except ImportError:
    pass

from streamlit_gsheets import GSheetsConnection

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Sistema RDC & PDE - ENESA", layout="wide", initial_sidebar_state="expanded")

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

st.markdown(f"""
    <div class="enesa-header">
        <h1 style="margin: 0; font-size: 1.8rem; font-weight: 700;">Sistema de Gestão RDC & PDE</h1>
        <p style="color: {cor_texto_sub}; font-size: 0.95rem; margin: 6px 0 0 0;">ENESA Engenharia · Controle Operacional de Efetivo</p>
    </div>
""", unsafe_allow_html=True)

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
            
        ws = wb.active
        ws[celula_encarregado] = encarregado_selecionado
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

# =================================================================
# BARRA LATERAL
# =================================================================
with st.sidebar:
    # CSS para esconder marcas d'água do Streamlit e botões quebrados
    st.markdown("""
        <style>
            #MainMenu {visibility: hidden;}
            header {visibility: hidden;}
            footer {visibility: hidden;}
            .st-emotion-cache-1z12wxe {display: none;} /* Esconder collapse button que quebrou */
        </style>
    """, unsafe_allow_html=True)

    if os.path.exists(caminho_logo):
        col1, col2, col3 = st.columns([1.5, 2, 1.5]) 
        with col2:
            st.image(caminho_logo, use_container_width=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
    st.header("📂 Arquivos Base")
    st.write("Atualize a base arrastando os ficheiros:")
    arquivo_pde = st.file_uploader("Base de Efetivo (.csv ou .xlsx):", type=["csv", "xlsx"])
    arquivo_modelo = st.file_uploader("📄 Layout MODELO.xlsx:", type=["xlsx"])
    
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
    st.markdown(
        """
        <div style='text-align: center; margin-top: 20px; font-size: 12px; color: #888;'>
            <p>Desenvolvido por</p>
            <p style='font-size: 16px; font-weight: bold; color: #ff4b4b; margin-top: -10px;'>Edson Garcia</p>
            <p style='margin-top: -10px;'>v5.0 · ENESA Engenharia (com ANTIGRAVITY)</p>
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
        st.session_state.df = df_carregado
        if conn:
            try:
                conn.update(worksheet="Página1", data=st.session_state.df)
                st.sidebar.success("☁️ Base salva no Google Sheets!")
            except Exception as e:
                st.sidebar.error(f"Erro Nuvem: {e}")

elif st.session_state.df is None:
    carregado_nuvem = False
    if conn:
        try:
            df_gsheets = conn.read(worksheet="Página1", ttl=5)
            df_gsheets = df_gsheets.dropna(how='all')
            if not df_gsheets.empty:
                st.session_state.df = df_gsheets
                carregado_nuvem = True
        except Exception:
            pass
            
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
    lista_encarregados = sorted([str(e) for e in df_atual["ENCARREGADO"].unique() if str(e).strip() != ""])

    tab_dashboard, tab_emissao, tab_cc, tab_ia = st.tabs(["📊 Dashboard", "🖨️ Emissão de RDC", "💰 Controle de C.C", "🤖 Leitor de RDC (IA)"])

    with tab_dashboard:
        st.markdown("### Painel de Gestão")
        m1, m2, m3 = st.columns(3)
        m1.metric("👷 Funcionários", len(df_atual))
        m2.metric("👔 Encarregados", len(lista_encarregados))
        m3.metric("🔧 Funções", df_atual["FUNÇÃO"].nunique())
        st.markdown("")
        
        mod_direta = ["MONTADOR DE ANDAIME", "SOLDADOR TIG/ER", "LIXADOR", "MECANICO MONTADOR", "ENCANADOR", "AJUDANTE", "SOLDADOR RX", "SOLDADOR TIG", "CALDEREIRO", "SOLDADOR MIG/ER"]
        df_mod = df_atual[df_atual["FUNÇÃO"].str.upper().isin(mod_direta)]
        
        if len(df_mod) > 0:
            st.markdown("**Mão de Obra Direta**")
            func_mod = df_mod["FUNÇÃO"].value_counts().reset_index()
            func_mod.columns = ["Função", "Quantidade"]
            chart1 = alt.Chart(func_mod).mark_bar(cornerRadiusTopRight=5, cornerRadiusBottomRight=5, color=cor_laranja).encode(x=alt.X("Quantidade:Q", title="Qtd"), y=alt.Y("Função:N", sort='-x', title="")).properties(height=max(200, len(func_mod) * 35))
            st.altair_chart(chart1, use_container_width=True)
            
        st.markdown("")
        st.markdown("**Base Completa**")
        st.dataframe(df_atual[["MATRICULA", "NOME", "FUNÇÃO", "ENCARREGADO"]], hide_index=True, use_container_width=True)

    with tab_emissao:
        st.markdown("### Emissão de RDC")
        if not lista_encarregados:
            st.warning("Nenhum encarregado encontrado na base.")
        else:
            encarregado_sel = st.selectbox("Escolha o Encarregado:", lista_encarregados)
            equipe = df_atual[df_atual["ENCARREGADO"] == encarregado_sel]
            st.markdown("")
            st.markdown(f"""<div style="background: {cor_card}; border-radius: 10px; padding: 20px; border: 1px solid {cor_borda}; margin-bottom: 16px;"><div style="text-align: center; border-bottom: 2px solid {cor_azul}; padding-bottom: 12px; margin-bottom: 12px;"><h3 style="margin: 0; font-size: 1.2rem; color: {cor_texto} !important;">RDC - Relatório Diário de Campo</h3><p style="color: {cor_texto_sub}; margin: 4px 0 0 0; font-size: 0.85rem;">ENESA Engenharia</p></div><table style="width: 100%; color: {cor_texto}; font-size: 0.9rem;"><tr><td style="padding: 4px 0;"><strong>Encarregado:</strong></td><td>{encarregado_sel}</td></tr><tr><td style="padding: 4px 0;"><strong>Data:</strong></td><td>{datetime.datetime.now().strftime("%d/%m/%Y")}</td></tr><tr><td style="padding: 4px 0;"><strong>Efetivo:</strong></td><td>{len(equipe)} colaborador(es)</td></tr></table></div>""", unsafe_allow_html=True)
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
                                for enc in lista_encarregados:
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
            chave_padrao = "AQ.Ab8RN6K3Sm_kxaKRqJDWdi5F9Xcpo2_ZGsuJ-eLATBTbampuhQ"
            api_key_input = st.text_input("🔑 Chave de API Gemini", value=chave_padrao, type="password")
            
            arquivos_scan = st.file_uploader("Upload de RDCs Escaneados (PDF, JPG, PNG)", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)
            
            if st.button("🚀 Processar Arquivos com IA", type="primary", use_container_width=True) and arquivos_scan:
                if not api_key_input:
                    st.warning("Insira a chave da API para continuar.")
                else:
                    client = genai.Client(api_key=api_key_input)
                    nomes_para_prompt = ", ".join(lista_encarregados)
                    
                    prompt_ia = f"""
                    Analise este documento (que pode ter várias páginas). Para CADA formulário de obra (RDC) encontrado no arquivo, extraia os dados.
                    Retorne APENAS um array (lista) em formato JSON válido. Exemplo do formato exato esperado:
                    [
                      {{
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
                    - DISCIPLINA: Extraia a disciplina ou função do topo, mas RETORNE APENAS A PRIMEIRA PALAVRA OU A PALAVRA PRINCIPAL (ex: MECÂNICA, SOLDA, TOPOGRAFIA, CALDEIRARIA). Se for montador de andaime escreva ANDAIME. Sempre apenas 1 palavra.
                    - ENCARREGADO: Extraia o nome do Encarregado escrito no papel. REGRA CRÍTICA: Compare o que está escrito com esta lista de encarregados válidos: [{nomes_para_prompt}]. Retorne EXATAMENTE o nome correspondente como está grafado na lista fornecida, corrigindo pequenos desvios do manuscrito. SE O NOME ESTIVER TOTALMENTE ILEGÍVEL OU NÃO ESTIVER NESSA LISTA DE FORMA ALGUMA, RETORNE EXATAMENTE O TEXTO 'AJUSTAR NOME'.
                    - TURNO: Analise os horários. De dia (ex: 07:00 as 17:00) = 'DIURNO'. De noite = 'NOTURNO'.
                    - ATIVIDADE: RESUMA A ATIVIDADE EM NO MÁXIMO 15 PALAVRAS!!! É ESTRITAMENTE PROIBIDO PASSAR DE 15 PALAVRAS, MESMO QUE O TEXTO ORIGINAL SEJA GIGANTE. Se passar de 15 palavras, corte o resto. Foco EXCLUSIVO na ação principal (ex: 'MONTAGEM DE SUPORTE NA CALDEIRA'). Corrija a ortografia e retorne TUDO EM LETRAS MAIÚSCULAS.
                    - CALDEIRA: Se mencionar 'caldeira de recuperação' = 'RB'. Se 'caldeira de potência' = 'PB'. Se a descrição da atividade mencionar 'PRECIPITADOR' ou 'ESP' = 'ESP'. Se nenhum = ''.
                    - LOCAL: Procure as caixinhas 'PB ( )' e 'RB ( )'. Se PB marcado com X = 'PB'. Se RB marcado com X = 'RB'. Se nenhum = ''.
                    - AREA: Procure as caixinhas de área/local de trabalho marcadas com X. As opções são: DUTO, EQUIPAMENTO, TUBULAÇÃO, ESTRUTURA MET, PRECIPITADOR, PRESSAO-MEC, PRESSAO-TUBULACAO, PRESSAO-FORNALHA, PINTURA, SOPRAGEM, ANDAIME. Retorne EXATAMENTE o nome da área marcada. Se nenhuma estiver marcada, retorne ''.

                    Não inclua crases, formatação markdown ou texto adicional, apenas o JSON puro começando com [ e terminando com ].
                    """

                    progresso = st.progress(0)
                    total_arquivos = len(arquivos_scan)
                    console_log = st.empty()
                    
                    for i, arquivo_scan in enumerate(arquivos_scan):
                        console_log.info(f"Processando arquivo {i+1} de {total_arquivos}: {arquivo_scan.name}...")
                        
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
                                        contents=[arquivo_up, prompt_ia]
                                    )
                                    texto_resposta = resposta.text.strip()
                                    if texto_resposta.startswith("```json"):
                                        texto_resposta = texto_resposta[7:-3].strip()
                                    elif texto_resposta.startswith("```"):
                                        texto_resposta = texto_resposta[3:-3].strip()

                                    dados_extraidos_lista = json.loads(texto_resposta)

                                    if isinstance(dados_extraidos_lista, dict):
                                        dados_extraidos_lista = [dados_extraidos_lista]

                                    # Mapeamento de AREA -> sufixo do C.C.
                                    mapa_area_sufixo = {
                                        'EQUIPAMENTO': '001', 'DUTO': '002', 'TUBULAÇÃO': '003', 'TUBULACAO': '003',
                                        'ESTRUTURA MET': '004', 'ESTRUTURA METALICA': '004', 'ESTRUTURA': '004',
                                        'PRECIPITADOR': '005', 'ESP': '005',
                                        'PRESSAO-MEC': '006', 'PRESSAO - MEC': '006', 'PRESSÃO-MEC': '006',
                                        'PRESSAO-TUBULACAO': '007', 'PRESSAO - TUBULACAO': '007', 'PRESSÃO-TUBULAÇÃO': '007',
                                        'PINTURA': '009', 'ANDAIME': '014', 'ANDAIMES': '014',
                                        'PRESSAO-FORNALHA': '003', 'PRESSAO - FORNALHA': '003',
                                        'SOPRAGEM': '015',
                                    }

                                    for dados in dados_extraidos_lista:
                                        ultimo_item = st.session_state.df_ia['ITEM'].max() if not st.session_state.df_ia.empty and pd.notna(st.session_state.df_ia['ITEM'].max()) else 0
                                        dados['ITEM'] = int(ultimo_item) + 1
                                        if 'LOCAL' not in dados:
                                            dados['LOCAL'] = ''
                                        if 'AREA' not in dados:
                                            dados['AREA'] = ''
                                        st.session_state.df_ia = pd.concat([st.session_state.df_ia, pd.DataFrame([dados])], ignore_index=True)
                                        
                                        # === ATUALIZAR C.C. COMPLETO NA BASE ===
                                        local_bruto = str(dados.get('LOCAL', '')).strip().upper()
                                        area_bruta = str(dados.get('AREA', '')).strip().upper()
                                        disciplina_lida = str(dados.get('DISCIPLINA', '')).strip().upper()
                                        enc_lido = str(dados.get('ENCARREGADO', '')).strip().upper()
                                        
                                        # Limpeza robusta do local (caso a IA retorne "PB (X)")
                                        local_lido = ''
                                        if 'PB' in local_bruto: local_lido = 'PB'
                                        elif 'RB' in local_bruto: local_lido = 'RB'
                                        # Fallback
                                        if not local_lido:
                                            cald = str(dados.get('CALDEIRA', '')).strip().upper()
                                            if 'PB' in cald: local_lido = 'PB'
                                            elif 'RB' in cald: local_lido = 'RB'
                                        
                                        # Limpeza robusta da área
                                        area_lida = ''
                                        for k in mapa_area_sufixo.keys():
                                            if k in area_bruta or k in disciplina_lida:
                                                area_lida = k
                                                break
                                                
                                        if enc_lido and enc_lido != 'AJUSTAR NOME' and 'C.C' in df_atual.columns:
                                            # Busca inteligente do Encarregado (Substring ou Aproximação)
                                            encarregados_unicos = df_atual['ENCARREGADO'].dropna().unique()
                                            enc_encontrado = None
                                            
                                            # 1. Match exato
                                            for e in encarregados_unicos:
                                                if str(e).strip().upper() == enc_lido:
                                                    enc_encontrado = e
                                                    break
                                            
                                            # 2. Substring (ex: IA lê "RAIMUNDO EUDE", base tem "RAIMUNDO EUDE DA SILVA")
                                            if not enc_encontrado:
                                                for e in encarregados_unicos:
                                                    if enc_lido in str(e).upper():
                                                        enc_encontrado = e
                                                        break
                                                        
                                            # 3. Fuzzy Match
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
                                                
                                                # Determinar prefixo (PB=125.02, RB=125.01)
                                                if local_lido in ['PB', 'RB']:
                                                    prefixo_novo = '125.02' if local_lido == 'PB' else '125.01'
                                                    
                                                    # Determinar sufixo pela AREA marcada
                                                    sufixo = mapa_area_sufixo.get(area_lida, '')
                                                    
                                                    if sufixo:
                                                        # C.C. completo: prefixo + sufixo
                                                        cc_novo = f"{prefixo_novo}.{sufixo}"
                                                        df_atual.loc[mask_enc, 'C.C'] = cc_novo
                                                        atualizado = True
                                                        st.toast(f"✅ C.C. de TODA A EQUIPE de {enc_encontrado} → {cc_novo}")
                                                    else:
                                                        # Sem área marcada, só troca o prefixo mantendo sufixo original
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
                                                    # Salvar fisicamente no Google Sheets
                                                    try:
                                                        conn_update = st.connection("gsheets", type=GSheetsConnection)
                                                        conn_update.update(worksheet="Página1", data=df_atual)
                                                    except Exception as e:
                                                        st.error(f"Erro ao salvar na nuvem: {e}")
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

                    console_log.success("🎉 Leitura concluída!")
                    time.sleep(2)
                    st.rerun()
                    
            if not st.session_state.df_ia.empty:
                st.markdown("#### Dados Extraídos")
                
                lista_com_alerta = lista_encarregados + ["AJUSTAR NOME"]
                df_filtrado = st.session_state.df_ia[st.session_state.df_ia['ENCARREGADO'].isin(lista_com_alerta)]
                
                col_dw1, col_dw2 = st.columns([1, 1])
                with col_dw1:
                    buf_excel = io.BytesIO()
                    df_filtrado.to_excel(buf_excel, index=False)
                    buf_excel.seek(0)
                    st.download_button("⬇️ Baixar Planilha RDC Lida (.xlsx)", data=buf_excel, file_name="Planilha_RDC_Lida.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
                with col_dw2:
                    if st.button("🗑️ Limpar Dados Lidos", use_container_width=True):
                        st.session_state.df_ia = pd.DataFrame(columns=['ITEM', 'DISCIPLINA', 'ENCARREGADO', 'TURNO', 'ATIVIDADE', 'CALDEIRA', 'LOCAL'])
                        st.rerun()
                
                st.dataframe(df_filtrado, use_container_width=True)

    with tab_cc:
        st.markdown("### 💰 Controle de Centro de Custo (C.C)")
        
        if "C.C" not in df_atual.columns or df_atual["C.C"].str.strip().eq("").all():
            st.warning("⚠️ A coluna de Centro de Custo (C.C) não foi encontrada na base de dados atual. Verifique se a planilha possui essa coluna.")
        else:
            lista_cc = sorted([str(cc) for cc in df_atual["C.C"].unique() if str(cc).strip() != ""])
            
            # Métricas gerais
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("🏢 Centros de Custo", len(lista_cc))
            mc2.metric("👷 Total Alocados", len(df_atual[df_atual["C.C"].str.strip() != ""]))
            mc3.metric("🔧 Funções Distintas", df_atual[df_atual["C.C"].str.strip() != ""]["FUNÇÃO"].nunique())
            st.markdown("")

            # Gráfico de distribuição por C.C.
            st.markdown("**Distribuição de Efetivo por Centro de Custo**")
            cc_contagem = df_atual[df_atual["C.C"].str.strip() != ""]["C.C"].value_counts().reset_index()
            cc_contagem.columns = ["Centro de Custo", "Quantidade"]
            chart_cc = alt.Chart(cc_contagem).mark_bar(cornerRadiusTopRight=5, cornerRadiusBottomRight=5, color=cor_destaque).encode(
                x=alt.X("Quantidade:Q", title="Qtd"),
                y=alt.Y("Centro de Custo:N", sort='-x', title="")
            ).properties(height=max(250, len(cc_contagem) * 28))
            st.altair_chart(chart_cc, use_container_width=True)
            
            st.markdown("---")
            
            # Filtros por C.C. e Equipe
            st.markdown("**Consulta Detalhada**")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                cc_selecionado = st.selectbox("Selecione o Centro de Custo:", ["TODOS"] + lista_cc)
            with col_f2:
                enc_selecionado = st.selectbox("Selecione a Equipe (Encarregado):", ["TODAS AS EQUIPES"] + lista_encarregados)
            
            df_cc_filtrado = df_atual[df_atual["C.C"].str.strip() != ""]
            
            if cc_selecionado != "TODOS":
                df_cc_filtrado = df_cc_filtrado[df_cc_filtrado["C.C"] == cc_selecionado]
                
            if enc_selecionado != "TODAS AS EQUIPES":
                df_cc_filtrado = df_cc_filtrado[df_cc_filtrado["ENCARREGADO"] == enc_selecionado]
            
            if len(df_cc_filtrado) > 0:
                # Resumo de funções no C.C. selecionado
                st.markdown(f"**Funções no C.C. selecionado** ({len(df_cc_filtrado)} colaboradores)")
                func_cc = df_cc_filtrado["FUNÇÃO"].value_counts().reset_index()
                func_cc.columns = ["Função", "Quantidade"]
                chart_func_cc = alt.Chart(func_cc).mark_bar(cornerRadiusTopRight=5, cornerRadiusBottomRight=5, color=cor_laranja).encode(
                    x=alt.X("Quantidade:Q", title="Qtd"),
                    y=alt.Y("Função:N", sort='-x', title="")
                ).properties(height=max(200, len(func_cc) * 30))
                st.altair_chart(chart_func_cc, use_container_width=True)
                
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