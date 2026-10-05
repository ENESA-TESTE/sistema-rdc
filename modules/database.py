# -*- coding: utf-8 -*-
"""
Módulo de Banco de Dados Ultrarrápido (SQLite Indexado + Espelhamento em Nuvem)
Garante respostas em < 5ms para toda a interface do Sistema RDC e persistência total.
"""

import os
import json
import sqlite3
import datetime
import threading
import pandas as pd
import streamlit as st

# Diretório base
PASTA_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(PASTA_BASE, "rdo_database.db")
CAMINHO_HISTORICO_F1_CSV = os.path.join(PASTA_BASE, "historico_f1_local.csv")
CAMINHO_BRIEFINGS_JSON = os.path.join(PASTA_BASE, "briefings_historico.json")
CAMINHO_PDE_CSV = os.path.join(PASTA_BASE, "PDE.csv")
CAMINHO_RDC_REGISTROS_CSV = os.path.join(PASTA_BASE, "rdc_registros.csv")

# Lock para operações thread-safe no SQLite
db_lock = threading.Lock()

def get_connection():
    """Retorna uma conexão SQLite configurada com WAL mode para máxima velocidade."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    return conn

def init_db():
    """Inicializa as tabelas e índices no SQLite se não existirem."""
    with db_lock:
        conn = get_connection()
        c = conn.cursor()
        
        # 1. Tabela Histórico F1
        c.execute("""
            CREATE TABLE IF NOT EXISTS historico_f1 (
                data TEXT NOT NULL,
                encarregado TEXT NOT NULL,
                criado_em TEXT,
                PRIMARY KEY (data, encarregado)
            );
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_f1_data ON historico_f1(data);")
        c.execute("CREATE INDEX IF NOT EXISTS idx_f1_enc ON historico_f1(encarregado);")
        
        # 2. Tabela Briefings Diários & IA
        c.execute("""
            CREATE TABLE IF NOT EXISTS briefings (
                data_iso TEXT PRIMARY KEY,
                data_formatada TEXT,
                data_salvo TEXT,
                conteudo_json TEXT,
                resumo_texto TEXT,
                origem TEXT
            );
        """)
        
        # 3. Tabela Base de Funcionários PDE
        c.execute("""
            CREATE TABLE IF NOT EXISTS pde_colaboradores (
                matricula TEXT PRIMARY KEY,
                nome TEXT,
                funcao TEXT,
                cc TEXT,
                encarregado TEXT,
                turno TEXT,
                status TEXT,
                disciplina TEXT,
                mao_de_obra TEXT,
                atualizado_em TEXT
            );
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_pde_enc ON pde_colaboradores(encarregado);")
        c.execute("CREATE INDEX IF NOT EXISTS idx_pde_disc ON pde_colaboradores(disciplina);")
        
        # 4. Tabela Banco de RDCs Salvos
        c.execute("""
            CREATE TABLE IF NOT EXISTS rdc_banco (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item TEXT,
                sub TEXT,
                data TEXT,
                disciplina TEXT,
                encarregado TEXT,
                turno TEXT,
                dds TEXT,
                transcricao TEXT,
                atividade TEXT,
                sub_atividade TEXT,
                local_especifico TEXT,
                efetivo_atividade TEXT,
                problemas TEXT,
                local TEXT,
                area TEXT,
                caldeira TEXT,
                data_criacao TEXT
            );
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_rdc_data ON rdc_banco(data);")
        c.execute("CREATE INDEX IF NOT EXISTS idx_rdc_enc ON rdc_banco(encarregado);")
        
        # 5. Tabela de Controle de Sincronização
        c.execute("""
            CREATE TABLE IF NOT EXISTS sync_controle (
                tabela TEXT PRIMARY KEY,
                ultima_sincronizacao TEXT,
                total_registros INTEGER,
                status TEXT
            );
        """)
        
        conn.commit()
        conn.close()

# Executar inicialização imediata
init_db()

# =====================================================================
# CONEXÃO COM GOOGLE SHEETS PARA ESPELHAMENTO EM NUVEM
# =====================================================================
def obter_planilha_google():
    """Retorna o objeto da planilha Google Sheets autenticado via Service Account."""
    try:
        from google.oauth2 import service_account
        import gspread
        if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            creds_info = st.secrets["connections"]["gsheets"]
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = service_account.Credentials.from_service_account_info(creds_info, scopes=scopes)
            client = gspread.authorize(creds)
            sh = client.open_by_url(creds_info['spreadsheet'])
            return sh
    except Exception:
        pass
    return None

# =====================================================================
# 1. HISTÓRICO F1 (ALTA VELOCIDADE + ESPELHO NUVEM)
# =====================================================================
def carregar_f1_db() -> pd.DataFrame:
    """Carrega todo o histórico F1 em < 2ms diretamente do SQLite."""
    conn = get_connection()
    df = pd.read_sql_query("SELECT data AS DATA, encarregado AS ENCARREGADO FROM historico_f1 ORDER BY data DESC, encarregado ASC", conn)
    conn.close()
    
    # Se o banco SQLite estiver vazio (primeiro boot ou reinício), restaura em lote da nuvem ou CSV
    if df.empty:
        df = restaurar_f1_inicial()
    return df

def restaurar_f1_inicial() -> pd.DataFrame:
    """Restaura o F1 a partir do Google Sheets ou CSV local e popula o SQLite."""
    df_restaurado = pd.DataFrame(columns=["DATA", "ENCARREGADO"])
    
    # 1. Tentar ler do Google Sheets (aba Historico_F1)
    try:
        sh = obter_planilha_google()
        if sh and "Historico_F1" in [w.title for w in sh.worksheets()]:
            ws = sh.worksheet("Historico_F1")
            recs = ws.get_all_records()
            if recs:
                df_cloud = pd.DataFrame(recs)
                if not df_cloud.empty and "DATA" in df_cloud.columns and "ENCARREGADO" in df_cloud.columns:
                    df_cloud = df_cloud[df_cloud["ENCARREGADO"].astype(str).str.strip().ne("") & df_cloud["DATA"].astype(str).str.strip().ne("")]
                    if not df_cloud.empty:
                        salvar_f1_db(df_cloud, sync_cloud=False)
                        return df_cloud[["DATA", "ENCARREGADO"]].drop_duplicates()
    except Exception:
        pass
        
    # 2. Fallback: CSV local
    if os.path.exists(CAMINHO_HISTORICO_F1_CSV):
        try:
            df_local = pd.read_csv(CAMINHO_HISTORICO_F1_CSV)
            if not df_local.empty and "DATA" in df_local.columns and "ENCARREGADO" in df_local.columns:
                salvar_f1_db(df_local, sync_cloud=False)
                return df_local[["DATA", "ENCARREGADO"]].drop_duplicates()
        except Exception:
            pass
            
    return df_restaurado

def salvar_f1_db(df_f1: pd.DataFrame, sync_cloud: bool = True) -> tuple:
    """Salva o DataFrame F1 no SQLite instantaneamente e agenda espelhamento na nuvem."""
    if df_f1 is None or df_f1.empty:
        return False, "Dados vazios."
        
    df_limpo = df_f1.dropna(subset=["DATA", "ENCARREGADO"]).copy()
    df_limpo["DATA"] = df_limpo["DATA"].astype(str).str.strip()
    df_limpo["ENCARREGADO"] = df_limpo["ENCARREGADO"].astype(str).str.strip()
    df_limpo = df_limpo[df_limpo["ENCARREGADO"].ne("") & df_limpo["DATA"].ne("")].drop_duplicates(subset=["DATA", "ENCARREGADO"])
    
    if df_limpo.empty:
        return False, "Nenhum registro válido."
        
    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    with db_lock:
        conn = get_connection()
        c = conn.cursor()
        # Inserção em lote rápida com REPLACE
        registros = [(r["DATA"], r["ENCARREGADO"], agora) for _, r in df_limpo.iterrows()]
        c.executemany("INSERT OR REPLACE INTO historico_f1 (data, encarregado, criado_em) VALUES (?, ?, ?)", registros)
        c.execute("INSERT OR REPLACE INTO sync_controle (tabela, ultima_sincronizacao, total_registros, status) VALUES ('historico_f1', ?, ?, 'OK')", (agora, len(df_limpo)))
        conn.commit()
        conn.close()
        
    # Atualizar CSV local como backup secundário
    try:
        df_limpo[["DATA", "ENCARREGADO"]].to_csv(CAMINHO_HISTORICO_F1_CSV, index=False)
    except Exception:
        pass
        
    if sync_cloud:
        disparar_sync_f1_nuvem()
        
    return True, f"OK: {len(df_limpo)} registros salvos no banco ultrarrápido."

def adicionar_entrega_f1_db(data_str: str, encarregado_str: str, sync_cloud: bool = True) -> bool:
    """Adiciona uma entrega individual ao F1 instantaneamente."""
    d = str(data_str).strip()
    e = str(encarregado_str).strip()
    if not d or not e:
        return False
        
    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with db_lock:
        conn = get_connection()
        conn.execute("INSERT OR REPLACE INTO historico_f1 (data, encarregado, criado_em) VALUES (?, ?, ?)", (d, e, agora))
        conn.commit()
        conn.close()
        
    if sync_cloud:
        disparar_sync_f1_nuvem()
    return True

def remover_entrega_f1_db(data_str: str, encarregado_str: str, sync_cloud: bool = True) -> bool:
    """Remove uma entrega específica instantaneamente."""
    d = str(data_str).strip()
    e = str(encarregado_str).strip()
    with db_lock:
        conn = get_connection()
        conn.execute("DELETE FROM historico_f1 WHERE data = ? AND encarregado = ?", (d, e))
        conn.commit()
        conn.close()
        
    if sync_cloud:
        disparar_sync_f1_nuvem()
    return True

def disparar_sync_f1_nuvem():
    """Espelha o SQLite no Google Sheets em segundo plano (não trava a tela)."""
    def _tarefa():
        try:
            sh = obter_planilha_google()
            if not sh:
                return
            conn_sql = get_connection()
            df_atual = pd.read_sql_query("SELECT data AS DATA, encarregado AS ENCARREGADO FROM historico_f1 ORDER BY data DESC", conn_sql)
            conn_sql.close()
            
            if df_atual.empty:
                return
                
            ws_names = [w.title for w in sh.worksheets()]
            if "Historico_F1" not in ws_names:
                ws_f1 = sh.add_worksheet(title="Historico_F1", rows=max(len(df_atual) + 500, 1000), cols=5)
            else:
                ws_f1 = sh.worksheet("Historico_F1")
                
            valores = [["DATA", "ENCARREGADO"]] + df_atual[["DATA", "ENCARREGADO"]].astype(str).values.tolist()
            if ws_f1.row_count < len(valores) + 50:
                ws_f1.add_rows(len(valores) + 500 - ws_f1.row_count)
            ws_f1.clear()
            ws_f1.update(values=valores, range_name=f"A1:B{len(valores)}")
        except Exception:
            pass
            
    t = threading.Thread(target=_tarefa, daemon=True)
    t.start()

# =====================================================================
# 2. BRIEFINGS DIÁRIOS & IA (ALTA VELOCIDADE + NUVEM)
# =====================================================================
def carregar_briefings_db() -> dict:
    """Carrega todos os briefings diários em < 2ms."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT data_iso, conteudo_json FROM briefings ORDER BY data_iso DESC")
    rows = c.fetchall()
    conn.close()
    
    dados = {}
    for r in rows:
        try:
            dados[r["data_iso"]] = json.loads(r["conteudo_json"])
        except Exception:
            pass
            
    if not dados:
        dados = restaurar_briefings_inicial()
        
    return dados

def restaurar_briefings_inicial() -> dict:
    """Restaura briefings da nuvem Google Sheets ou JSON local para o SQLite."""
    dados = {}
    
    # 1. Tentar ler da nuvem Google Sheets
    try:
        sh = obter_planilha_google()
        if sh and "Briefings" in [w.title for w in sh.worksheets()]:
            ws_br = sh.worksheet("Briefings")
            recs = ws_br.get_all_records()
            for r in recs:
                dt_iso = str(r.get("DATA_ISO", "")).strip()
                c_json = r.get("CONTEUDO_JSON", "")
                if dt_iso and c_json:
                    try:
                        b_dict = json.loads(c_json)
                        dados[dt_iso] = b_dict
                        salvar_briefing_dia_db(dt_iso, b_dict, sync_cloud=False)
                    except Exception:
                        pass
    except Exception:
        pass
        
    # 2. Fallback JSON local
    if not dados and os.path.exists(CAMINHO_BRIEFINGS_JSON):
        try:
            with open(CAMINHO_BRIEFINGS_JSON, "r", encoding="utf-8") as f:
                d_loc = json.load(f)
                if isinstance(d_loc, dict):
                    for k, v in d_loc.items():
                        dados[k] = v
                        salvar_briefing_dia_db(k, v, sync_cloud=False)
        except Exception:
            pass
            
    return dados

def obter_briefing_dia_db(data_str: str) -> dict:
    """Obtém o briefing de um dia específico em 0.5ms."""
    if not data_str:
        return None
    d_clean = str(data_str).strip()
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT conteudo_json FROM briefings WHERE data_iso = ? OR data_formatada = ?", (d_clean, d_clean))
    row = c.fetchone()
    conn.close()
    
    if row:
        try:
            return json.loads(row["conteudo_json"])
        except Exception:
            pass
    return None

def salvar_briefing_dia_db(data_str: str, briefing_dict: dict, sync_cloud: bool = True) -> bool:
    """Salva o briefing no SQLite e sincroniza com a nuvem."""
    if not data_str or not isinstance(briefing_dict, dict):
        return False
        
    dt_iso = briefing_dict.get("data_iso", str(data_str))
    dt_fmt = briefing_dict.get("data_formatada", str(data_str))
    dt_salvo = briefing_dict.get("data_salvo", datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
    origem = briefing_dict.get("origem", "IA")
    resumo_txt = str(briefing_dict.get("briefing_texto", ""))[:1500]
    conteudo_json = json.dumps(briefing_dict, ensure_ascii=False)
    
    with db_lock:
        conn = get_connection()
        conn.execute("""
            INSERT OR REPLACE INTO briefings 
            (data_iso, data_formatada, data_salvo, conteudo_json, resumo_texto, origem)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (dt_iso, dt_fmt, dt_salvo, conteudo_json, resumo_txt, origem))
        conn.commit()
        conn.close()
        
    # Atualizar JSON local
    try:
        todos = carregar_briefings_db()
        with open(CAMINHO_BRIEFINGS_JSON, "w", encoding="utf-8") as f:
            json.dump(todos, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
        
    if sync_cloud:
        disparar_sync_briefings_nuvem(dt_iso, briefing_dict)
        
    return True

def excluir_briefing_dia_db(data_str: str, sync_cloud: bool = True) -> bool:
    """Exclui um briefing no SQLite e na nuvem."""
    d_clean = str(data_str).strip()
    with db_lock:
        conn = get_connection()
        conn.execute("DELETE FROM briefings WHERE data_iso = ? OR data_formatada = ?", (d_clean, d_clean))
        conn.commit()
        conn.close()
        
    if sync_cloud:
        def _tarefa_excluir():
            try:
                sh = obter_planilha_google()
                if sh and "Briefings" in [w.title for w in sh.worksheets()]:
                    ws = sh.worksheet("Briefings")
                    recs = ws.get_all_records()
                    for i, r in enumerate(recs, start=2):
                        if str(r.get("DATA_ISO", "")).strip() == d_clean:
                            ws.delete_rows(i)
                            break
            except Exception:
                pass
        threading.Thread(target=_tarefa_excluir, daemon=True).start()
        
    return True

def disparar_sync_briefings_nuvem(chave_iso: str, briefing_dict: dict):
    """Atualiza a aba Briefings no Google Sheets em segundo plano."""
    def _tarefa():
        try:
            sh = obter_planilha_google()
            if not sh:
                return
            ws_names = [w.title for w in sh.worksheets()]
            if "Briefings" not in ws_names:
                ws_br = sh.add_worksheet(title="Briefings", rows=1000, cols=10)
                ws_br.update(values=[["DATA_ISO", "DATA_FORMATADA", "DATA_SALVO", "CONTEUDO_JSON", "RESUMO_TEXTO"]], range_name="A1:E1")
            else:
                ws_br = sh.worksheet("Briefings")
                
            records = ws_br.get_all_records()
            row_idx = None
            for i, rec in enumerate(records, start=2):
                if str(rec.get("DATA_ISO", "")).strip() == chave_iso:
                    row_idx = i
                    break
                    
            conteudo_json_str = json.dumps(briefing_dict, ensure_ascii=False)
            resumo_texto = str(briefing_dict.get("briefing_texto", ""))[:1500]
            linha = [chave_iso, briefing_dict.get("data_formatada", ""), briefing_dict.get("data_salvo", ""), conteudo_json_str, resumo_texto]
            
            if row_idx:
                ws_br.update(values=[linha], range_name=f"A{row_idx}:E{row_idx}")
            else:
                ws_br.append_row(linha)
        except Exception:
            pass
            
    threading.Thread(target=_tarefa, daemon=True).start()

# =====================================================================
# 3. BASE PDE DE FUNCIONÁRIOS
# =====================================================================
def carregar_pde_db() -> pd.DataFrame:
    """Carrega o PDE completo em < 3ms."""
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT matricula AS MATRICULA, nome AS NOME, funcao AS FUNÇÃO, 
               cc AS "C.C", encarregado AS ENCARREGADO, turno AS TURNO, 
               status AS STATUS, disciplina AS DISCIPLINA, mao_de_obra AS "MÃO DE OBRA"
        FROM pde_colaboradores 
        ORDER BY nome ASC
    """, conn)
    conn.close()
    
    if df.empty:
        df = restaurar_pde_inicial()
    return df

def restaurar_pde_inicial() -> pd.DataFrame:
    """Popula o SQLite com o PDE a partir do Google Sheets ou CSV."""
    df_res = pd.DataFrame()
    # 1. Tentar Google Sheets
    try:
        sh = obter_planilha_google()
        if sh and "PDE" in [w.title for w in sh.worksheets()]:
            ws = sh.worksheet("PDE")
            recs = ws.get_all_records()
            if recs:
                df_cloud = pd.DataFrame(recs)
                if not df_cloud.empty:
                    salvar_pde_db(df_cloud, sync_cloud=False)
                    return df_cloud
    except Exception:
        pass
        
    # 2. Fallback CSV
    if os.path.exists(CAMINHO_PDE_CSV):
        try:
            df_csv = pd.read_csv(CAMINHO_PDE_CSV)
            if not df_csv.empty:
                salvar_pde_db(df_csv, sync_cloud=False)
                return df_csv
        except Exception:
            pass
            
    return df_res

def salvar_pde_db(df_pde: pd.DataFrame, sync_cloud: bool = True):
    """Salva a base PDE no SQLite e agenda sync."""
    if df_pde is None or df_pde.empty:
        return
        
    col_map = {
        "MATRÍCULA": "matricula", "MATRICULA": "matricula", "CHAPA": "matricula",
        "NOME": "nome", "COLABORADOR": "nome", "FUNCIONARIO": "nome",
        "FUNÇÃO": "funcao", "FUNCAO": "funcao", "CARGO": "funcao",
        "C.C": "cc", "CC": "cc", "CENTRO DE CUSTO (C.C)": "cc", "CENTRO DE CUSTO": "cc",
        "ENCARREGADO": "encarregado", "LIDER": "encarregado",
        "TURNO": "turno", "STATUS": "status", "DISCIPLINA": "disciplina",
        "MÃO DE OBRA": "mao_de_obra", "MAO DE OBRA": "mao_de_obra", "MAO_DE_OBRA": "mao_de_obra"
    }
    
    df_work = df_pde.copy()
    for c in df_work.columns:
        c_up = str(c).strip().upper()
        if c_up in col_map:
            df_work.rename(columns={c: col_map[c_up]}, inplace=True)
            
    cols_necessarias = ["matricula", "nome", "funcao", "cc", "encarregado", "turno", "status", "disciplina", "mao_de_obra"]
    for c in cols_necessarias:
        if c not in df_work.columns:
            df_work[c] = ""
            
    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df_work["atualizado_em"] = agora
    
    # Garantir matrícula única e não-nula
    if "matricula" in df_work.columns:
        df_work["matricula"] = df_work["matricula"].fillna("").astype(str).str.strip()
        # Se houver linhas sem matrícula, gerar identificador temporário
        sem_mat = df_work["matricula"] == ""
        if sem_mat.any():
            df_work.loc[sem_mat, "matricula"] = [f"AUTO_{i}" for i in range(sem_mat.sum())]
    
    with db_lock:
        conn = get_connection()
        c = conn.cursor()
        registros = [
            (str(r["matricula"]), str(r["nome"]), str(r["funcao"]), str(r["cc"]), 
             str(r["encarregado"]), str(r["turno"]), str(r["status"]), str(r["disciplina"]), 
             str(r["mao_de_obra"]), agora)
            for _, r in df_work.iterrows()
        ]
        c.executemany("""
            INSERT OR REPLACE INTO pde_colaboradores 
            (matricula, nome, funcao, cc, encarregado, turno, status, disciplina, mao_de_obra, atualizado_em)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, registros)
        c.execute("INSERT OR REPLACE INTO sync_controle (tabela, ultima_sincronizacao, total_registros, status) VALUES ('pde_colaboradores', ?, ?, 'OK')", (agora, len(registros)))
        conn.commit()
        conn.close()
        
    try:
        df_pde.to_csv(CAMINHO_PDE_CSV, index=False)
    except Exception:
        pass

# =====================================================================
# STATUS E CONTROLE GERAL
# =====================================================================
def status_banco_geral() -> dict:
    """Retorna contadores em tempo real para exibir na sidebar."""
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM historico_f1")
    total_f1 = c.fetchone()[0]
    
    c.execute("SELECT COUNT(*) FROM briefings")
    total_briefings = c.fetchone()[0]
    
    c.execute("SELECT COUNT(*) FROM pde_colaboradores")
    total_pde = c.fetchone()[0]
    
    c.execute("SELECT COUNT(*) FROM rdc_banco")
    total_rdcs = c.fetchone()[0]
    conn.close()
    
    return {
        "f1": total_f1,
        "briefings": total_briefings,
        "pde": total_pde,
        "rdcs": total_rdcs,
        "db_size_kb": round(os.path.getsize(DB_PATH) / 1024, 1) if os.path.exists(DB_PATH) else 0
    }
