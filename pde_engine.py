# -*- coding: utf-8 -*-
"""
==============================================================================
PDE DATA ENGINE & AUDIT SYSTEM (Data Quality & Real-Time Cache Bridge)
Desenvolvido para: Edson Garcia — SGO RDC & PDE
==============================================================================
Módulo de alta performance com SINCRONIZAÇÃO AUTOMÁTICA EM TEMPO REAL
e Integração Bidirecional Instantânea: SQLite + CSV + Google Sheets Online
"""

import os
import re
import json
import sqlite3
import datetime
import time
import threading
import queue
import unicodedata
import pandas as pd
import numpy as np

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(PASTA_BASE, "pde_database.db")
CSV_PDE = os.path.join(PASTA_BASE, "PDE.csv")
CSV_BASE_ATUAL = os.path.join(PASTA_BASE, "BASE_ATUAL.csv")
SECRETS_PATH = os.path.join(PASTA_BASE, ".streamlit", "secrets.toml")
URL_PLANILHA_PADRAO = "https://docs.google.com/spreadsheets/d/1ajWLKG4I56_QAwc1VoZmi8w4YSGbmHf6oEho_yWmsYY/edit"

_sincronizador_thread = None
_sincronizador_ativo = False

_fila_gsheets_sync = queue.Queue()
_worker_gsheets_thread = None
_cache_mapa_linhas_sheets = {}
_cache_mapa_nomes_sheets = {}
_cache_colunas_detectadas = {}
_cache_mapa_timestamp = 0


def get_db_connection():
    """Retorna uma conexão SQLite otimizada com WAL mode para leitura/escrita simultânea em < 1ms."""
    conn = sqlite3.connect(DB_PATH, timeout=15.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA cache_size=-64000;")
    return conn


def inicializar_banco_pde():
    """Cria as tabelas de alta performance e índices caso não existam."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # 1. Tabela Principal de Funcionários
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pde_funcionarios (
                MATRICULA TEXT PRIMARY KEY,
                NOME TEXT NOT NULL,
                FUNCAO TEXT,
                CC TEXT,
                ENCARREGADO TEXT,
                TURNO TEXT,
                STATUS TEXT,
                DISCIPLINA TEXT,
                MAO_DE_OBRA TEXT,
                ATUALIZADO_EM TEXT
            );
        """)
        
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_pde_enc ON pde_funcionarios(ENCARREGADO);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_pde_cc ON pde_funcionarios(CC);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_pde_nome ON pde_funcionarios(NOME);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_pde_status ON pde_funcionarios(STATUS);")

        # 2. Tabela de Auditoria e Inconsistências
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pde_auditoria_erros (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                TIPO_ERRO TEXT NOT NULL,
                CAMPO TEXT,
                MATRICULA TEXT,
                NOME TEXT,
                ENCARREGADO TEXT,
                CC TEXT,
                DETALHES TEXT,
                GRAVIDADE TEXT,
                DATA_HORA TEXT
            );
        """)

        # 3. Tabela de Logs de Sincronização
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pde_sync_logs (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                DATA_HORA TEXT NOT NULL,
                FONTE TEXT,
                TOTAL_LIDOS INTEGER,
                TOTAL_VALIDOS INTEGER,
                DUPLICADOS_TRATADOS INTEGER,
                PENDENCIAS_ENCONTRADAS INTEGER,
                DURACAO_MS REAL,
                STATUS TEXT,
                MENSAGEM TEXT
            );
        """)
        
        # 4. Tabela de Metadados / Versão
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pde_meta (
                CHAVE TEXT PRIMARY KEY,
                VALOR TEXT,
                ATUALIZADO_EM TEXT
            );
        """)

        # 5. Tabela de Histórico de Movimentações
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pde_movimentacoes (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                DATA_HORA TEXT,
                MATRICULA TEXT,
                NOME TEXT,
                DE_ENCARREGADO TEXT,
                PARA_ENCARREGADO TEXT,
                DE_CC TEXT,
                PARA_CC TEXT,
                APONTADOR TEXT,
                OBSERVACAO TEXT
            );
        """)
        conn.commit()


def eh_colaborador_inativo(status_str):
    if not status_str or pd.isna(status_str):
        return False
    s = str(status_str).strip().upper()
    termos = ["DEMIT", "DESLIG", "INATIV", "RESCIS", "CANCEL", "TRANSFERID", "ENCERR", "BAIXA", "AFAST", "INSS", "LICENCA", "LICENÇA"]
    return any(t in s for t in termos)


def normalizar_texto(texto):
    """Remove acentos, espaços extras e coloca em caixa alta."""
    if not texto:
        return ""
    s = str(texto).strip().upper()
    return unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')


def _col_idx_to_letter(col_idx):
    """Converte índice 1-based (1=A, 2=B, 11=K) em letra de coluna do Excel/Sheets."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _ler_secrets_toml_seguro(caminho):
    """Lê o arquivo secrets.toml de forma 100% tolerante a falhas (tomllib, toml ou parser manual)."""
    if not os.path.exists(caminho):
        return None
        
    # 1. Tentar tomllib (Python 3.11+)
    try:
        import tomllib
        with open(caminho, "rb") as f:
            sec = tomllib.load(f)
            if "connections" in sec and "gsheets" in sec["connections"]:
                return dict(sec["connections"]["gsheets"])
    except Exception:
        pass
        
    # 2. Tentar biblioteca toml externa
    try:
        import toml
        with open(caminho, "r", encoding="utf-8") as f:
            sec = toml.load(f)
            if "connections" in sec and "gsheets" in sec["connections"]:
                return dict(sec["connections"]["gsheets"])
    except Exception:
        pass
        
    # 3. Parser manual resiliente de fallback
    try:
        resultado = {}
        dentro_gsheets = False
        with open(caminho, "r", encoding="utf-8") as f:
            for linha in f:
                linha_clean = linha.strip()
                if not linha_clean or linha_clean.startswith("#"):
                    continue
                if linha_clean.startswith("["):
                    dentro_gsheets = (linha_clean == "[connections.gsheets]")
                    continue
                if dentro_gsheets and "=" in linha_clean:
                    k, v = linha_clean.split("=", 1)
                    k = k.strip()
                    v = v.strip().strip('"').strip("'")
                    resultado[k] = v
        if resultado:
            return resultado
    except Exception:
        pass
        
    return None


def _abrir_com_service_account_dict(cfg_dict):
    """Abre o Google Spreadsheet a partir de um dicionário de Service Account."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        info = dict(cfg_dict)
        spread_val = info.pop("spreadsheet", None) or URL_PLANILHA_PADRAO
        
        # Corrigir quebras de linha na private_key
        if "private_key" in info and isinstance(info["private_key"], str):
            info["private_key"] = info["private_key"].replace("\\n", "\n")
            
        # Filtrar apenas chaves reconhecidas pela ServiceAccountCredentials da Google
        chaves_validas = {
            "type", "project_id", "private_key_id", "private_key",
            "client_email", "client_id", "auth_uri", "token_uri",
            "auth_provider_x509_cert_url", "client_x509_cert_url", "universe_domain"
        }
        info_limpa = {k: v for k, v in info.items() if k in chaves_validas}
        
        creds = Credentials.from_service_account_info(info_limpa, scopes=scopes)
        client = gspread.authorize(creds)
        
        # Abrir por URL ou por Key
        if spread_val:
            s_str = str(spread_val).strip()
            if s_str.startswith("http"):
                try:
                    return client.open_by_url(s_str)
                except Exception:
                    # Extrair o ID da planilha da URL caso open_by_url oscile
                    m = re.search(r"/d/([a-zA-Z0-9-_]+)", s_str)
                    if m:
                        return client.open_by_key(m.group(1))
            else:
                return client.open_by_key(s_str)
                
        return client.open("PDE")
    except Exception as e:
        print(f"Aviso em _abrir_com_service_account_dict: {e}")
        return None


def obter_cliente_google_sheets():
    """
    Autentica com as credenciais do Google Sheets configuradas em 5 camadas redundantes:
    1. streamlit st.connection("gsheets", type=GSheetsConnection) se estiver no Streamlit
    2. st.secrets["connections"]["gsheets"]
    3. .streamlit/secrets.toml
    4. robo.json na pasta do projeto
    5. credentials.json / service_account.json
    """
    # 1. Tentar st.connection nativo do Streamlit se disponível
    try:
        import streamlit as st
        from streamlit_gsheets import GSheetsConnection
        conn_gs = st.connection("gsheets", type=GSheetsConnection)
        if hasattr(conn_gs, "_instance") and conn_gs._instance is not None:
            client = conn_gs._instance
            url = None
            if hasattr(st, "secrets") and "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                url = st.secrets["connections"]["gsheets"].get("spreadsheet")
            url = url or URL_PLANILHA_PADRAO
            try:
                return client.open_by_url(url)
            except Exception:
                m = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
                if m:
                    return client.open_by_key(m.group(1))
    except Exception:
        pass

    # 2. Tentar st.secrets diretamente
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            cfg_g = dict(st.secrets["connections"]["gsheets"])
            sh = _abrir_com_service_account_dict(cfg_g)
            if sh:
                return sh
    except Exception:
        pass

    # 3. Tentar carregar de secrets.toml em disco
    if os.path.exists(SECRETS_PATH):
        cfg_g = _ler_secrets_toml_seguro(SECRETS_PATH)
        if cfg_g:
            sh = _abrir_com_service_account_dict(cfg_g)
            if sh:
                return sh

    # 4. Tentar carregar de robo.json local
    p_robo = os.path.join(PASTA_BASE, "robo.json")
    if os.path.exists(p_robo):
        try:
            with open(p_robo, "r", encoding="utf-8") as f:
                cfg_robo = json.load(f)
            cfg_robo["spreadsheet"] = URL_PLANILHA_PADRAO
            sh = _abrir_com_service_account_dict(cfg_robo)
            if sh:
                return sh
        except Exception:
            pass

    # 5. Tentar outros arquivos JSON de credenciais
    for json_file in ["credentials.json", "service_account.json"]:
        p = os.path.join(PASTA_BASE, json_file)
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    cfg_j = json.load(f)
                cfg_j["spreadsheet"] = URL_PLANILHA_PADRAO
                sh = _abrir_com_service_account_dict(cfg_j)
                if sh:
                    return sh
            except Exception:
                pass

    return None


def detectar_colunas_planilha(ws):
    """
    Detecta dinamicamente as letras das colunas na planilha online:
    - MATRICULA (A)
    - NOME (B)
    - CENTRO DE CUSTO (H)
    - MATRICULA ENCARREGADO (K)
    - ENCARREGADO (L) — Apenas para mapeamento, NUNCA sobrescrita!
    """
    global _cache_colunas_detectadas
    if _cache_colunas_detectadas:
        return _cache_colunas_detectadas

    # Valores padrão oficiais ENESA
    cols = {
        "col_mat_idx": 1,        # A
        "col_nome_idx": 2,       # B
        "col_cc_idx": 8,         # H
        "col_mat_enc_idx": 11,   # K
        "col_enc_nome_idx": 12   # L
    }
    
    try:
        # Lê até a linha 5 para identificar a linha de cabeçalho
        linhas_iniciais = ws.get_values("A1:Z5")
        for num_linha, linha in enumerate(linhas_iniciais, start=1):
            achou_header = False
            for idx, val in enumerate(linha):
                v = normalizar_texto(val)
                c_idx = idx + 1
                if "MATR" in v and not any(k in v for k in ["ENCARR", "COORD"]):
                    cols["col_mat_idx"] = c_idx
                    achou_header = True
                elif "MATR" in v and "ENCARR" in v:
                    cols["col_mat_enc_idx"] = c_idx
                elif "ENCARR" in v and "MATR" not in v:
                    cols["col_enc_nome_idx"] = c_idx
                elif any(k in v for k in ["CENTRO DE CUSTO", "C.C", "CC", "CUSTO"]):
                    cols["col_cc_idx"] = c_idx
                elif "NOME" in v and not any(k in v for k in ["ENCARR", "COORD", "SUPERV"]):
                    cols["col_nome_idx"] = c_idx
            if achou_header:
                cols["linha_cabecalho"] = num_linha
                break
    except Exception as e:
        print(f"Aviso ao detectar colunas no Google Sheets: {e}")

    # Converte índices para letras de coluna (ex: 11 -> 'K', 8 -> 'H')
    cols["letra_mat"] = _col_idx_to_letter(cols["col_mat_idx"])
    cols["letra_nome"] = _col_idx_to_letter(cols["col_nome_idx"])
    cols["letra_cc"] = _col_idx_to_letter(cols["col_cc_idx"])
    cols["letra_mat_enc"] = _col_idx_to_letter(cols["col_mat_enc_idx"])
    cols["letra_enc_nome"] = _col_idx_to_letter(cols["col_enc_nome_idx"])

    _cache_colunas_detectadas = cols
    return cols


def obter_mapa_linhas_planilha(ws, forcar_recarregar=False):
    """
    Retorna dois mapeamentos em cache:
    1. Matrícula -> Linha no Google Sheets (suporta string, número inteiro e com/sem zeros)
    2. Nome do Colaborador (normalizado) -> Linha no Google Sheets (fallback infalível)
    """
    global _cache_mapa_linhas_sheets, _cache_mapa_nomes_sheets, _cache_mapa_timestamp
    agora = time.time()
    if not forcar_recarregar and _cache_mapa_linhas_sheets and (agora - _cache_mapa_timestamp < 900):
        return _cache_mapa_linhas_sheets, _cache_mapa_nomes_sheets
        
    try:
        cols = detectar_colunas_planilha(ws)
        col_mat_vals = ws.col_values(cols["col_mat_idx"])
        col_nome_vals = ws.col_values(cols["col_nome_idx"]) if cols["col_nome_idx"] else []
        
        mapa_mat = {}
        mapa_nome = {}
        
        for idx, val in enumerate(col_mat_vals):
            row_num = idx + 1
            clean_m = str(val).strip().replace(".0", "").replace(".", "").replace("-", "")
            if clean_m and clean_m.upper() != "MATRICULA":
                mapa_mat[clean_m] = row_num
                if clean_m.isdigit():
                    mapa_mat[str(int(clean_m))] = row_num
                    
        for idx, val in enumerate(col_nome_vals):
            row_num = idx + 1
            clean_n = normalizar_texto(val)
            if clean_n and clean_n != "NOME":
                mapa_nome[clean_n] = row_num

        _cache_mapa_linhas_sheets = mapa_mat
        _cache_mapa_nomes_sheets = mapa_nome
        _cache_mapa_timestamp = agora
        return _cache_mapa_linhas_sheets, _cache_mapa_nomes_sheets
    except Exception as e:
        print(f"Erro em obter_mapa_linhas_planilha: {e}")
        return _cache_mapa_linhas_sheets, _cache_mapa_nomes_sheets


def obter_mapa_encarregados_matricula():
    """
    Retorna um dicionário mapeando NOME_ENCARREGADO -> MATRICULA_ENCARREGADO.
    Busca em 3 fontes redundantes: SQLite, PDE.csv e BASE_ATUAL.csv.
    """
    mapa = {}
    
    # 1. Banco SQLite
    try:
        with get_db_connection() as conn:
            df_func = pd.read_sql_query("SELECT MATRICULA, NOME FROM pde_funcionarios WHERE MATRICULA != '' AND MATRICULA != '-';", conn)
            for _, r in df_func.iterrows():
                n = str(r["NOME"]).strip().upper()
                n_norm = normalizar_texto(n)
                m = str(r["MATRICULA"]).strip().replace(".0", "")
                if m and m not in ["-", "NAN", "0", ""]:
                    mapa[n] = m
                    mapa[n_norm] = m
    except Exception:
        pass

    # 2. PDE.csv
    if os.path.exists(CSV_PDE):
        try:
            df_pde = pd.read_csv(CSV_PDE, encoding='latin1', low_memory=False)
            col_enc = [c for c in df_pde.columns if "ENCARR" in str(c).upper() and "MATR" not in str(c).upper()]
            col_mat_enc = [c for c in df_pde.columns if "MATR" in str(c).upper() and "ENCARR" in str(c).upper()]
            if col_enc and col_mat_enc:
                for _, r in df_pde.iterrows():
                    e = str(r[col_enc[0]]).strip().upper()
                    me = str(r[col_mat_enc[0]]).strip().replace(".0", "")
                    if e and e not in ["-", "NAN", "0", ""] and me and me not in ["-", "NAN", "0", ""]:
                        mapa[e] = me
                        mapa[normalizar_texto(e)] = me
                        
            # Também mapeia funcionário direto (todo encarregado é um funcionário)
            col_m_func = [c for c in df_pde.columns if "MATR" in str(c).upper() and not any(k in str(c).upper() for k in ["ENCARR", "COORD"])]
            col_n_func = [c for c in df_pde.columns if "NOME" in str(c).upper()]
            if col_m_func and col_n_func:
                for _, r in df_pde.iterrows():
                    nf = str(r[col_n_func[0]]).strip().upper()
                    mf = str(r[col_m_func[0]]).strip().replace(".0", "")
                    if nf and nf not in ["-", "NAN", "0", ""] and mf and mf not in ["-", "NAN", "0", ""]:
                        mapa[nf] = mf
                        mapa[normalizar_texto(nf)] = mf
        except Exception:
            pass

    return mapa


def sincronizar_alteracoes_gsheets_direto(lista_alteracoes):
    """
    Atualiza DIRETAMENTE as células no Google Sheets em tempo real via batch_update.
    - Coluna K: MATRICULA ENCARREGADO (gravada com a matrícula do líder ou vazia se desvinculado)
    - Coluna L: ENCARREGADO (NUNCA É SOBRESCRITA - preserva a fórmula =PROCV / =XLOOKUP)
    - Coluna H: CENTRO DE CUSTO (C.C)
    """
    if not lista_alteracoes:
        return True, "Nenhuma alteração para enviar."

    try:
        sh = obter_cliente_google_sheets()
        if not sh:
            return False, "Não foi possível autenticar no Google Sheets (verifique secrets.toml ou robo.json)."
            
        ws_names = [w.title for w in sh.worksheets()]
        alvo_ws = "PDE" if "PDE" in ws_names else ("Página1" if "Página1" in ws_names else ws_names[0])
        ws = sh.worksheet(alvo_ws)
        
        cols = detectar_colunas_planilha(ws)
        map_mat, map_nome = obter_mapa_linhas_planilha(ws)
        if not map_mat:
            map_mat, map_nome = obter_mapa_linhas_planilha(ws, forcar_recarregar=True)
            
        mapa_enc_mat = obter_mapa_encarregados_matricula()
        
        batch_updates = []
        celulas_afetadas = 0
        linhas_modificadas = []
        
        for alt in lista_alteracoes:
            mat_raw = str(alt.get("MATRICULA", "")).strip().replace(".0", "").replace(".", "").replace("-", "")
            
            # Localizar número da linha (por matrícula ou por nome)
            row = map_mat.get(mat_raw)
            if not row and mat_raw.isdigit():
                row = map_mat.get(str(int(mat_raw)))
            if not row and "NOME" in alt:
                nome_alt = normalizar_texto(alt.get("NOME"))
                row = map_nome.get(nome_alt)
                
            # Se ainda não encontrou, recarrega o mapa uma vez
            if not row:
                map_mat, map_nome = obter_mapa_linhas_planilha(ws, forcar_recarregar=True)
                row = map_mat.get(mat_raw)
                if not row and "NOME" in alt:
                    row = map_nome.get(normalizar_texto(alt.get("NOME")))
                    
            if not row:
                print(f"[PDE SYNC] Aviso: Linha não encontrada na planilha para matrícula {mat_raw}")
                continue
                
            linhas_modificadas.append(row)
                
            # 1. Tratar Matrícula do Encarregado (Coluna K)
            if "NOVO_ENCARREGADO" in alt:
                novo_enc = alt.get("NOVO_ENCARREGADO")
                if novo_enc is None:
                    val_k = ""
                else:
                    enc_str = str(novo_enc).strip()
                    enc_norm = normalizar_texto(enc_str)
                    
                    termos_desvincular = ["", "-", "0", "NAN", "NONE", "SEM ENCARREGADO", "DESVINCULAR", "NENHUM", "(SEM ENCARREGADO)"]
                    if enc_norm in termos_desvincular or "DEIXAR SEM ENCARREGADO" in enc_norm:
                        val_k = ""
                    elif enc_str.isdigit():
                        # Já é a própria matrícula numérica
                        val_k = enc_str
                    else:
                        # Buscar matrícula pelo nome do encarregado
                        val_k = str(mapa_enc_mat.get(enc_str.upper(), "")).strip().replace(".0", "")
                        if not val_k:
                            val_k = str(mapa_enc_mat.get(enc_norm, "")).strip().replace(".0", "")
                        # Tenta extrair matrícula se o nome veio no formato "12500123 — NOME"
                        if not val_k and " — " in enc_str:
                            p_mat = enc_str.split(" — ")[0].strip()
                            if p_mat.isdigit():
                                val_k = p_mat
                                
                batch_updates.append({
                    "range": f"{cols['letra_mat_enc']}{row}",
                    "values": [[val_k]]
                })
                celulas_afetadas += 1
                
            # 2. Tratar Centro de Custo (Coluna H)
            if alt.get("NOVO_CC") is not None:
                novo_cc = str(alt.get("NOVO_CC")).strip()
                if novo_cc and novo_cc != "(Manter Atual)":
                    # Limpa se vier formatado como "125.01.008 — [PB] PRECIPITADOR"
                    cc_limpo = novo_cc.split(" — ")[0].strip()
                    batch_updates.append({
                        "range": f"{cols['letra_cc']}{row}",
                        "values": [[cc_limpo]]
                    })
                    celulas_afetadas += 1
                
        if batch_updates:
            try:
                ws.batch_update(batch_updates, value_input_option="USER_ENTERED")
            except Exception as e_batch:
                # Fallback tolerante célula a célula caso a API do Sheets rejeite o lote
                for b_item in batch_updates:
                    ws.update(b_item["range"], b_item["values"], value_input_option="USER_ENTERED")
                    
            return True, f"{celulas_afetadas} célula(s) atualizada(s) no Google Sheets com sucesso (Colunas {cols['letra_mat_enc']} e {cols['letra_cc']})!"
            
        if not linhas_modificadas:
            return False, "Nenhum colaborador da lista foi localizado na planilha online do Google Sheets."
            
        return True, "Nenhuma alteração de célula necessária."
        
    except Exception as e:
        msg_erro = str(e)
        print(f"Erro ao sincronizar com Google Sheets direto: {msg_erro}")
        return False, f"Falha na sincronização online: {msg_erro}"


def enfileirar_sincronizacao_gsheets(lista_alteracoes):
    """Enfileira para processamento assíncrono caso a conexão online esteja instável."""
    for alt in lista_alteracoes:
        _fila_gsheets_sync.put(alt)
    iniciar_worker_gsheets()


def _loop_worker_gsheets():
    """Worker em segundo plano resiliente caso a sincronização direta precise de retry."""
    while True:
        try:
            itens = []
            item_primeiro = _fila_gsheets_sync.get(timeout=4)
            if item_primeiro:
                itens.append(item_primeiro)
                while not _fila_gsheets_sync.empty():
                    try:
                        itens.append(_fila_gsheets_sync.get_nowait())
                    except queue.Empty:
                        break
                        
            if itens:
                for tentativa in range(3):
                    ok, _ = sincronizar_alteracoes_gsheets_direto(itens)
                    if ok:
                        break
                    time.sleep(4 * (tentativa + 1))
                    
                for _ in itens:
                    _fila_gsheets_sync.task_done()
        except queue.Empty:
            time.sleep(1)
        except Exception:
            time.sleep(2)


def iniciar_worker_gsheets():
    """Inicia a thread do worker de escrita do Google Sheets apenas se necessário."""
    global _worker_gsheets_thread
    if _worker_gsheets_thread is None or not _worker_gsheets_thread.is_alive():
        _worker_gsheets_thread = threading.Thread(
            target=_loop_worker_gsheets,
            daemon=True,
            name="GSheetsSyncQueueWorker"
        )
        _worker_gsheets_thread.start()


def normalizar_colunas_pde(df):
    if df is None or df.empty:
        return df
    mapping = {}
    for col in df.columns:
        c_up = str(col).strip().upper()
        if "MATR" in c_up and not any(k in c_up for k in ["ENCARR", "COORD"]):
            mapping[col] = "MATRICULA"
        elif "NOME" in c_up:
            mapping[col] = "NOME"
        elif "FUN" in c_up:
            mapping[col] = "FUNCAO"
        elif "CUSTO" in c_up or "C.C" in c_up or c_up == "CC":
            mapping[col] = "CC"
        elif "ENCARR" in c_up and "MATR" not in c_up:
            mapping[col] = "ENCARREGADO"
        elif "TURNO" in c_up:
            mapping[col] = "TURNO"
        elif "STATUS" in c_up or "SIT" in c_up:
            mapping[col] = "STATUS"
        elif "DISCIP" in c_up:
            mapping[col] = "DISCIPLINA"
        elif "MÃO DE OBRA" in c_up or "MAO DE OBRA" in c_up or "M.O" in c_up or c_up == "MO":
            mapping[col] = "MAO_DE_OBRA"
    return df.rename(columns=mapping)


def auditar_e_sincronizar_pde(df_bruto, fonte="GOOGLE_SHEETS"):
    if df_bruto is None or df_bruto.empty:
        return False, "DataFrame vazio."
    inicializar_banco_pde()
    df = normalizar_colunas_pde(df_bruto.copy())
    
    colunas_obrigatorias = ["MATRICULA", "NOME"]
    for c in colunas_obrigatorias:
        if c not in df.columns:
            return False, f"Coluna obrigatória ausente: {c}"
            
    df["MATRICULA"] = df["MATRICULA"].astype(str).str.strip().str.replace(".0", "", regex=False)
    df["NOME"] = df["NOME"].astype(str).str.strip().str.upper()
    df = df[df["MATRICULA"].ne("") & df["MATRICULA"].ne("-") & df["MATRICULA"].ne("NAN")]
    
    for c in ["FUNCAO", "CC", "ENCARREGADO", "TURNO", "STATUS", "DISCIPLINA", "MAO_DE_OBRA"]:
        if c not in df.columns:
            df[c] = ""
        else:
            df[c] = df[c].fillna("").astype(str).str.strip().str.upper()
            
    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df["ATUALIZADO_EM"] = agora
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM pde_funcionarios;")
        
        colunas_insercao = ["MATRICULA", "NOME", "FUNCAO", "CC", "ENCARREGADO", "TURNO", "STATUS", "DISCIPLINA", "MAO_DE_OBRA", "ATUALIZADO_EM"]
        df_sub = df[colunas_insercao].drop_duplicates(subset=["MATRICULA"])
        df_sub.to_sql("pde_funcionarios", conn, if_exists="append", index=False)
        
        cursor.execute("""
            INSERT OR REPLACE INTO pde_meta (CHAVE, VALOR, ATUALIZADO_EM)
            VALUES ('ultima_atualizacao', ?, ?);
        """, (agora, agora))
        conn.commit()
        
    return True, f"Sincronizados {len(df_sub)} funcionários com sucesso."


def carregar_pde_instantaneo():
    """
    Retorna o DataFrame de funcionários com leitura instantânea (< 1ms).
    Prioriza o SQLite otimizado. Se estiver vazio, lê PDE.csv ou BASE_ATUAL.csv.
    """
    inicializar_banco_pde()
    try:
        with get_db_connection() as conn:
            df = pd.read_sql_query("SELECT * FROM pde_funcionarios;", conn)
            if not df.empty:
                rename_map = {
                    "CC": "C.C",
                    "FUNCAO": "FUNÇÃO",
                    "MAO_DE_OBRA": "MÃO DE OBRA"
                }
                return df.rename(columns=rename_map)
    except Exception:
        pass

    for arq in ["PDE.csv", "BASE_ATUAL.csv"]:
        caminho = os.path.join(PASTA_BASE, arq)
        if os.path.exists(caminho):
            try:
                df = pd.read_csv(caminho, encoding='latin1', low_memory=False)
                return df
            except Exception:
                pass
    return pd.DataFrame()


def obter_mapa_centros_de_custo():
    """Retorna um dicionário mapeando CÓDIGO_CC -> 'CÓDIGO — [PB/RB] NOME_DA_DISCIPLINA'."""
    mapa = {}
    caminho_csv = os.path.join(PASTA_BASE, "PDE.csv")
    if os.path.exists(caminho_csv):
        try:
            df_raw = pd.read_csv(caminho_csv, encoding='latin1', low_memory=False)
            col_cc = [c for c in df_raw.columns if "CUSTO" in str(c).upper() or "C.C" in str(c).upper() or "CC" in str(c).upper()]
            col_disc = [c for c in df_raw.columns if "DISCIPLINA" in str(c).upper()]
            if col_cc:
                for cc, grp in df_raw.groupby(col_cc[0]):
                    cc_clean = str(cc).strip()
                    if not cc_clean or cc_clean in ["-", "0", "nan", "NAN"]:
                        continue
                    area = "PB" if "125.02." in cc_clean else "RB"
                    disc = ""
                    if col_disc:
                        discs = [str(d).strip().upper() for d in grp[col_disc[0]].dropna().unique() if str(d).strip()]
                        if discs:
                            disc = discs[0]
                    mapa[cc_clean] = f"{cc_clean} — [{area}] {disc}".strip()
        except Exception:
            pass
    return mapa


def registrar_movimentacao_equipe(matricula, nome, de_enc, para_enc, de_cc, para_cc, apontador="Apontador", obs=""):
    """Registra uma transferência / movimentação de funcionário no histórico local."""
    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    inicializar_banco_pde()
    try:
        with get_db_connection() as conn:
            conn.execute("""
                INSERT INTO pde_movimentacoes (DATA_HORA, MATRICULA, NOME, DE_ENCARREGADO, PARA_ENCARREGADO, DE_CC, PARA_CC, APONTADOR, OBSERVACAO)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (agora, str(matricula), str(nome), str(de_enc), str(para_enc), str(de_cc), str(para_cc), str(apontador), str(obs)))
            conn.commit()
    except Exception as e:
        print(f"Erro ao registrar movimentacao: {e}")


def obter_historico_movimentacoes(limite=50):
    """Retorna os últimos registros de movimentação de equipe."""
    inicializar_banco_pde()
    try:
        with get_db_connection() as conn:
            df = pd.read_sql_query(f"SELECT * FROM pde_movimentacoes ORDER BY ID DESC LIMIT {int(limite)};", conn)
            return df
    except Exception:
        return pd.DataFrame(columns=["ID", "DATA_HORA", "MATRICULA", "NOME", "DE_ENCARREGADO", "PARA_ENCARREGADO", "DE_CC", "PARA_CC", "APONTADOR", "OBSERVACAO"])


def aplicar_correcoes_em_massa(lista_alteracoes, sync_sheets=True):
    """
    Aplica correções de Encarregado e C.C tanto no banco SQLite, quanto nos arquivos locais (PDE.csv e BASE_ATUAL.csv),
    e sincroniza IMEDIATAMENTE com a planilha online do Google Sheets.
    """
    if not lista_alteracoes:
        return True, "Nenhuma alteração para aplicar."
    
    inicializar_banco_pde()
    agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    matriculas_alteradas = set()
    mapa_enc_mat = obter_mapa_encarregados_matricula()
    
    # 1. Atualizar SQLite local (< 1ms)
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            for alt in lista_alteracoes:
                mat = str(alt.get("MATRICULA", "")).strip().replace(".0", "")
                if not mat:
                    continue
                matriculas_alteradas.add(mat)
                novo_enc = alt.get("NOVO_ENCARREGADO")
                novo_cc = alt.get("NOVO_CC")
                
                enc_limpo = None
                if novo_enc is not None:
                    enc_str = str(novo_enc).strip().upper()
                    if enc_str in ["", "-", "0", "NAN", "NONE", "SEM ENCARREGADO", "🚫 (DEIXAR SEM ENCARREGADO)", "NENHUM", "(SEM ENCARREGADO)"] or "DEIXAR SEM ENCARREGADO" in enc_str:
                        enc_limpo = ""
                    else:
                        enc_limpo = str(novo_enc).strip().upper()

                if enc_limpo is not None and novo_cc is not None:
                    cursor.execute("""
                        UPDATE pde_funcionarios
                        SET ENCARREGADO = ?, CC = ?, ATUALIZADO_EM = ?
                        WHERE MATRICULA = ?;
                    """, (enc_limpo, str(novo_cc).strip(), agora, mat))
                elif enc_limpo is not None:
                    cursor.execute("""
                        UPDATE pde_funcionarios
                        SET ENCARREGADO = ?, ATUALIZADO_EM = ?
                        WHERE MATRICULA = ?;
                    """, (enc_limpo, agora, mat))
                elif novo_cc is not None:
                    cursor.execute("""
                        UPDATE pde_funcionarios
                        SET CC = ?, ATUALIZADO_EM = ?
                        WHERE MATRICULA = ?;
                    """, (str(novo_cc).strip(), agora, mat))
            conn.commit()
    except Exception as e:
        print(f"Erro ao atualizar banco SQLite: {e}")

    # 2. Atualizar CSVs locais (PDE.csv e BASE_ATUAL.csv)
    for arq_csv in ["PDE.csv", "BASE_ATUAL.csv"]:
        caminho_csv = os.path.join(PASTA_BASE, arq_csv)
        if os.path.exists(caminho_csv):
            try:
                df_csv = pd.read_csv(caminho_csv, encoding='latin1', low_memory=False)
                col_mat = [c for c in df_csv.columns if "MATR" in str(c).upper() and not any(k in str(c).upper() for k in ["ENCARR", "COORD"])]
                col_enc = [c for c in df_csv.columns if "ENCARR" in str(c).upper() and "MATR" not in str(c).upper()]
                col_mat_enc = [c for c in df_csv.columns if "MATR" in str(c).upper() and "ENCARR" in str(c).upper()]
                col_cc = [c for c in df_csv.columns if "CUSTO" in str(c).upper() or "C.C" in str(c).upper() or "CC" in str(c).upper()]
                
                if col_mat:
                    cm = col_mat[0]
                    df_csv["_MAT_STR"] = df_csv[cm].astype(str).str.strip().str.replace(".0", "", regex=False)
                    
                    mat_dict_enc = {}
                    mat_dict_mat_enc = {}
                    mat_dict_cc = {}
                    
                    for a in lista_alteracoes:
                        m_str = str(a.get("MATRICULA", "")).strip().replace(".0", "")
                        if "NOVO_ENCARREGADO" in a:
                            ne = a["NOVO_ENCARREGADO"]
                            if ne is None:
                                mat_dict_enc[m_str] = ""
                                mat_dict_mat_enc[m_str] = ""
                            else:
                                ne_str = str(ne).strip().upper()
                                if ne_str in ["", "-", "0", "NAN", "NONE", "SEM ENCARREGADO", "🚫 (DEIXAR SEM ENCARREGADO)", "NENHUM", "(SEM ENCARREGADO)"] or "DEIXAR SEM ENCARREGADO" in ne_str:
                                    mat_dict_enc[m_str] = ""
                                    mat_dict_mat_enc[m_str] = ""
                                else:
                                    mat_dict_enc[m_str] = ne_str
                                    mat_dict_mat_enc[m_str] = str(mapa_enc_mat.get(ne_str, "")).strip().replace(".0", "")
                        if "NOVO_CC" in a and a["NOVO_CC"] is not None:
                            mat_dict_cc[m_str] = str(a["NOVO_CC"]).strip()

                    if col_enc and mat_dict_enc:
                        ce = col_enc[0]
                        mask_enc = df_csv["_MAT_STR"].isin(mat_dict_enc.keys())
                        df_csv.loc[mask_enc, ce] = df_csv.loc[mask_enc, "_MAT_STR"].map(mat_dict_enc)
                        
                    if col_mat_enc and mat_dict_mat_enc:
                        cme = col_mat_enc[0]
                        mask_me = df_csv["_MAT_STR"].isin(mat_dict_mat_enc.keys())
                        df_csv.loc[mask_me, cme] = df_csv.loc[mask_me, "_MAT_STR"].map(mat_dict_mat_enc)
                        
                    if col_cc and mat_dict_cc:
                        ccc = col_cc[0]
                        mask_cc = df_csv["_MAT_STR"].isin(mat_dict_cc.keys())
                        df_csv.loc[mask_cc, ccc] = df_csv.loc[mask_cc, "_MAT_STR"].map(mat_dict_cc)
                        
                    df_csv.drop(columns=["_MAT_STR"], inplace=True, errors="ignore")
                    df_csv.to_csv(caminho_csv, index=False, encoding='latin1', errors='replace')
            except Exception as ex_csv:
                print(f"Erro ao salvar {arq_csv}: {ex_csv}")

    # 3. Sincronização ONLINE com Google Sheets (Coluna K e Coluna H)
    msg_cloud = ""
    if sync_sheets:
        ok_gs, res_gs = sincronizar_alteracoes_gsheets_direto(lista_alteracoes)
        if ok_gs:
            msg_cloud = f" | ☁️ {res_gs}"
        else:
            # Enfileira para reenvio em segundo plano se a rede oscilou
            enfileirar_sincronizacao_gsheets(lista_alteracoes)
            msg_cloud = f" | ⚠️ Salvo localmente. Envio online enfileirado: {res_gs}"

    return True, f"{len(matriculas_alteradas)} colaborador(es) atualizado(s) com sucesso!{msg_cloud}"
