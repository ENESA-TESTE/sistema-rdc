from __future__ import annotations
import json, os
from datetime import datetime, timezone
from typing import Any
import requests
import streamlit as st

def _secret(nome: str, padrao: str = "") -> str:
    try:
        valor = st.secrets.get(nome, "")
        if valor:
            return str(valor).strip()
    except Exception:
        pass
    return str(os.getenv(nome, padrao)).strip()

def configurado() -> bool:
    return bool(_secret("SUPABASE_URL") and _secret("SUPABASE_SERVICE_ROLE_KEY"))

def contrato_id() -> str:
    return _secret("SGO_CONTRATO_ID", "contrato-piloto")

def _headers(prefer: str | None = None) -> dict[str, str]:
    key = _secret("SUPABASE_SERVICE_ROLE_KEY")
    headers = {"apikey": key, "Content-Type": "application/json"}
    # sb_secret_* e uma chave opaca nova: use no cabecalho apikey.
    # service_role legado e JWT: tambem requer Authorization Bearer.
    if not key.startswith("sb_secret_"):
        headers["Authorization"] = f"Bearer {key}"
    if prefer:
        headers["Prefer"] = prefer
    return headers

def _url() -> str:
    return _secret("SUPABASE_URL").rstrip("/") + "/rest/v1/sgo_estado"

def salvar_estado(chave: str, valor: Any, timeout: int = 20) -> tuple[bool, str]:
    if not configurado():
        return False, "Supabase nao configurado."
    payload = {"contrato_id": contrato_id(), "chave": chave, "valor": valor,
               "atualizado_em": datetime.now(timezone.utc).isoformat()}
    try:
        resp = requests.post(_url()+"?on_conflict=contrato_id,chave",
            headers=_headers("resolution=merge-duplicates,return=minimal"),
            data=json.dumps(payload,ensure_ascii=False,default=str).encode("utf-8"),timeout=timeout)
        resp.raise_for_status()
        return True, "Salvo no banco persistente."
    except Exception as exc:
        return False, f"Falha no banco persistente: {exc}"

def carregar_estado(chave: str, padrao: Any = None, timeout: int = 20) -> Any:
    if not configurado():
        return padrao
    params={"contrato_id":f"eq.{contrato_id()}","chave":f"eq.{chave}","select":"valor,atualizado_em","limit":"1"}
    try:
        resp=requests.get(_url(),headers=_headers(),params=params,timeout=timeout)
        resp.raise_for_status(); dados=resp.json()
        return dados[0].get("valor",padrao) if dados else padrao
    except Exception:
        return padrao

def apagar_estado(chave: str, timeout: int = 20) -> tuple[bool, str]:
    if not configurado():
        return False, "Supabase nao configurado."
    params={"contrato_id":f"eq.{contrato_id()}","chave":f"eq.{chave}"}
    try:
        resp=requests.delete(_url(),headers=_headers("return=minimal"),params=params,timeout=timeout)
        resp.raise_for_status(); return True,"Registro removido."
    except Exception as exc:
        return False,f"Falha ao remover: {exc}"
