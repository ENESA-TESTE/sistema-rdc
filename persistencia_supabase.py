"""Persistencia remota do SGO RDC & PDE via Supabase REST.

As credenciais devem ficar em .streamlit/secrets.toml:
SUPABASE_URL = "https://SEU-PROJETO.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "SUA_CHAVE_SERVICE_ROLE"
SGO_CONTRATO_ID = "inocencia"
"""
from __future__ import annotations

import json
import os
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
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


def _url() -> str:
    return _secret("SUPABASE_URL").rstrip("/") + "/rest/v1/sgo_estado"


def salvar_estado(chave: str, valor: Any, timeout: int = 20) -> tuple[bool, str]:
    """Faz upsert de um estado JSON por contrato e chave."""
    if not configurado():
        return False, "Supabase não configurado."
    payload = {
        "contrato_id": contrato_id(),
        "chave": chave,
        "valor": valor,
        "atualizado_em": datetime.now(timezone.utc).isoformat(),
    }
    try:
        resp = requests.post(
            _url() + "?on_conflict=contrato_id,chave",
            headers=_headers("resolution=merge-duplicates,return=minimal"),
            data=json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8"),
            timeout=timeout,
        )
        resp.raise_for_status()
        return True, "Salvo no banco persistente."
    except Exception as exc:
        return False, f"Falha no banco persistente: {exc}"


def carregar_estado(chave: str, padrao: Any = None, timeout: int = 20) -> Any:
    """Carrega estado JSON. Retorna padrao quando ausente ou indisponível."""
    if not configurado():
        return padrao
    params = {
        "contrato_id": f"eq.{contrato_id()}",
        "chave": f"eq.{chave}",
        "select": "valor,atualizado_em",
        "limit": "1",
    }
    try:
        resp = requests.get(_url(), headers=_headers(), params=params, timeout=timeout)
        resp.raise_for_status()
        dados = resp.json()
        if not dados:
            return padrao
        return dados[0].get("valor", padrao)
    except Exception:
        return padrao


def apagar_estado(chave: str, timeout: int = 20) -> tuple[bool, str]:
    if not configurado():
        return False, "Supabase não configurado."
    params = {"contrato_id": f"eq.{contrato_id()}", "chave": f"eq.{chave}"}
    try:
        resp = requests.delete(_url(), headers=_headers("return=minimal"), params=params, timeout=timeout)
        resp.raise_for_status()
        return True, "Registro removido."
    except Exception as exc:
        return False, f"Falha ao remover: {exc}"
