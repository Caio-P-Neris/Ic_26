#!/usr/bin/env python3
"""
ibov_wayback.py  (v2 — download direto do CSV)
═══════════════════════════════════════════════
Coleta a composição quadrimestral do IBOV (2018-2025) via Wayback Machine,
priorizando o download direto do arquivo CSV que existe na página B3.

Estratégias por ordem de prioridade:
  1. CSV download — sistemaswebb3 (URLs descobertas via CDX, ~2020+)
  2. CSV download — CDX ao vivo por data (API sistemaswebb3)
  3. CSV legado   — URLs B3/BM&FBovespa antigas (lumis, data/files)
  4. JSON paginado — API sistemaswebb3 (todas as páginas, fallback)
  5. HTML parse   — página principal B3 (último recurso)

Roda primeiro uma fase de DESCOBERTA via CDX para mapear quais URLs
de CSV existem no arquivo do WM, depois usa essa lista na coleta.

Uso:
  pip install requests pandas openpyxl
  python ibov_wayback.py
  python ibov_wayback.py --skip-existing
  python ibov_wayback.py --only 2021-Q2
  python ibov_wayback.py --discover-only     # só mostra URLs encontradas
  python ibov_wayback.py --rediscover        # refaz a varredura CDX
"""

import argparse
import base64
import io
import json
import os
import re
import sys
import time
from datetime import datetime
from typing import Optional

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ══════════════════════════════════════════════════════════════════════════════
# Configurações
# ══════════════════════════════════════════════════════════════════════════════

CDX_API  = "https://web.archive.org/cdx/search/cdx"
WB_BASE  = "https://web.archive.org/web"

# API nova B3 (~2020+) — retorna JSON paginado E CSV via download
B3_API_NEW_PREFIX = (
    "https://sistemaswebb3-listados.b3.com.br"
    "/indexProxy/indexCall/GetPortfolioDay/"
)

# Página principal (fallback HTML)
B3_PAGE = (
    "https://www.b3.com.br/pt_br/market-data-e-indices/indices/"
    "indices-amplos/indice-ibovespa-ibovespa-composicao-da-carteira.htm"
)

# Padrões legados de URL (2018-2019, BM&FBovespa / B3 antiga)
B3_LEGACY_PREFIXES = [
    "https://www.b3.com.br/lumis/portal/file/fileDownload.jsp",
    "https://www.b3.com.br/data/files/",
    "https://www.bmfbovespa.com.br/indices/download/",
    "https://www.bmfbovespa.com.br/lumis/portal/file/fileDownload.jsp",
    "https://www.b3.com.br/pt_br/market-data-e-indices/indices/indices-amplos/",
]

OUTPUT_DIR      = "ibov_composicao"
OUTPUT_CSV      = os.path.join(OUTPUT_DIR, "ibov_quadrimestral_2018_2025.csv")
OUTPUT_XLSX     = os.path.join(OUTPUT_DIR, "ibov_quadrimestral_2018_2025.xlsx")
DISCOVERY_CACHE = os.path.join(OUTPUT_DIR, "_discovered_urls.json")

DELAY   = 2    # segundos entre requests (respeitar WM)
MAX_RET = 3    # tentativas de retry HTTP
CDX_WIN = 60   # dias de tolerância padrão para snapshot

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/csv,application/json,text/html,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9",
}


# ══════════════════════════════════════════════════════════════════════════════
# Session HTTP com retry automático
# ══════════════════════════════════════════════════════════════════════════════

def build_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(total=MAX_RET, backoff_factor=2.0,
                  status_forcelist=[429, 500, 502, 503, 504])
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.mount("http://",  HTTPAdapter(max_retries=retry))
    s.headers.update(HEADERS)
    return s


SESSION = build_session()


# ══════════════════════════════════════════════════════════════════════════════
# Datas dos quadrimestres (jan / mai / set de 2018 a 2025)
# ══════════════════════════════════════════════════════════════════════════════

QUADRIMESTRES: list[dict] = [
    {
        "label":     f"{y}-Q{q}",
        "ano":       y,
        "quad":      q,
        "mes":       m,
        "data_alvo": datetime(y, m, 10),   # dia 10 = carteira já publicada
    }
    for y in range(2018, 2026)
    for q, m in [(1, 1), (2, 5), (3, 9)]
    if datetime(y, m, 10) <= datetime(2025, 12, 31)
]


# ══════════════════════════════════════════════════════════════════════════════
# CDX helpers
# ══════════════════════════════════════════════════════════════════════════════

def cdx_query(params: dict, timeout: int = 45) -> list[list]:
    """Executa uma consulta CDX e retorna as linhas (sem cabeçalho)."""
    try:
        r = SESSION.get(CDX_API, params=params, timeout=timeout)
        r.raise_for_status()
        rows = r.json()
        return rows[1:] if len(rows) > 1 else []
    except Exception as exc:
        print(f"    [CDX] {exc}")
        return []


def cdx_find_snapshot(
    url: str,
    target: datetime,
    window_days: int = CDX_WIN,
    match_type: str = "exact",
    mime_filter: Optional[str] = None,
) -> Optional[str]:
    """
    Retorna o timestamp WM (YYYYMMDDHHMMSS) do snapshot 200 mais próximo de
    `target` dentro de `window_days` dias. None se não encontrar.
    """
    filters = ["statuscode:200"]
    if mime_filter:
        filters.append(f"mimetype:{mime_filter}")

    params = {
        "url":       url,
        "output":    "json",
        "closest":   target.strftime("%Y%m%d"),
        "limit":     10,
        "fl":        "timestamp,statuscode,mimetype",
        "filter":    filters,
        "matchType": match_type,
    }
    for row in cdx_query(params):
        ts = row[0]
        snap_dt = datetime.strptime(ts[:8], "%Y%m%d")
        if abs((snap_dt - target).days) <= window_days:
            return ts
    return None


def cdx_list_all(url_prefix: str, mime_filter: Optional[str] = None,
                 limit: int = 5000) -> list[dict]:
    """
    Lista todos os snapshots arquivados (status 200) para um prefixo de URL.
    collapse=urlkey → uma entrada por URL única (evita duplicatas de snapshots).
    """
    params = {
        "url":       url_prefix,
        "output":    "json",
        "fl":        "timestamp,original,mimetype,statuscode",
        "filter":    ["statuscode:200"],
        "matchType": "prefix",
        "limit":     limit,
        "collapse":  "urlkey",
    }
    if mime_filter:
        params["filter"].append(f"mimetype:{mime_filter}")

    rows = cdx_query(params, timeout=90)
    return [{"timestamp": r[0], "url": r[1], "mimetype": r[2]} for r in rows]


# ══════════════════════════════════════════════════════════════════════════════
# Fase de descoberta — mapeia URLs de CSV/JSON do IBOV no Wayback Machine
# ══════════════════════════════════════════════════════════════════════════════

def discover_ibov_urls(force: bool = False) -> dict:
    """
    Varre o CDX para encontrar URLs B3 relacionadas ao IBOV arquivadas.
    Resultado salvo em cache JSON para evitar reprocessamento.

    Retorna:
      {
        "csv_new":    [...],   # CSV via API sistemaswebb3
        "json_new":   [...],   # JSON via API sistemaswebb3
        "csv_legacy": [...],   # CSV/arquivos legados B3
      }
    """
    if not force and os.path.isfile(DISCOVERY_CACHE):
        with open(DISCOVERY_CACHE, encoding="utf-8") as f:
            cached = json.load(f)
        total = sum(len(v) for v in cached.values())
        print(f"  📂  Cache de descoberta carregado ({total} entradas): {DISCOVERY_CACHE}")
        return cached

    print("\n" + "═"*60)
    print("  FASE DE DESCOBERTA — varredura CDX")
    print("═"*60)

    result: dict = {"csv_new": [], "json_new": [], "csv_legacy": []}

    # ── API nova: CSV ─────────────────────────────────────────────────────
    print(f"\n  [1/3] CSV em {B3_API_NEW_PREFIX[:55]} …")
    hits = cdx_list_all(B3_API_NEW_PREFIX, mime_filter="text/csv")
    result["csv_new"] = hits
    print(f"        → {len(hits)} snapshots CSV")
    time.sleep(DELAY)

    # ── API nova: JSON ────────────────────────────────────────────────────
    print(f"  [2/3] JSON em {B3_API_NEW_PREFIX[:55]} …")
    hits_j = cdx_list_all(B3_API_NEW_PREFIX, mime_filter="application/json")
    result["json_new"] = hits_j
    print(f"        → {len(hits_j)} snapshots JSON")
    time.sleep(DELAY)

    # ── URLs legadas ──────────────────────────────────────────────────────
    print("  [3/3] URLs legadas B3/BM&FBovespa …")
    for pfx in B3_LEGACY_PREFIXES:
        hits_leg = cdx_list_all(pfx)
        # Filtra por palavras-chave relacionadas ao IBOV
        keywords = ["ibov", "composicao", "carteira", "portfolio", "indice"]
        filtered = [
            h for h in hits_leg
            if any(kw in h["url"].lower() for kw in keywords)
        ]
        result["csv_legacy"].extend(filtered)
        print(f"        {pfx[:55]:55s}  → {len(filtered)}/{len(hits_leg)}")
        time.sleep(DELAY)

    # Remove duplicatas em csv_legacy
    seen = set()
    unique_legacy = []
    for e in result["csv_legacy"]:
        key = e["url"]
        if key not in seen:
            seen.add(key)
            unique_legacy.append(e)
    result["csv_legacy"] = unique_legacy

    # Salva cache
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(DISCOVERY_CACHE, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    total = sum(len(v) for v in result.values())
    print(f"\n  💾  Cache salvo: {DISCOVERY_CACHE}  ({total} entradas totais)")
    return result


# ══════════════════════════════════════════════════════════════════════════════
# Download via Wayback Machine
# ══════════════════════════════════════════════════════════════════════════════

def wayback_get(url: str, timestamp: str) -> Optional[bytes]:
    """
    Baixa o recurso arquivado e retorna os bytes brutos.
    Usa sufixo `if_` para obter o recurso puro (sem toolbar do WM).
    """
    wb_url = f"{WB_BASE}/{timestamp}if_/{url}"
    try:
        r = SESSION.get(wb_url, timeout=90)
        if r.status_code == 200 and len(r.content) > 100:
            return r.content
        print(f"      [WB] status={r.status_code}  size={len(r.content)}")
    except Exception as exc:
        print(f"      [WB GET] {exc}")
    return None


# ══════════════════════════════════════════════════════════════════════════════
# Parsers
# ══════════════════════════════════════════════════════════════════════════════

def parse_csv_bytes(data: bytes) -> Optional[pd.DataFrame]:
    """
    Tenta ler bytes como CSV.
    B3 usa ponto-e-vírgula como separador e vírgula como decimal.
    Testa múltiplas combinações de encoding e separador.
    """
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        for sep in (";", ",", "\t"):
            try:
                text = data.decode(enc)
                df = pd.read_csv(
                    io.StringIO(text),
                    sep=sep,
                    decimal=",",
                    thousands=".",
                    engine="python",
                    on_bad_lines="skip",
                )
                if df.shape[1] >= 2 and df.shape[0] >= 5:
                    df = normalize_df(df)
                    if "codigo" in df.columns and not df.empty:
                        return df
            except Exception:
                pass
    return None


def parse_json_bytes(data: bytes) -> Optional[pd.DataFrame]:
    """Interpreta bytes como JSON da API sistemaswebb3."""
    try:
        raw = json.loads(data.decode("utf-8"))
    except Exception:
        return None

    if isinstance(raw, list):
        results = raw
    elif isinstance(raw, dict):
        for key in ("results", "Results", "items", "data", "portfolioDay"):
            if key in raw and isinstance(raw[key], list):
                results = raw[key]
                break
        else:
            candidates = [v for v in raw.values() if isinstance(v, list) and v]
            results = candidates[0] if candidates else []
    else:
        return None

    if not results:
        return None

    df = pd.DataFrame(results)
    return normalize_df(df)


def parse_html_bytes(data: bytes) -> Optional[pd.DataFrame]:
    """
    Parse do HTML da página B3.
    Tenta extrair JSON embutido nos scripts da página e, depois,
    tabelas HTML convencionais.
    """
    html = None
    for enc in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            html = data.decode(enc)
            break
        except Exception:
            pass
    if not html:
        return None

    # 1. JSON embutido no JS da página
    json_patterns = [
        r'"results"\s*:\s*(\[\s*\{.*?\}\s*\])',
        r'"portfolioDay"\s*:\s*(\[\s*\{.*?\}\s*\])',
        r'"composition"\s*:\s*(\[\s*\{.*?\}\s*\])',
        r'(\[\s*\{"cod"\s*:.*?\}\s*\])',
    ]
    for pat in json_patterns:
        m = re.search(pat, html, re.DOTALL | re.IGNORECASE)
        if m:
            try:
                df = pd.DataFrame(json.loads(m.group(1)))
                df = normalize_df(df)
                if "codigo" in df.columns and len(df) >= 5:
                    return df
            except Exception:
                pass

    # 2. Tabelas HTML
    try:
        tables = pd.read_html(html, decimal=",", thousands=".")
        for t in tables:
            joined = " ".join(str(c).lower() for c in t.columns)
            if any(kw in joined for kw in ["cod", "papel", "ativo", "participação"]):
                t = normalize_df(t)
                if "codigo" in t.columns and len(t) >= 5:
                    return t
    except Exception:
        pass

    return None


# ══════════════════════════════════════════════════════════════════════════════
# Normalização de colunas
# ══════════════════════════════════════════════════════════════════════════════

_COL_MAP = {
    # código do ativo
    "cod": "codigo", "codRaw": "codigo", "asset": "codigo",
    "Cod": "codigo", "COD": "codigo", "Asset": "codigo",
    "Código": "codigo", "Papel": "codigo", "ticker": "codigo", "Ativo": "codigo",
    # nome da empresa
    "Ação": "nome", "Empresa": "nome", "name": "nome", "Name": "nome",
    # participação %
    "part": "part_pct", "Part": "part_pct", "PART": "part_pct",
    "participation": "part_pct", "percentual": "part_pct",
    "Participação": "part_pct", "Participacao": "part_pct",
    "Part. (%)": "part_pct", "Part.(%)": "part_pct",
    "Peso": "part_pct", "weight": "part_pct",
    # quantidade teórica
    "theoricQty": "qtd_teorica", "TheoricQty": "qtd_teorica",
    "Qtde. Teórica": "qtd_teorica", "Quantidade Teórica": "qtd_teorica",
    "qtdTeorica": "qtd_teorica",
    # tipo / segmento
    "type": "tipo", "Type": "tipo", "Tipo": "tipo",
    "setor": "segmento", "segment": "segmento",
    "Segment": "segmento", "Segmento": "segmento",
}

_INVALID_CODES = {"COD", "CÓDIGO", "PAPEL", "ATIVO", "TICKER", "NAN", "CÓDIGO"}


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.rename(columns={k: v for k, v in _COL_MAP.items() if k in df.columns})

    # Fallback: detecta coluna de código por substring
    if "codigo" not in df.columns:
        for col in df.columns:
            if any(kw in str(col).lower() for kw in ["cod", "papel", "ativo", "ticker"]):
                df = df.rename(columns={col: "codigo"})
                break

    if "codigo" in df.columns:
        df["codigo"] = df["codigo"].astype(str).str.strip().str.upper()
        df = df[df["codigo"].notna() & (df["codigo"] != "") & (df["codigo"] != "NAN")]
        df = df[~df["codigo"].isin(_INVALID_CODES)]

    return df.reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
# Builder de URLs da API B3
# ══════════════════════════════════════════════════════════════════════════════

def b3_api_url(page: int = 1, page_size: int = 150) -> str:
    """Gera URL da API sistemaswebb3 com parâmetros codificados em base64."""
    params = json.dumps(
        {"language": "pt-br", "pageNumber": page, "pageSize": page_size,
         "index": "IBOV", "segment": "1"},
        separators=(",", ":"),
    )
    return B3_API_NEW_PREFIX + base64.b64encode(params.encode()).decode()


# ══════════════════════════════════════════════════════════════════════════════
# Seleciona snapshot mais próximo dentre descobertos
# ══════════════════════════════════════════════════════════════════════════════

def closest_entry(entries: list[dict], target: datetime,
                  window_days: int = CDX_WIN) -> Optional[dict]:
    """
    Dentre uma lista de entradas descobertas {timestamp, url, …},
    retorna a mais próxima de `target` dentro de `window_days` dias.
    """
    best, best_delta = None, window_days + 1
    for e in entries:
        try:
            dt    = datetime.strptime(e["timestamp"][:8], "%Y%m%d")
            delta = abs((dt - target).days)
            if delta < best_delta:
                best_delta, best = delta, e
        except Exception:
            pass
    return best


# ══════════════════════════════════════════════════════════════════════════════
# Coleta de um quadrimestre
# ══════════════════════════════════════════════════════════════════════════════

def fetch_quarter(
    quad: dict, discovered: dict
) -> tuple[Optional[pd.DataFrame], Optional[str], str]:
    """
    Tenta obter a composição IBOV para o quadrimestre usando 5 estratégias.
    Retorna (DataFrame, timestamp_WM, nome_da_estratégia).
    """
    label  = quad["label"]
    target = quad["data_alvo"]

    print(f"\n{'─'*60}")
    print(f"  📅  {label}  |  alvo: {target.strftime('%d/%m/%Y')}")
    print(f"{'─'*60}")

    # ── [1] CSV descoberto — API sistemaswebb3 ────────────────────────────
    print("  [1] CSV download (URLs descobertas — API nova) …")
    entry = closest_entry(discovered.get("csv_new", []), target)
    if entry:
        snap_dt = datetime.strptime(entry["timestamp"][:8], "%Y%m%d")
        delta   = abs((snap_dt - target).days)
        print(f"      snapshot: {entry['timestamp']}  ({snap_dt.strftime('%d/%m/%Y')}, Δ={delta}d)")
        data = wayback_get(entry["url"], entry["timestamp"])
        time.sleep(DELAY)
        if data:
            df = parse_csv_bytes(data)
            if df is not None and not df.empty:
                print(f"      ✅  {len(df)} ativos  (CSV direto — API nova)")
                return df, entry["timestamp"], "csv_api_nova"
    else:
        print("      Sem entradas CSV descobertas para esta época.")

    # ── [2] CSV ao vivo via CDX por data — API sistemaswebb3 ──────────────
    print("  [2] CSV download (CDX ao vivo — API nova) …")
    ts = cdx_find_snapshot(B3_API_NEW_PREFIX, target,
                           match_type="prefix", mime_filter="text/csv")
    time.sleep(1)
    if ts:
        snap_dt = datetime.strptime(ts[:8], "%Y%m%d")
        print(f"      snapshot: {ts}  ({snap_dt.strftime('%d/%m/%Y')})")
        # Recupera a URL real desse timestamp
        rows = cdx_query({
            "url": B3_API_NEW_PREFIX, "output": "json",
            "from": ts, "to": ts,
            "fl": "timestamp,original",
            "matchType": "prefix",
            "filter": ["statuscode:200", "mimetype:text/csv"],
            "limit": 1,
        })
        real_url = rows[0][1] if rows else b3_api_url()
        data = wayback_get(real_url, ts)
        time.sleep(DELAY)
        if data:
            df = parse_csv_bytes(data)
            if df is not None and not df.empty:
                print(f"      ✅  {len(df)} ativos  (CSV CDX ao vivo)")
                return df, ts, "csv_cdx_live"

    # ── [3] CSV legado (URLs descobertas) ─────────────────────────────────
    print("  [3] CSV legado B3 (lumis / data/files / bmfbovespa) …")
    entry_leg = closest_entry(discovered.get("csv_legacy", []), target)
    if entry_leg:
        snap_dt = datetime.strptime(entry_leg["timestamp"][:8], "%Y%m%d")
        delta   = abs((snap_dt - target).days)
        print(f"      snapshot: {entry_leg['timestamp']}  ({snap_dt.strftime('%d/%m/%Y')}, Δ={delta}d)")
        print(f"      URL: {entry_leg['url'][:75]}")
        data = wayback_get(entry_leg["url"], entry_leg["timestamp"])
        time.sleep(DELAY)
        if data:
            df = parse_csv_bytes(data)
            if df is not None and not df.empty:
                print(f"      ✅  {len(df)} ativos  (CSV legado)")
                return df, entry_leg["timestamp"], "csv_legado"
            # Tenta também como JSON
            df = parse_json_bytes(data)
            if df is not None and not df.empty:
                print(f"      ✅  {len(df)} ativos  (JSON legado)")
                return df, entry_leg["timestamp"], "json_legado"
    else:
        print("      Sem entradas legadas descobertas.")

    # ── [4] JSON paginado — API sistemaswebb3 ─────────────────────────────
    print("  [4] JSON paginado — API sistemaswebb3 …")
    all_dfs, last_ts = [], None
    for page in range(1, 5):   # até 4 páginas (IBOV ~84 ativos → 1-2 pág.)
        api_url = b3_api_url(page=page)
        ts_p    = cdx_find_snapshot(api_url, target)
        time.sleep(1)
        if not ts_p:
            break
        if page == 1:
            snap_dt = datetime.strptime(ts_p[:8], "%Y%m%d")
            print(f"      snapshot: {ts_p}  ({snap_dt.strftime('%d/%m/%Y')})")
        last_ts = ts_p
        data = wayback_get(api_url, ts_p)
        time.sleep(DELAY)
        if not data:
            break
        df_p = parse_json_bytes(data)
        if df_p is None or df_p.empty:
            break
        all_dfs.append(df_p)
        if len(df_p) < 100:    # última página (menos que pageSize)
            break

    if all_dfs:
        df = pd.concat(all_dfs, ignore_index=True).drop_duplicates(subset=["codigo"])
        print(f"      ✅  {len(df)} ativos  (JSON, {len(all_dfs)} pág.)")
        return df, last_ts, "json_api"

    # ── [5] HTML — página principal B3 ───────────────────────────────────
    print("  [5] HTML — página principal B3 (janela ±90 dias) …")
    ts = cdx_find_snapshot(B3_PAGE, target, window_days=90)
    time.sleep(1)
    if ts:
        snap_dt = datetime.strptime(ts[:8], "%Y%m%d")
        print(f"      snapshot: {ts}  ({snap_dt.strftime('%d/%m/%Y')})")
        data = wayback_get(B3_PAGE, ts)
        time.sleep(DELAY)
        if data:
            df = parse_html_bytes(data)
            if df is not None and not df.empty:
                print(f"      ✅  {len(df)} ativos  (HTML parse)")
                return df, ts, "html_parse"

    print(f"  ❌  Sem dados para {label}")
    return None, None, "falhou"


# ══════════════════════════════════════════════════════════════════════════════
# Salvar resultados consolidados
# ══════════════════════════════════════════════════════════════════════════════

PRIORITY_COLS = [
    "periodo", "ano", "quadrimestre", "mes_vigencia",
    "codigo", "nome", "tipo", "part_pct", "qtd_teorica",
    "segmento", "estrategia", "snapshot_ts", "snapshot_data",
]


def save_results(frames: list[pd.DataFrame]) -> None:
    if not frames:
        print("\n⚠️   Nenhum dado coletado — arquivos não gerados.")
        return

    combined = pd.concat(frames, ignore_index=True)

    present  = [c for c in PRIORITY_COLS if c in combined.columns]
    rest     = [c for c in combined.columns if c not in present]
    combined = combined[present + rest]

    combined.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"\n  💾  CSV  → {OUTPUT_CSV}")

    try:
        combined.to_excel(OUTPUT_XLSX, index=False, engine="openpyxl")
        print(f"  💾  XLSX → {OUTPUT_XLSX}")
    except ImportError:
        print("  ⚠️   pip install openpyxl  para gerar XLSX")

    print(f"\n  Registros  : {len(combined):,}")
    print(f"  Períodos   : {combined['periodo'].nunique()}")
    print(f"  Estratégias: {combined['estrategia'].value_counts().to_dict()}"
          if "estrategia" in combined.columns else "")


# ══════════════════════════════════════════════════════════════════════════════
# Resumo de descoberta
# ══════════════════════════════════════════════════════════════════════════════

def print_discovery_summary(discovered: dict) -> None:
    print("\n" + "═"*60)
    print("  URLS DESCOBERTAS NO WAYBACK MACHINE")
    print("═"*60)
    for cat, entries in discovered.items():
        print(f"\n  [{cat}]  {len(entries)} entradas")
        for e in entries[:8]:
            snap_dt = datetime.strptime(e["timestamp"][:8], "%Y%m%d").strftime("%d/%m/%Y")
            print(f"    {snap_dt}  {e['url'][:75]}")
        if len(entries) > 8:
            print(f"    … e mais {len(entries)-8} entradas")


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--skip-existing", action="store_true",
                   help="Pula quadrimestres cujo CSV individual já existe")
    p.add_argument("--only", metavar="LABEL",
                   help="Coleta apenas um período (ex: 2021-Q2)")
    p.add_argument("--discover-only", action="store_true",
                   help="Executa só a fase de descoberta e exibe resultado")
    p.add_argument("--rediscover", action="store_true",
                   help="Força nova varredura CDX (ignora cache)")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("═"*60)
    print("  IBOV — Composição Quadrimestral 2018-2025")
    print("  Wayback Machine  |  v2 — download direto CSV")
    print("═"*60)

    # ── Fase 1: Descoberta CDX ────────────────────────────────────────────
    discovered = discover_ibov_urls(force=args.rediscover)
    print_discovery_summary(discovered)

    if args.discover_only:
        sys.exit(0)

    # ── Fase 2: Coleta quadrimestral ──────────────────────────────────────
    quads = QUADRIMESTRES
    if args.only:
        quads = [q for q in quads if q["label"] == args.only]
        if not quads:
            sys.exit(f"Período '{args.only}' não encontrado. Formato: YYYY-Q[1|2|3]")

    print(f"\n  Períodos a processar: {len(quads)}")

    all_frames: list[pd.DataFrame] = []
    summary:    list[dict]         = []

    for quad in quads:
        ind_csv = os.path.join(OUTPUT_DIR, f"ibov_{quad['label']}.csv")

        # Modo incremental
        if args.skip_existing and os.path.isfile(ind_csv):
            print(f"\n  ⏭️   {quad['label']}  — já coletado, pulando.")
            df_ex = pd.read_csv(ind_csv)
            all_frames.append(df_ex)
            summary.append({
                "periodo":    quad["label"],
                "n_ativos":   len(df_ex),
                "status":     "existente",
                "estrategia": "—",
                "snapshot":   df_ex.get("snapshot_ts", pd.Series(["?"])).iloc[0],
            })
            continue

        df, ts, estrategia = fetch_quarter(quad, discovered)
        ok = df is not None and not df.empty

        summary.append({
            "periodo":    quad["label"],
            "n_ativos":   len(df) if ok else 0,
            "status":     "ok" if ok else "falhou",
            "estrategia": estrategia,
            "snapshot":   ts or "—",
        })

        if ok:
            df = df.copy()
            df["periodo"]       = quad["label"]
            df["ano"]           = quad["ano"]
            df["quadrimestre"]  = quad["quad"]
            df["mes_vigencia"]  = quad["mes"]
            df["estrategia"]    = estrategia
            df["snapshot_ts"]   = ts
            df["snapshot_data"] = (
                datetime.strptime(ts[:8], "%Y%m%d").strftime("%d/%m/%Y") if ts else ""
            )
            df.to_csv(ind_csv, index=False, encoding="utf-8-sig")
            print(f"      📁  Salvo: {ind_csv}")
            all_frames.append(df)

        time.sleep(DELAY)

    # ── Resumo final ──────────────────────────────────────────────────────
    print("\n" + "═"*60)
    print("  RESUMO DA COLETA")
    print("═"*60)
    df_summ = pd.DataFrame(summary)
    print(df_summ.to_string(index=False))

    ok_n   = df_summ[df_summ["status"].isin(["ok", "existente"])].shape[0]
    fail_n = len(summary) - ok_n
    print(f"\n  ✅  Coletados : {ok_n}/{len(summary)}")
    print(f"  ❌  Falharam  : {fail_n}/{len(summary)}")

    save_results(all_frames)


if __name__ == "__main__":
    main()