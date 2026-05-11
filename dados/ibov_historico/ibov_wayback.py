#!/usr/bin/env python3
"""
ibov_wayback.py  (v3 — HTML/iframe + JSON API)
═══════════════════════════════════════════════
Coleta a composição quadrimestral do IBOV (2018-2025) via Wayback Machine.

A página da B3 embute o conteúdo em um iframe apontando para
sistemaswebb3-listados.b3.com.br. O Wayback Machine arquiva o iframe
como URL separada, onde está a tabela com class="table table-responsive-md".

Estratégias por ordem:
  1. JSON API   — sistemaswebb3/GetPortfolioDay  (2021+, confirmado)
  2. HTML iframe — página principal → extrai src do iframe → WM → BeautifulSoup
  3. HTML direto — busca URL do iframe diretamente no CDX (sem passar pela página pai)
  4. HTML página — tenta a página principal como fallback (captura server-side render)

Dependências:
  pip install requests pandas beautifulsoup4 lxml openpyxl

Uso:
  python ibov_wayback.py
  python ibov_wayback.py --skip-existing
  python ibov_wayback.py --only 2019-Q1
  python ibov_wayback.py --rediscover
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
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ══════════════════════════════════════════════════════════════════════════════
# Configurações
# ══════════════════════════════════════════════════════════════════════════════

CDX_API = "https://web.archive.org/cdx/search/cdx"
WB_BASE = "https://web.archive.org/web"

# Página principal da carteira IBOV
B3_PAGE = (
    "https://www.b3.com.br/pt_br/market-data-e-indices/indices/"
    "indices-amplos/indice-ibovespa-ibovespa-composicao-da-carteira.htm"
)

# API JSON paginada (vigente ~2020+, 59 snapshots confirmados no WM)
B3_API_PREFIX = (
    "https://sistemaswebb3-listados.b3.com.br"
    "/indexProxy/indexCall/GetPortfolioDay/"
)

# Prefixo do iframe — o WM arquivou essa URL separadamente
IFRAME_PREFIX = "https://sistemaswebb3-listados.b3.com.br/indexPage/"

# Classe CSS da tabela de composição (confirmada na inspeção)
TABLE_CLASS = "table-responsive-md"   # substring que identifica a tabela certa

OUTPUT_DIR      = "ibov_composicao"
OUTPUT_CSV      = os.path.join(OUTPUT_DIR, "ibov_quadrimestral_2018_2025.csv")
OUTPUT_XLSX     = os.path.join(OUTPUT_DIR, "ibov_quadrimestral_2018_2025.xlsx")
IFRAME_CACHE    = os.path.join(OUTPUT_DIR, "_iframe_urls.json")

DELAY   = 2     # segundos entre requests ao WM
MAX_RET = 3
CDX_WIN = 60    # dias de tolerância padrão

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/json,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9",
}


# ══════════════════════════════════════════════════════════════════════════════
# Session HTTP
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
# Datas dos quadrimestres
# ══════════════════════════════════════════════════════════════════════════════

QUADRIMESTRES: list[dict] = [
    {
        "label":     f"{y}-Q{q}",
        "ano":       y,
        "quad":      q,
        "mes":       m,
        "data_alvo": datetime(y, m, 10),
    }
    for y in range(2018, 2026)
    for q, m in [(1, 1), (2, 5), (3, 9)]
    if datetime(y, m, 10) <= datetime(2025, 12, 31)
]


# ══════════════════════════════════════════════════════════════════════════════
# CDX helpers
# ══════════════════════════════════════════════════════════════════════════════

def cdx_query(params: dict, timeout: int = 45) -> list[list]:
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
    """Retorna o timestamp WM mais próximo de `target` dentro de `window_days`."""
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
        if abs((datetime.strptime(ts[:8], "%Y%m%d") - target).days) <= window_days:
            return ts
    return None


def cdx_list_prefix(url_prefix: str, limit: int = 3000) -> list[dict]:
    """Lista todos os snapshots 200 para um prefixo de URL (collapse por URL única)."""
    params = {
        "url":       url_prefix,
        "output":    "json",
        "fl":        "timestamp,original,mimetype",
        "filter":    ["statuscode:200"],
        "matchType": "prefix",
        "limit":     limit,
        "collapse":  "urlkey",
    }
    rows = cdx_query(params, timeout=90)
    return [{"timestamp": r[0], "url": r[1], "mimetype": r[2]} for r in rows]


# ══════════════════════════════════════════════════════════════════════════════
# Descoberta das URLs de iframe arquivadas
# ══════════════════════════════════════════════════════════════════════════════

def discover_iframe_urls(force: bool = False) -> list[dict]:
    """
    Mapeia todas as URLs do iframe sistemaswebb3/indexPage arquivadas no WM.
    Essas são as URLs que contêm a tabela HTML da composição IBOV.
    Resultado em cache para evitar reprocessamento.
    """
    if not force and os.path.isfile(IFRAME_CACHE):
        with open(IFRAME_CACHE, encoding="utf-8") as f:
            cached = json.load(f)
        print(f"  📂  Cache iframe carregado ({len(cached)} entradas): {IFRAME_CACHE}")
        return cached

    print("\n  🔍  Descobrindo URLs de iframe no Wayback Machine …")
    entries = cdx_list_prefix(IFRAME_PREFIX)

    # Filtra entradas relevantes: HTML e que contenham "ibov" ou "ibovespa" na URL
    filtered = [
        e for e in entries
        if any(kw in e["url"].lower() for kw in ["ibov", "ibovespa", "indice"])
        or "text/html" in e.get("mimetype", "")
    ]

    # Se não filtrou nada relevante, mantém tudo (pode ter URL genérica)
    result = filtered if filtered else entries

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(IFRAME_CACHE, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print(f"  💾  {len(result)} URLs de iframe salvas em {IFRAME_CACHE}")
    return result


def closest_iframe(entries: list[dict], target: datetime,
                   window_days: int = CDX_WIN) -> Optional[dict]:
    """Retorna a entrada descoberta mais próxima de `target`."""
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
# Download via Wayback Machine
# ══════════════════════════════════════════════════════════════════════════════

def wayback_get(url: str, timestamp: str) -> Optional[bytes]:
    """Baixa o recurso arquivado (sufixo if_ = sem toolbar do WM)."""
    wb_url = f"{WB_BASE}/{timestamp}if_/{url}"
    try:
        r = SESSION.get(wb_url, timeout=90)
        if r.status_code == 200 and len(r.content) > 200:
            return r.content
        print(f"      [WB] status={r.status_code}  bytes={len(r.content)}")
    except Exception as exc:
        print(f"      [WB] {exc}")
    return None


# ══════════════════════════════════════════════════════════════════════════════
# Extração do src do iframe na página principal
# ══════════════════════════════════════════════════════════════════════════════

def extract_iframe_src(html_bytes: bytes) -> Optional[str]:
    """
    Encontra a URL do iframe na página principal da B3.
    Procura por <iframe> e <object> que apontem para sistemaswebb3 ou b3.com.br.
    """
    for enc in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            html = html_bytes.decode(enc)
            break
        except Exception:
            html = None
    if not html:
        return None

    soup = BeautifulSoup(html, "lxml")

    # 1. <iframe src="...">
    for tag in soup.find_all("iframe"):
        src = tag.get("src", "")
        if src and ("sistemaswebb3" in src or "b3.com.br" in src):
            return src if src.startswith("http") else "https:" + src

    # 2. <object data="..."> ou <embed src="...">
    for tag in soup.find_all(["object", "embed"]):
        src = tag.get("data", tag.get("src", ""))
        if src and "b3.com.br" in src:
            return src if src.startswith("http") else "https:" + src

    # 3. URLs de iframe em scripts JS inline (às vezes injetado via JS)
    matches = re.findall(
        r"""(?:src|url|iframe)['":\s]+(['"](https?://sistemaswebb3[^'"]+)['")])""",
        html, re.IGNORECASE
    )
    if matches:
        return matches[0][1] if isinstance(matches[0], tuple) else matches[0]

    return None


# ══════════════════════════════════════════════════════════════════════════════
# Parser HTML — BeautifulSoup na tabela conhecida
# ══════════════════════════════════════════════════════════════════════════════

def parse_table_bs4(html_bytes: bytes) -> Optional[pd.DataFrame]:
    """
    Usa BeautifulSoup para encontrar a tabela de composição IBOV.
    Procura por class="table table-responsive-sm table-responsive-md"
    e faz o parse linha a linha.
    """
    for enc in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            html = html_bytes.decode(enc)
            break
        except Exception:
            html = None
    if not html:
        return None

    soup = BeautifulSoup(html, "lxml")

    # Tenta encontrar a tabela pela classe confirmada na inspeção
    table = (
        soup.find("table", class_=lambda c: c and TABLE_CLASS in c)
        or soup.find("table", class_=lambda c: c and "table-responsive" in (c or ""))
        or soup.find("table")   # fallback: primeira tabela encontrada
    )

    if table is None:
        # Pode haver iframes aninhados no próprio HTML arquivado
        nested_docs = soup.find_all("html")
        for doc in nested_docs:
            table = doc.find("table", class_=lambda c: c and "table" in (c or ""))
            if table:
                break

    if table is None:
        return None

    # Cabeçalho
    thead = table.find("thead")
    headers = []
    if thead:
        headers = [th.get_text(strip=True) for th in thead.find_all(["th", "td"])]

    # Linhas do corpo
    tbody = table.find("tbody")
    rows  = []
    if tbody:
        for tr in tbody.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if cells and any(c for c in cells):   # ignora linhas vazias
                rows.append(cells)

    if not rows:
        return None

    try:
        if headers and len(headers) == len(rows[0]):
            df = pd.DataFrame(rows, columns=headers)
        else:
            df = pd.DataFrame(rows)
            # Usa primeira linha como cabeçalho se parecer um header
            if df.shape[0] > 1 and all(isinstance(v, str) for v in df.iloc[0]):
                df.columns = df.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)
    except Exception:
        return None

    return normalize_df(df)


# ══════════════════════════════════════════════════════════════════════════════
# Parser JSON (API sistemaswebb3)
# ══════════════════════════════════════════════════════════════════════════════

def parse_json_bytes(data: bytes) -> Optional[pd.DataFrame]:
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

    return normalize_df(pd.DataFrame(results))


# ══════════════════════════════════════════════════════════════════════════════
# Normalização de colunas
# ══════════════════════════════════════════════════════════════════════════════

_COL_MAP = {
    "cod": "codigo", "codRaw": "codigo", "asset": "codigo",
    "Cod": "codigo", "COD": "codigo", "Código": "codigo",
    "Papel": "codigo", "ticker": "codigo", "Ativo": "codigo",
    "Ação": "nome", "Empresa": "nome", "name": "nome",
    "part": "part_pct", "Part": "part_pct",
    "Participação": "part_pct", "Part. (%)": "part_pct",
    "Peso": "part_pct", "participation": "part_pct",
    "theoricQty": "qtd_teorica", "TheoricQty": "qtd_teorica",
    "Qtde. Teórica": "qtd_teorica", "Quantidade Teórica": "qtd_teorica",
    "type": "tipo", "Type": "tipo", "Tipo": "tipo",
    "setor": "segmento", "segment": "segmento", "Segmento": "segmento",
}

_INVALID = {"COD", "CÓDIGO", "PAPEL", "ATIVO", "TICKER", "NAN", ""}


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.rename(columns={k: v for k, v in _COL_MAP.items() if k in df.columns})

    if "codigo" not in df.columns:
        for col in df.columns:
            if any(kw in str(col).lower() for kw in ["cod", "papel", "ativo", "ticker"]):
                df = df.rename(columns={col: "codigo"})
                break

    if "codigo" in df.columns:
        df["codigo"] = df["codigo"].astype(str).str.strip().str.upper()
        df = df[~df["codigo"].isin(_INVALID) & df["codigo"].notna()]

    return df.reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
# Builder URL JSON API
# ══════════════════════════════════════════════════════════════════════════════

def b3_api_url(page: int = 1, page_size: int = 150) -> str:
    params = json.dumps(
        {"language": "pt-br", "pageNumber": page, "pageSize": page_size,
         "index": "IBOV", "segment": "1"},
        separators=(",", ":"),
    )
    return B3_API_PREFIX + base64.b64encode(params.encode()).decode()


# ══════════════════════════════════════════════════════════════════════════════
# Coleta de um quadrimestre
# ══════════════════════════════════════════════════════════════════════════════

def fetch_quarter(
    quad: dict, iframe_entries: list[dict]
) -> tuple[Optional[pd.DataFrame], Optional[str], str]:
    """
    Tenta 4 estratégias para obter a composição do IBOV no quadrimestre.
    Retorna (DataFrame, timestamp_WM, estratégia_usada).
    """
    label  = quad["label"]
    target = quad["data_alvo"]

    print(f"\n{'─'*60}")
    print(f"  📅  {label}  |  alvo: {target.strftime('%d/%m/%Y')}")
    print(f"{'─'*60}")

    # ── [1] JSON API (funciona para 2021+) ───────────────────────────────
    print("  [1] JSON API — sistemaswebb3/GetPortfolioDay …")
    all_dfs, last_ts = [], None
    for page in range(1, 5):
        api_url = b3_api_url(page=page)
        ts = cdx_find_snapshot(api_url, target)
        time.sleep(1)
        if not ts:
            if page == 1:
                print("      Sem snapshot na janela padrão.")
            break
        if page == 1:
            snap_dt = datetime.strptime(ts[:8], "%Y%m%d")
            print(f"      snapshot: {ts}  ({snap_dt.strftime('%d/%m/%Y')})")
        last_ts = ts
        data = wayback_get(api_url, ts)
        time.sleep(DELAY)
        if not data:
            break
        df_p = parse_json_bytes(data)
        if df_p is None or df_p.empty:
            break
        all_dfs.append(df_p)
        if len(df_p) < 100:   # menos que pageSize → última página
            break

    if all_dfs:
        df = pd.concat(all_dfs, ignore_index=True).drop_duplicates(subset=["codigo"])
        print(f"      ✅  {len(df)} ativos  (JSON API, {len(all_dfs)} pág.)")
        return df, last_ts, "json_api"

    # ── [2] HTML iframe (descoberto via CDX) ─────────────────────────────
    print("  [2] HTML iframe — URLs descobertas …")
    entry = closest_iframe(iframe_entries, target)
    if entry:
        snap_dt = datetime.strptime(entry["timestamp"][:8], "%Y%m%d")
        delta   = abs((snap_dt - target).days)
        print(f"      snapshot: {entry['timestamp']}  ({snap_dt.strftime('%d/%m/%Y')}, Δ={delta}d)")
        print(f"      URL: {entry['url'][:75]}")
        data = wayback_get(entry["url"], entry["timestamp"])
        time.sleep(DELAY)
        if data:
            df = parse_table_bs4(data)
            if df is not None and not df.empty:
                print(f"      ✅  {len(df)} ativos  (HTML iframe CDX)")
                return df, entry["timestamp"], "html_iframe_cdx"
            else:
                print("      Tabela não encontrada no iframe.")
    else:
        print("      Sem entradas de iframe descobertas.")

    # ── [3] HTML iframe via página principal → seguir src ────────────────
    print("  [3] HTML iframe — extraindo src da página principal …")
    ts_main = cdx_find_snapshot(B3_PAGE, target, window_days=90)
    time.sleep(1)
    if ts_main:
        snap_dt = datetime.strptime(ts_main[:8], "%Y%m%d")
        print(f"      snapshot página: {ts_main}  ({snap_dt.strftime('%d/%m/%Y')})")
        main_data = wayback_get(B3_PAGE, ts_main)
        time.sleep(DELAY)
        if main_data:
            iframe_src = extract_iframe_src(main_data)
            if iframe_src:
                print(f"      iframe src: {iframe_src[:75]}")
                # Busca o iframe com timestamp próximo ao da página pai
                ts_iframe = cdx_find_snapshot(iframe_src, target, window_days=90)
                time.sleep(1)
                if ts_iframe:
                    snap_dt2 = datetime.strptime(ts_iframe[:8], "%Y%m%d")
                    print(f"      snapshot iframe: {ts_iframe}  ({snap_dt2.strftime('%d/%m/%Y')})")
                    iframe_data = wayback_get(iframe_src, ts_iframe)
                    time.sleep(DELAY)
                    if iframe_data:
                        df = parse_table_bs4(iframe_data)
                        if df is not None and not df.empty:
                            print(f"      ✅  {len(df)} ativos  (HTML iframe seguido)")
                            return df, ts_iframe, "html_iframe_seguido"
                        else:
                            print("      Tabela não encontrada no iframe seguido.")
                else:
                    print("      Sem snapshot do iframe no WM.")
            else:
                print("      Nenhum iframe encontrado na página principal.")

    # ── [4] HTML iframe direto via CDX (prefixo) ─────────────────────────
    print("  [4] HTML iframe — CDX ao vivo (prefixo sistemaswebb3/indexPage) …")
    ts_if = cdx_find_snapshot(
        IFRAME_PREFIX, target, match_type="prefix",
        mime_filter="text/html", window_days=90,
    )
    time.sleep(1)
    if ts_if:
        snap_dt = datetime.strptime(ts_if[:8], "%Y%m%d")
        print(f"      snapshot: {ts_if}  ({snap_dt.strftime('%d/%m/%Y')})")
        # Pega a URL real arquivada nesse timestamp
        rows = cdx_query({
            "url":       IFRAME_PREFIX,
            "output":    "json",
            "from":      ts_if,
            "to":        ts_if,
            "fl":        "timestamp,original",
            "matchType": "prefix",
            "filter":    ["statuscode:200", "mimetype:text/html"],
            "limit":     5,
        })
        urls_to_try = [r[1] for r in rows] if rows else [IFRAME_PREFIX]
        for iframe_url in urls_to_try:
            print(f"      tentando: {iframe_url[:75]}")
            data = wayback_get(iframe_url, ts_if)
            time.sleep(DELAY)
            if data:
                df = parse_table_bs4(data)
                if df is not None and not df.empty:
                    print(f"      ✅  {len(df)} ativos  (HTML iframe CDX ao vivo)")
                    return df, ts_if, "html_iframe_live"
    else:
        print("      Sem snapshot de iframe na janela ±90 dias.")

    print(f"  ❌  Sem dados para {label}")
    return None, None, "falhou"


# ══════════════════════════════════════════════════════════════════════════════
# Salvar resultados
# ══════════════════════════════════════════════════════════════════════════════

PRIORITY_COLS = [
    "periodo", "ano", "quadrimestre", "mes_vigencia",
    "codigo", "nome", "tipo", "part_pct", "qtd_teorica",
    "segmento", "estrategia", "snapshot_ts", "snapshot_data",
]


def save_results(frames: list[pd.DataFrame]) -> None:
    if not frames:
        print("\n⚠️   Nenhum dado coletado.")
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
    if "estrategia" in combined.columns:
        print(f"  Estratégias: {combined.groupby('estrategia')['periodo'].nunique().to_dict()}")


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--skip-existing", action="store_true",
                   help="Pula quadrimestres cujo CSV individual já existe")
    p.add_argument("--only", metavar="LABEL",
                   help="Coleta apenas um período (ex: 2019-Q1)")
    p.add_argument("--rediscover", action="store_true",
                   help="Refaz a varredura CDX dos iframes (ignora cache)")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("═"*60)
    print("  IBOV — Composição Quadrimestral 2018-2025")
    print("  Wayback Machine  |  v3 — HTML iframe + JSON API")
    print("═"*60)

    # Fase de descoberta: mapeia URLs de iframe no WM
    iframe_entries = discover_iframe_urls(force=args.rediscover)

    if iframe_entries:
        print(f"\n  Iframes no WM: {len(iframe_entries)} URLs únicas")
        # Mostra intervalo de datas coberto
        dts = sorted(e["timestamp"][:8] for e in iframe_entries)
        d0  = datetime.strptime(dts[0],  "%Y%m%d").strftime("%d/%m/%Y")
        d1  = datetime.strptime(dts[-1], "%Y%m%d").strftime("%d/%m/%Y")
        print(f"  Cobertura   : {d0} → {d1}")
        # Mostra exemplos de URLs únicas
        unique_urls = list({e["url"] for e in iframe_entries})[:5]
        print("  Exemplos de URLs:")
        for u in unique_urls:
            print(f"    {u[:78]}")
    else:
        print("\n  ⚠️   Nenhuma URL de iframe encontrada no WM.")

    # Coleta
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

        if args.skip_existing and os.path.isfile(ind_csv):
            print(f"\n  ⏭️   {quad['label']}  — já coletado, pulando.")
            df_ex = pd.read_csv(ind_csv)
            all_frames.append(df_ex)
            summary.append({
                "periodo":    quad["label"],
                "n_ativos":   len(df_ex),
                "status":     "existente",
                "estrategia": "—",
            })
            continue

        df, ts, estrategia = fetch_quarter(quad, iframe_entries)
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
            print(f"      📁  {ind_csv}")
            all_frames.append(df)

        time.sleep(DELAY)

    # Resumo
    print("\n" + "═"*60)
    print("  RESUMO DA COLETA")
    print("═"*60)
    df_summ = pd.DataFrame(summary)
    print(df_summ.to_string(index=False))

    ok_n = df_summ[df_summ["status"].isin(["ok", "existente"])].shape[0]
    print(f"\n  ✅  Coletados : {ok_n}/{len(summary)}")
    print(f"  ❌  Falharam  : {len(summary)-ok_n}/{len(summary)}")

    save_results(all_frames)


if __name__ == "__main__":
    main()