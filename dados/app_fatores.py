"""
app_fatores.py
==============
Interface web (Streamlit) para montar uma planilha de fatores fundamentalistas
a partir do ranking_historico.csv já baixado da CVM.

O usuário escolhe:
    • o período (faixa de trimestres),
    • os ativos (por ticker, mapeados via brapi.dev),
    • quais fatores deseja calcular.

Fatores que dependem de preço de mercado (P/L, EV/EBIT, Preço Graham, Margem de
Segurança, Market Cap, LPA, VPA) buscam cotações no yfinance apenas para os
ativos selecionados. Os demais saem direto do CSV.

Detecção robusta: cada fator só fica habilitado se as colunas necessárias
existirem no CSV carregado (funciona com o formato antigo e o novo).

Como rodar:
    pip install streamlit pandas numpy requests yfinance rapidfuzz
    streamlit run app_fatores.py
"""

from pathlib import Path

import numpy as np
import pandas as pd
import requests
import streamlit as st
import yfinance as yf
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Fatores Fundamentalistas — CVM", layout="wide")

CAMINHO_PADRAO = Path(__file__).parent / "ranking_historico.csv"


# ===========================================================================
# Definição dos fatores
#   precisa_preco : exige cotação do yfinance
#   cols          : colunas do CSV necessárias (além de preço)
# ===========================================================================
FATORES = {
    # --- Fundamentalistas (direto do CSV) ---
    "roic":            dict(label="ROIC",                 grupo="Fundamentalista", precisa_preco=False),
    "roe":             dict(label="ROE",                  grupo="Fundamentalista", precisa_preco=False, cols=["ll_anualizado", "pl"]),
    "margem_liquida":  dict(label="Margem Líquida",       grupo="Fundamentalista", precisa_preco=False, cols=["ll_anualizado", "receita_anualizada"]),
    "ebitda":          dict(label="EBITDA",               grupo="Fundamentalista", precisa_preco=False, cols=["ebitda"]),
    "dl_ebitda":       dict(label="Dív. Líq. / EBITDA",   grupo="Fundamentalista", precisa_preco=False, cols=["divida_liquida", "ebitda"]),
    "divida_liquida":  dict(label="Dívida Líquida",       grupo="Fundamentalista", precisa_preco=False, cols=["divida_liquida"]),
    "cagr_lucro":      dict(label="CAGR Lucro 5a",        grupo="Fundamentalista", precisa_preco=False, cols=["ll_anualizado"]),
    "cagr_receita":    dict(label="CAGR Receita 5a",      grupo="Fundamentalista", precisa_preco=False, cols=["receita_anualizada"]),
    # --- Dependentes de preço (yfinance) ---
    "preco":           dict(label="Preço (fechamento)",   grupo="De preço",        precisa_preco=True,  cols=[]),
    "market_cap":      dict(label="Market Cap",           grupo="De preço",        precisa_preco=True,  cols=["qt_acoes"]),
    "ev":              dict(label="Enterprise Value (EV)",grupo="De preço",        precisa_preco=True,  cols=["qt_acoes", "divida_liquida"]),
    "ev_ebit":         dict(label="EV/EBIT",              grupo="De preço",        precisa_preco=True,  cols=["qt_acoes", "divida_liquida", "ebit_anualizado"]),
    "p_l":             dict(label="P/L",                  grupo="De preço",        precisa_preco=True,  cols=["qt_acoes", "ll_anualizado"]),
    "lpa":             dict(label="LPA",                  grupo="De preço",        precisa_preco=True,  cols=["qt_acoes", "ll_anualizado"]),
    "vpa":             dict(label="VPA",                  grupo="De preço",        precisa_preco=True,  cols=["qt_acoes", "pl"]),
    "preco_graham":    dict(label="Preço Justo (Graham)", grupo="De preço",        precisa_preco=True,  cols=["qt_acoes", "ll_anualizado", "pl"]),
    "margem_seguranca":dict(label="Margem de Segurança",  grupo="De preço",        precisa_preco=True,  cols=["qt_acoes", "ll_anualizado", "pl"]),
}


def fator_disponivel(chave: str, colunas: set) -> bool:
    """Diz se um fator pode ser calculado com as colunas presentes no CSV."""
    if chave == "roic":
        return ("roic" in colunas) or ({"nopat", "capital_investido"} <= colunas)
    return set(FATORES[chave].get("cols", [])) <= colunas


# ===========================================================================
# Carregamento e mapeamentos (com cache)
# ===========================================================================

@st.cache_data(show_spinner=False)
def carregar_base(conteudo_ou_caminho) -> pd.DataFrame:
    if isinstance(conteudo_ou_caminho, (str, Path)):
        df = pd.read_csv(conteudo_ou_caminho)
    else:
        df = pd.read_csv(conteudo_ou_caminho)
    df.columns = [c.lstrip("﻿") for c in df.columns]
    df["DT_FIM_EXERC"] = pd.to_datetime(df["DT_FIM_EXERC"], errors="coerce")
    df = df.dropna(subset=["DT_FIM_EXERC"])
    return df


def _normalizar_nome(nome: str) -> str:
    remover = [
        "S.A.", "S/A", "SA", "LTDA", "LTDA.", "S.A", "/SA",
        "CIA.", "CIA", "COMPANHIA", "PARTICIPACOES", "PARTICIPAÇÕES",
        "HOLDING", "GROUP", "BRASIL", "DO BRASIL",
        "EM RECUPERACAO JUDICIAL", "EM LIQUIDACAO EXTRAJUDICIAL",
    ]
    nome = str(nome).upper()
    for t in remover:
        nome = nome.replace(t, "")
    return " ".join(nome.split())


@st.cache_data(show_spinner="Mapeando empresas → tickers (brapi.dev)...")
def mapear_tickers(empresas: pd.DataFrame) -> pd.DataFrame:
    """
    empresas: DataFrame [CNPJ_CIA, DENOM_CIA] (únicos).
    Retorna [CNPJ_CIA, DENOM_CIA, ticker] para os que casaram.
    """
    try:
        r = requests.get("https://brapi.dev/api/quote/list", timeout=30)
        r.raise_for_status()
        b = pd.DataFrame(r.json().get("stocks", []))[["stock", "name"]]
        b.columns = ["ticker", "nome"]
        b["ticker"] = b["ticker"].str.upper().str.strip()
        b = b[b["ticker"].str.match(r"^[A-Z]{4}\d{1,2}$")].drop_duplicates("ticker")
    except Exception as e:
        st.error(f"Falha ao consultar a brapi.dev: {e}")
        return pd.DataFrame(columns=["CNPJ_CIA", "DENOM_CIA", "ticker"])

    nomes_norm = b["nome"].apply(_normalizar_nome).tolist()
    tickers = b["ticker"].tolist()

    out = empresas.copy()
    def _match(nome):
        m = process.extractOne(_normalizar_nome(nome), nomes_norm, scorer=fuzz.token_sort_ratio)
        if m and m[1] >= 72:
            return tickers[nomes_norm.index(m[0])], m[1]
        return None, 0

    res = out["DENOM_CIA"].map(_match)
    out["ticker"] = [r[0] for r in res]
    out["_score"] = [r[1] for r in res]
    out = out.dropna(subset=["ticker"])
    # 1 ticker → 1 empresa (maior score): evita que holdings/subsidiárias de nome
    # parecido roubem o ticker da principal (ex.: ENGI11, LAVV3 com falsos positivos).
    out = out.sort_values("_score", ascending=False).drop_duplicates("ticker", keep="first")
    return out[["CNPJ_CIA", "DENOM_CIA", "ticker"]]


@st.cache_data(show_spinner=False)
def buscar_precos(ticker: str, datas: tuple) -> pd.DataFrame:
    """Preço de fechamento mensal mais próximo de cada data, via yfinance."""
    datas_dt = pd.to_datetime(list(datas))
    try:
        hist = yf.Ticker(f"{ticker}.SA").history(
            start=datas_dt.min() - pd.DateOffset(days=15),
            end=datas_dt.max() + pd.DateOffset(days=15),
            interval="1mo", auto_adjust=True,
        )
        if hist.empty:
            return pd.DataFrame(columns=["DT_FIM_EXERC", "preco"])
        hist = hist[["Close"]].reset_index()
        hist["Date"] = pd.to_datetime(hist["Date"]).dt.tz_localize(None)
        rows = [{"DT_FIM_EXERC": d, "preco": hist.loc[(hist["Date"] - d).abs().idxmin(), "Close"]}
                for d in datas_dt]
        return pd.DataFrame(rows)
    except Exception:
        return pd.DataFrame(columns=["DT_FIM_EXERC", "preco"])


# ===========================================================================
# Cálculo dos fatores
# ===========================================================================

def _normalizar_escala_acoes(qt, pl, preco, limite_pb=0.05):
    """
    Corrige a escala de qt_acoes. A composição de capital da CVM às vezes informa
    as ações em MILHARES (o arquivo não tem coluna de escala), deixando qt ~1000x
    menor e inflando LPA/VPA/Graham. Detecta pelo P/B implícito (market_cap / PL):
    enquanto < limite_pb, multiplica qt por 1000 (até 2x). Só ajusta linhas com
    PL > 0 e preço disponível.
    """
    qt    = np.asarray(qt,    dtype="float64").copy()
    pl    = np.asarray(pl,    dtype="float64")
    preco = np.asarray(preco, dtype="float64")
    for _ in range(2):
        pb = np.where(
            (pl > 0) & np.isfinite(preco) & np.isfinite(qt) & (qt > 0),
            (preco * qt) / pl, np.nan,
        )
        corrige = np.isfinite(pb) & (pb < limite_pb)
        qt = np.where(corrige, qt * 1000.0, qt)
    return qt


def _cagr_5anos(df: pd.DataFrame) -> pd.DataFrame:
    """
    CAGR de lucro e receita: dezembro X vs dezembro X-5. Retorna
    [CNPJ_CIA, data_cagr, cagr_lucro, cagr_receita] (data_cagr = 31/dez do ano X)
    para propagar via merge_asof (carry-forward).
    """
    dez = df[df["DT_FIM_EXERC"].dt.month == 12].copy()
    if dez.empty:
        return pd.DataFrame(columns=["CNPJ_CIA", "data_cagr", "cagr_lucro", "cagr_receita"])
    dez["ano"] = dez["DT_FIM_EXERC"].dt.year
    atual = dez.drop_duplicates(["CNPJ_CIA", "ano"])
    base = atual.rename(columns={"ano": "ano_base"})
    base["ano"] = base["ano_base"] + 5
    cols_base = {"ll_anualizado": "ll_base", "receita_anualizada": "rec_base"}
    base = base.rename(columns=cols_base)
    keep = ["CNPJ_CIA", "ano"] + [v for k, v in cols_base.items() if v in base.columns]
    m = atual.merge(base[keep], on=["CNPJ_CIA", "ano"], how="left")

    def _cagr(a, b):
        ok = (b > 0) & (a > 0)
        return np.where(ok, (a / b) ** (1 / 5) - 1, np.nan)

    res = m[["CNPJ_CIA", "ano"]].copy()
    if "ll_base" in m:
        res["cagr_lucro"] = _cagr(m["ll_anualizado"], m["ll_base"])
    if "rec_base" in m:
        res["cagr_receita"] = _cagr(m["receita_anualizada"], m["rec_base"])
    res["data_cagr"] = pd.to_datetime(res["ano"].astype(str) + "-12-31")
    return res.drop(columns=["ano"])


def calcular_fatores(base: pd.DataFrame, df_precos: pd.DataFrame,
                     fatores: list[str]) -> pd.DataFrame:
    """
    base     : linhas dos ativos escolhidos (TODAS as datas — CAGR usa histórico).
    df_precos: [ticker, DT_FIM_EXERC, preco] dos ativos no período (pode ser vazio).
    fatores  : chaves de FATORES selecionadas.
    """
    df = base.copy()
    if df_precos is not None and not df_precos.empty:
        df = df.merge(df_precos, on=["ticker", "DT_FIM_EXERC"], how="left")
    if "preco" not in df.columns:
        df["preco"] = np.nan

    # Corrige escala de qt_acoes (composição às vezes em MILHARES) via P/B implícito
    if "qt_acoes" in df.columns and "pl" in df.columns:
        df["qt_acoes"] = _normalizar_escala_acoes(df["qt_acoes"], df["pl"], df["preco"])

    g = lambda c: df[c] if c in df.columns else pd.Series(np.nan, index=df.index)

    # CAGR (usa histórico completo): calculado em dezembro e propagado aos
    # trimestres seguintes via merge_asof (carry-forward; só dezembros passados).
    if "cagr_lucro" in fatores or "cagr_receita" in fatores:
        cagr = _cagr_5anos(df).dropna(subset=["data_cagr"]).sort_values("data_cagr")
        df = df.sort_values("DT_FIM_EXERC")
        if not cagr.empty:
            df = pd.merge_asof(
                df, cagr, left_on="DT_FIM_EXERC", right_on="data_cagr",
                by="CNPJ_CIA", direction="backward",
            ).drop(columns=["data_cagr"])

    saida = {}
    qt = g("qt_acoes")
    ll = g("ll_anualizado")
    pl = g("pl")
    preco = g("preco")
    div_liq = g("divida_liquida")

    # intermediários reutilizados
    market_cap = preco * qt
    lpa = np.where(qt > 0, ll / qt, np.nan)
    vpa = np.where(qt > 0, pl / qt, np.nan)
    graham = np.where((lpa > 0) & (vpa > 0), np.sqrt(22.5 * lpa * vpa), np.nan)

    for f in fatores:
        if f == "roic":
            if "roic" in df.columns:
                saida["roic"] = g("roic")
            else:
                saida["roic"] = np.where(g("capital_investido") != 0,
                                         g("nopat") / g("capital_investido"), np.nan)
        elif f == "roe":
            saida["roe"] = np.where(pl > 0, ll / pl, np.nan)
        elif f == "margem_liquida":
            rec = g("receita_anualizada")
            saida["margem_liquida"] = np.where(rec.abs() > 0, ll / rec, np.nan)
        elif f == "ebitda":
            saida["ebitda"] = g("ebitda")
        elif f == "dl_ebitda":
            eb = g("ebitda")
            saida["dl_ebitda"] = np.where(eb > 0, div_liq / eb, np.nan)
        elif f == "divida_liquida":
            saida["divida_liquida"] = div_liq
        elif f in ("cagr_lucro", "cagr_receita"):
            saida[f] = g(f)
        elif f == "preco":
            saida["preco"] = preco
        elif f == "market_cap":
            saida["market_cap"] = market_cap
        elif f == "ev":
            saida["ev"] = market_cap + div_liq
        elif f == "ev_ebit":
            eb = g("ebit_anualizado")
            saida["ev_ebit"] = np.where(eb != 0, (market_cap + div_liq) / eb, np.nan)
        elif f == "p_l":
            saida["p_l"] = np.where(ll > 0, market_cap / ll, np.nan)
        elif f == "lpa":
            saida["lpa"] = lpa
        elif f == "vpa":
            saida["vpa"] = vpa
        elif f == "preco_graham":
            saida["preco_graham"] = graham
        elif f == "margem_seguranca":
            saida["margem_seguranca"] = np.where(
                (graham > 0) & preco.notna(), (graham - preco) / graham, np.nan)

    # Métricas que não se aplicam a instituições financeiras → NaN
    if "is_financeiro" in df.columns:
        fin = df["is_financeiro"].astype(str).str.strip().str.lower().isin(
            ["true", "1", "1.0"]).to_numpy()
    else:
        fin = np.zeros(len(df), dtype=bool)
    for col in ("roic", "dl_ebitda", "ev", "ev_ebit"):
        if col in saida:
            saida[col] = np.where(fin, np.nan, np.asarray(saida[col], dtype="float64"))

    ident = ["ticker", "DENOM_CIA"]
    if "SETOR_ATIV" in df.columns:
        ident.append("SETOR_ATIV")
    ident.append("DT_FIM_EXERC")

    out = df[ident].copy()
    for col, val in saida.items():
        out[col] = val
    return out.sort_values(["ticker", "DT_FIM_EXERC"]).reset_index(drop=True)


# ===========================================================================
# Interface
# ===========================================================================

st.title("📥 Download de Fatores Fundamentalistas")
st.caption("Base: ranking_historico.csv (CVM). Selecione período, ativos e fatores.")

# --- Fonte do CSV (arquivo gerado pelo pipeline) ---
if not CAMINHO_PADRAO.exists():
    st.error(f"Arquivo não encontrado: {CAMINHO_PADRAO}. "
             "Rode o pipeline do notebook (fund_tri.ipynb) para gerar ranking_historico.csv.")
    st.stop()
base = carregar_base(str(CAMINHO_PADRAO))
st.sidebar.success(f"Base: {CAMINHO_PADRAO.name}")

if "DENOM_CIA" not in base.columns or "CNPJ_CIA" not in base.columns:
    st.error("O CSV precisa ter as colunas CNPJ_CIA e DENOM_CIA.")
    st.stop()

colunas = set(base.columns)

# --- Mapeamento de tickers ---
empresas = base[["CNPJ_CIA", "DENOM_CIA"]].drop_duplicates("CNPJ_CIA").dropna()
mapa = mapear_tickers(empresas)
base = base.merge(mapa[["CNPJ_CIA", "ticker"]], on="CNPJ_CIA", how="left")

tickers_disp = sorted(mapa["ticker"].dropna().unique().tolist())

# --- Período (apenas a partir de 2020Q1 — antes não há CAGR de 5 anos) ---
PISO = pd.Period("2020Q1", "Q")
trimestres = sorted(t for t in base["DT_FIM_EXERC"].dt.to_period("Q").unique() if t >= PISO)
trimestres_str = [str(t) for t in trimestres]

with st.sidebar:
    st.header("2. Período")
    if not trimestres_str:
        st.error("A base não possui trimestres a partir de 2020Q1.")
        st.stop()
    if len(trimestres_str) >= 2:
        ini, fim = st.select_slider(
            "Faixa de trimestres (a partir de 2020Q1)",
            options=trimestres_str,
            value=(trimestres_str[0], trimestres_str[-1]),
        )
    else:
        ini = fim = trimestres_str[0]
        st.write(f"Trimestre único: {ini}")

    st.header("3. Ativos")
    if not tickers_disp:
        st.error("Nenhum ticker mapeado (brapi indisponível?).")
        st.stop()
    sel_todos = st.checkbox("Selecionar todos os ativos mapeados", value=False)
    tickers_sel = tickers_disp if sel_todos else st.multiselect(
        "Tickers", tickers_disp, default=tickers_disp[:5] if tickers_disp else [])

    st.header("4. Fatores")
    fatores_sel = []
    for grupo in ["Fundamentalista", "De preço"]:
        st.markdown(f"**{grupo}**")
        for chave, meta in FATORES.items():
            if meta["grupo"] != grupo:
                continue
            disp = fator_disponivel(chave, colunas)
            marcado = st.checkbox(
                meta["label"] + ("" if disp else "  _(faltam colunas no CSV)_"),
                value=False, disabled=not disp, key=f"f_{chave}",
            )
            if marcado and disp:
                fatores_sel.append(chave)

    gerar = st.button("🚀 Gerar planilha", type="primary", use_container_width=True)

# --- Execução ---
if not gerar:
    st.info("Configure as opções na barra lateral e clique em **Gerar planilha**.")
    st.stop()

if not tickers_sel:
    st.error("Selecione ao menos um ativo.")
    st.stop()
if not fatores_sel:
    st.error("Selecione ao menos um fator.")
    st.stop()

# Filtra ativos (todas as datas — CAGR precisa do histórico)
base_sel = base[base["ticker"].isin(tickers_sel)].copy()

# Período de exibição
p_ini, p_fim = pd.Period(ini, "Q"), pd.Period(fim, "Q")
periodo_mask = base_sel["DT_FIM_EXERC"].dt.to_period("Q").between(p_ini, p_fim)

# Preços (só se houver fator de preço) — apenas datas do período selecionado
precisa_preco = any(FATORES[f]["precisa_preco"] for f in fatores_sel)
df_precos = pd.DataFrame(columns=["ticker", "DT_FIM_EXERC", "preco"])
if precisa_preco:
    datas_periodo = base_sel.loc[periodo_mask, "DT_FIM_EXERC"].dropna().unique()
    barra = st.progress(0.0, text="Buscando cotações...")
    frames = []
    for i, tk in enumerate(tickers_sel):
        dts = base_sel.loc[base_sel["ticker"] == tk, "DT_FIM_EXERC"].dropna().unique()
        dts = [d for d in dts if d in datas_periodo]
        if dts:
            p = buscar_precos(tk, tuple(pd.to_datetime(dts).tolist()))
            if not p.empty:
                p["ticker"] = tk
                frames.append(p)
        barra.progress((i + 1) / len(tickers_sel), text=f"Cotações {i+1}/{len(tickers_sel)}")
    barra.empty()
    if frames:
        df_precos = pd.concat(frames, ignore_index=True)
        df_precos["DT_FIM_EXERC"] = pd.to_datetime(df_precos["DT_FIM_EXERC"])

# Calcula sobre o histórico e depois restringe ao período
resultado = calcular_fatores(base_sel, df_precos, fatores_sel)
resultado = resultado[resultado["DT_FIM_EXERC"].dt.to_period("Q").between(p_ini, p_fim)]

st.subheader(f"Resultado — {len(resultado):,} linhas ({ini} a {fim})")
st.dataframe(resultado, use_container_width=True, hide_index=True)

csv = resultado.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "💾 Baixar CSV", data=csv,
    file_name=f"fatores_{ini}_{fim}.csv", mime="text/csv",
    use_container_width=True,
)
