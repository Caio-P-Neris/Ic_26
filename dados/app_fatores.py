"""
app_fatores.py
==============
Interface web (Streamlit) para baixar fatores fundamentalistas JÁ CALCULADOS
pelo pipeline (ranking.csv). O usuário escolhe período, ativos (por ticker) e
quais fatores quer, e baixa o recorte em CSV.

Os fatores já vêm prontos do ranking.csv (gerado por processar_ano +
adicionar_ev + calcular_metricas), incluindo: correção de escala de ações,
CAGR (carry-forward) e tratamento do setor financeiro — para bancos, ROIC,
DL/EBITDA, EV e EV/EBIT já vêm vazios (não se aplicam). Por isso o app só
filtra e seleciona colunas — não usa rede nem recalcula nada.

Como rodar:
    pip install streamlit pandas
    streamlit run app_fatores.py
"""

from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Fatores Fundamentalistas — CVM", layout="wide")

CAMINHO_PADRAO = Path(__file__).parent / "ranking.csv"
PISO = pd.Period("2020Q1", "Q")   # antes disso não há qt_acoes (composição de capital)

# Fatores selecionáveis → cada chave é o nome da coluna no ranking.csv
FATORES = {
    # Fundamentalistas
    "roic":             ("ROIC",                  "Fundamentalista"),
    "roe":              ("ROE",                   "Fundamentalista"),
    "margem_liquida":   ("Margem Líquida",        "Fundamentalista"),
    "ebitda":           ("EBITDA",                "Fundamentalista"),
    "dl_ebitda":        ("Dív. Líq. / EBITDA",    "Fundamentalista"),
    "divida_liquida":   ("Dívida Líquida",        "Fundamentalista"),
    "cagr_lucro":       ("CAGR Lucro 5a",         "Fundamentalista"),
    "cagr_receita":     ("CAGR Receita 5a",       "Fundamentalista"),
    # Preço / mercado
    "preco":            ("Preço (fechamento)",    "Preço / mercado"),
    "market_cap":       ("Market Cap",            "Preço / mercado"),
    "ev":               ("Enterprise Value (EV)", "Preço / mercado"),
    "ev_ebit":          ("EV/EBIT",               "Preço / mercado"),
    "p_l":              ("P/L",                   "Preço / mercado"),
    "lpa":              ("LPA",                   "Preço / mercado"),
    "vpa":              ("VPA",                   "Preço / mercado"),
    "preco_graham":     ("Preço Justo (Graham)",  "Preço / mercado"),
    "margem_seguranca": ("Margem de Segurança",   "Preço / mercado"),
}
GRUPOS = ["Fundamentalista", "Preço / mercado"]


@st.cache_data(show_spinner=False)
def carregar_base(caminho: str) -> pd.DataFrame:
    df = pd.read_csv(caminho)
    df.columns = [c.lstrip("﻿") for c in df.columns]
    df["DT_FIM_EXERC"] = pd.to_datetime(df["DT_FIM_EXERC"], errors="coerce")
    return df.dropna(subset=["DT_FIM_EXERC"])


# ===========================================================================
# Interface
# ===========================================================================

st.title("📥 Download de Fatores Fundamentalistas")
st.caption("Fonte: ranking.csv (fatores já calculados pelo pipeline). "
           "Selecione período, ativos e fatores e baixe em CSV.")

if not CAMINHO_PADRAO.exists():
    st.error(f"Arquivo não encontrado: {CAMINHO_PADRAO}. "
             "Rode o pipeline do notebook (fund_tri.ipynb) para gerar ranking.csv.")
    st.stop()

base = carregar_base(str(CAMINHO_PADRAO))

for col in ("ticker", "DT_FIM_EXERC"):
    if col not in base.columns:
        st.error(f"O ranking.csv não tem a coluna obrigatória '{col}'.")
        st.stop()

colunas = set(base.columns)

with st.sidebar:
    st.header("1. Base de dados")
    st.success(f"Base: {CAMINHO_PADRAO.name} ({len(base):,} linhas)")

    # --- Período (a partir de 2020Q1 — antes não há ações/Graham/CAGR) ---
    st.header("2. Período")
    trimestres = sorted(t for t in base["DT_FIM_EXERC"].dt.to_period("Q").unique() if t >= PISO)
    trimestres_str = [str(t) for t in trimestres]
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

    # --- Ativos (por ticker, já presente no ranking.csv) ---
    st.header("3. Ativos")
    tickers_disp = sorted(base["ticker"].dropna().unique().tolist())
    if not tickers_disp:
        st.error("Nenhum ticker na base.")
        st.stop()
    sel_todos = st.checkbox("Selecionar todos os ativos", value=False)
    tickers_sel = tickers_disp if sel_todos else st.multiselect(
        "Tickers", tickers_disp, default=tickers_disp[:5])

    # --- Fatores ---
    st.header("4. Fatores")
    fatores_sel = []
    for grupo in GRUPOS:
        st.markdown(f"**{grupo}**")
        for chave, (label, g) in FATORES.items():
            if g != grupo:
                continue
            disp = chave in colunas
            marcado = st.checkbox(
                label + ("" if disp else "  _(ausente no CSV)_"),
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

p_ini, p_fim = pd.Period(ini, "Q"), pd.Period(fim, "Q")
mask = (
    base["ticker"].isin(tickers_sel)
    & base["DT_FIM_EXERC"].dt.to_period("Q").between(p_ini, p_fim)
)

ident = [c for c in ["ticker", "DENOM_CIA", "SETOR_ATIV", "is_financeiro", "DT_FIM_EXERC"]
         if c in base.columns]
resultado = (
    base.loc[mask, ident + fatores_sel]
        .sort_values(["ticker", "DT_FIM_EXERC"])
        .reset_index(drop=True)
)

st.subheader(f"Resultado — {len(resultado):,} linhas ({ini} a {fim})")
st.dataframe(resultado, use_container_width=True, hide_index=True)

csv = resultado.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "💾 Baixar CSV", data=csv,
    file_name=f"fatores_{ini}_{fim}.csv", mime="text/csv",
    use_container_width=True,
)
