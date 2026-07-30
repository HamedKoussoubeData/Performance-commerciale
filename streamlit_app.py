
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Insurance Brokerage Analytics",
    layout="wide"
)

st.title("Insurance Brokerage Analytics")
st.caption("Données anonymisées et transformées à des fins de démonstration.")

data_candidates = [
    Path("data/BASE_COMMERCIALE_PUBLIQUE_ANONYMISEE.xlsx"),
    Path("BASE_COMMERCIALE_PUBLIQUE_ANONYMISEE.xlsx"),
]

existing = [path for path in data_candidates if path.exists()]

if not existing:
    st.error("Le fichier Excel anonymisé est introuvable.")
    st.stop()

df = pd.read_excel(existing[0], sheet_name="BASE_PUBLIQUE")
df["DATE_EMISSION"] = pd.to_datetime(df["DATE_EMISSION"], errors="coerce")
df["DELAI_PAIEMENT_COMMISSION_JOURS"] = pd.to_numeric(
    df["DELAI_PAIEMENT_COMMISSION_JOURS"],
    errors="coerce"
)

col1, col2, col3, col4 = st.columns(4)
col1.metric("Opérations", f"{len(df):,}".replace(",", " "))
col2.metric("Clients", f"{df['CLIENT_ID'].nunique():,}".replace(",", " "))
col3.metric(
    "Primes TTC",
    f"{df['PRIME_TTC_FCFA'].sum():,.0f} FCFA".replace(",", " ")
)
col4.metric(
    "Commissions",
    f"{df['COMMISSION_FCFA'].sum():,.0f} FCFA".replace(",", " ")
)

st.sidebar.header("Filtres")

company_options = sorted(df["COMPAGNIE_ID"].dropna().unique())
selected_companies = st.sidebar.multiselect(
    "Compagnies",
    company_options,
    default=company_options
)

branch_options = sorted(df["BRANCHE"].dropna().unique())
selected_branches = st.sidebar.multiselect(
    "Branches",
    branch_options,
    default=branch_options
)

filtered = df[
    df["COMPAGNIE_ID"].isin(selected_companies)
    & df["BRANCHE"].isin(selected_branches)
].copy()

monthly = (
    filtered.dropna(subset=["DATE_EMISSION"])
    .groupby(pd.Grouper(key="DATE_EMISSION", freq="MS"))
    .agg(
        PRIME_TTC_FCFA=("PRIME_TTC_FCFA", "sum"),
        COMMISSION_FCFA=("COMMISSION_FCFA", "sum"),
    )
    .reset_index()
)

st.subheader("Évolution mensuelle")
fig = px.line(
    monthly,
    x="DATE_EMISSION",
    y=["PRIME_TTC_FCFA", "COMMISSION_FCFA"],
    markers=True
)
st.plotly_chart(fig, use_container_width=True)

left, right = st.columns(2)

with left:
    st.subheader("Primes par compagnie")
    company = (
        filtered.groupby("COMPAGNIE_ID")["PRIME_TTC_FCFA"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    st.plotly_chart(
        px.bar(company, x="COMPAGNIE_ID", y="PRIME_TTC_FCFA"),
        use_container_width=True
    )

with right:
    st.subheader("Primes par branche")
    branch = (
        filtered.groupby("BRANCHE")["PRIME_TTC_FCFA"]
        .sum()
        .nlargest(12)
        .sort_values()
        .reset_index()
    )
    st.plotly_chart(
        px.bar(
            branch,
            x="PRIME_TTC_FCFA",
            y="BRANCHE",
            orientation="h"
        ),
        use_container_width=True
    )

st.subheader("Données filtrées")
st.dataframe(filtered, use_container_width=True)
