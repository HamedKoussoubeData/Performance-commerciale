from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots


# -------------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------------
st.set_page_config(
    page_title="Insurance Brokerage Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent
FILE_NAME = "BASE_COMMERCIALE_PUBLIQUE_ANONYMISEE.xlsx"

POSSIBLE_DATA_PATHS = [
    APP_DIR / "data" / FILE_NAME,
    Path.cwd() / "data" / FILE_NAME,
    APP_DIR / FILE_NAME,
    Path.cwd() / FILE_NAME,
]

DATA_PATH = next(
    (path for path in POSSIBLE_DATA_PATHS if path.exists()),
    None,
)

COLORS = {
    "navy": "#0B1F3A",
    "blue": "#2563EB",
    "blue_light": "#DBEAFE",
    "teal": "#0F9D8B",
    "orange": "#F59E0B",
    "red": "#DC2626",
    "green": "#16A34A",
    "gray_900": "#111827",
    "gray_700": "#374151",
    "gray_500": "#6B7280",
    "gray_300": "#D1D5DB",
    "gray_100": "#F3F4F6",
    "white": "#FFFFFF",
}

PLOTLY_COLORS = [
    COLORS["blue"],
    COLORS["teal"],
    COLORS["orange"],
    "#7C3AED",
    "#0891B2",
    "#DB2777",
    "#65A30D",
    "#EA580C",
]


# -------------------------------------------------------------------
# STYLE
# -------------------------------------------------------------------
st.markdown(
    f"""
    <style>
        .stApp {{
            background: #F6F8FC;
        }}

        .block-container {{
            max-width: 1500px;
            padding-top: 1.4rem;
            padding-bottom: 2rem;
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, {COLORS["navy"]} 0%, #102B50 100%);
        }}

        [data-testid="stSidebar"] * {{
            color: #FFFFFF;
        }}

        [data-testid="stSidebar"] .stMultiSelect div[data-baseweb="select"] > div,
        [data-testid="stSidebar"] .stDateInput div[data-baseweb="input"] > div {{
            background-color: rgba(255,255,255,0.10);
            border-color: rgba(255,255,255,0.25);
        }}

        .hero {{
            background: linear-gradient(120deg, {COLORS["navy"]} 0%, #173E73 65%, {COLORS["blue"]} 130%);
            border-radius: 20px;
            padding: 30px 34px;
            color: white;
            box-shadow: 0 14px 38px rgba(11, 31, 58, 0.18);
            margin-bottom: 22px;
        }}

        .hero-eyebrow {{
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            opacity: 0.78;
            margin-bottom: 8px;
        }}

        .hero-title {{
            font-size: 2.05rem;
            line-height: 1.15;
            font-weight: 800;
            margin: 0;
        }}

        .hero-subtitle {{
            margin-top: 10px;
            font-size: 1rem;
            line-height: 1.55;
            opacity: 0.88;
            max-width: 900px;
        }}

        .hero-meta {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            margin-top: 16px;
            padding: 7px 11px;
            border-radius: 999px;
            background: rgba(255,255,255,0.12);
            font-size: 0.83rem;
        }}

        .section-title {{
            font-size: 1.18rem;
            font-weight: 750;
            color: {COLORS["navy"]};
            margin-top: 0.2rem;
            margin-bottom: 0.25rem;
        }}

        .section-note {{
            color: {COLORS["gray_500"]};
            font-size: 0.9rem;
            margin-bottom: 0.9rem;
        }}

        [data-testid="stMetric"] {{
            background: white;
            border: 1px solid #E5EAF2;
            border-radius: 16px;
            padding: 15px 17px;
            box-shadow: 0 7px 20px rgba(15, 23, 42, 0.055);
        }}

        [data-testid="stMetricLabel"] {{
            color: {COLORS["gray_500"]};
            font-weight: 650;
        }}

        [data-testid="stMetricValue"] {{
            color: {COLORS["navy"]};
            font-weight: 800;
        }}

        [data-testid="stTabs"] button {{
            font-weight: 700;
            padding-left: 18px;
            padding-right: 18px;
        }}

        .insight-box {{
            background: white;
            border: 1px solid #E5EAF2;
            border-left: 5px solid {COLORS["blue"]};
            border-radius: 14px;
            padding: 15px 17px;
            margin: 8px 0 12px 0;
            color: {COLORS["gray_700"]};
            box-shadow: 0 5px 16px rgba(15, 23, 42, 0.04);
        }}

        .warning-box {{
            background: #FFFBEB;
            border: 1px solid #FDE68A;
            border-left: 5px solid {COLORS["orange"]};
            border-radius: 14px;
            padding: 15px 17px;
            color: #78350F;
            margin: 8px 0 12px 0;
        }}

        .footer {{
            margin-top: 32px;
            padding-top: 18px;
            border-top: 1px solid #E5E7EB;
            color: {COLORS["gray_500"]};
            font-size: 0.82rem;
            text-align: center;
        }}

        div[data-testid="stDataFrame"] {{
            background: white;
            border-radius: 14px;
            overflow: hidden;
            border: 1px solid #E5EAF2;
        }}

        .stDownloadButton button {{
            background: {COLORS["navy"]};
            color: white;
            border: none;
            border-radius: 10px;
            font-weight: 700;
        }}

        .stDownloadButton button:hover {{
            background: {COLORS["blue"]};
            color: white;
            border: none;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------------------------
# FONCTIONS
# -------------------------------------------------------------------
def format_fcfa(value: float, compact: bool = False) -> str:
    """Formate un montant en FCFA avec une présentation française."""
    if pd.isna(value):
        return "—"

    value = float(value)

    if compact:
        absolute = abs(value)
        if absolute >= 1_000_000_000:
            return f"{value / 1_000_000_000:.2f} Md FCFA".replace(".", ",")
        if absolute >= 1_000_000:
            return f"{value / 1_000_000:.1f} M FCFA".replace(".", ",")
        if absolute >= 1_000:
            return f"{value / 1_000:.1f} k FCFA".replace(".", ",")

    return f"{value:,.0f} FCFA".replace(",", " ")


def format_integer(value: float) -> str:
    if pd.isna(value):
        return "—"
    return f"{int(value):,}".replace(",", " ")


def style_figure(
    fig: go.Figure,
    title: str | None = None,
    height: int = 430,
    show_legend: bool = True,
) -> go.Figure:
    fig.update_layout(
        title={
            "text": title or "",
            "x": 0.01,
            "xanchor": "left",
            "font": {"size": 18, "color": COLORS["navy"]},
        },
        height=height,
        margin={"l": 15, "r": 15, "t": 58 if title else 25, "b": 15},
        paper_bgcolor=COLORS["white"],
        plot_bgcolor=COLORS["white"],
        font={"family": "Arial, sans-serif", "color": COLORS["gray_700"]},
        legend={
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "right",
            "x": 1,
        },
        showlegend=show_legend,
        hoverlabel={"bgcolor": "white", "font_size": 13},
    )
    fig.update_xaxes(
        showgrid=False,
        linecolor="#E5E7EB",
        tickfont={"color": COLORS["gray_500"]},
        title_font={"color": COLORS["gray_700"]},
    )
    fig.update_yaxes(
        gridcolor="#EEF2F7",
        zeroline=False,
        tickfont={"color": COLORS["gray_500"]},
        title_font={"color": COLORS["gray_700"]},
    )
    return fig


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    data = pd.read_excel(
        path,
        sheet_name="BASE_PUBLIQUE",
        engine="openpyxl",
    )

    date_columns = [
        "DATE_EFFET",
        "DATE_ECHEANCE",
        "DATE_EMISSION",
        "DATE_ENCAISSEMENT_PRIME",
        "DATE_ENCAISSEMENT_COMMISSION",
        "DATE_PAIEMENT_COMMISSION",
    ]

    numeric_columns = [
        "NUMERO_AVENANT",
        "PRIME_NETTE_FCFA",
        "PRIME_RC_FCFA",
        "TAXE_FCFA",
        "COUT_POLICE_FCFA",
        "PRIME_TTC_FCFA",
        "TAUX_COMMISSION_PCT",
        "COMMISSION_FCFA",
        "FRAIS_ACCESSOIRES_FCFA",
        "FRAIS_GESTION_FCFA",
        "ANNEE_EMISSION",
        "DELAI_PAIEMENT_COMMISSION_JOURS",
    ]

    for column in date_columns:
        if column in data.columns:
            data[column] = pd.to_datetime(data[column], errors="coerce")

    for column in numeric_columns:
        if column in data.columns:
            data[column] = pd.to_numeric(data[column], errors="coerce")

    data["ANNEE_EMISSION"] = data["DATE_EMISSION"].dt.year
    data["MOIS_EMISSION_DATE"] = data["DATE_EMISSION"].dt.to_period("M").dt.to_timestamp()
    data["DELAI_RECALCULE"] = (
        data["DATE_PAIEMENT_COMMISSION"]
        - data["DATE_ENCAISSEMENT_COMMISSION"]
    ).dt.days

    data["FLAG_EFFET_APRES_ECHEANCE"] = (
        data["DATE_EFFET"].notna()
        & data["DATE_ECHEANCE"].notna()
        & (data["DATE_EFFET"] > data["DATE_ECHEANCE"])
    )

    data["FLAG_PAIEMENT_AVANT_ENCAISSEMENT"] = (
        data["DATE_ENCAISSEMENT_COMMISSION"].notna()
        & data["DATE_PAIEMENT_COMMISSION"].notna()
        & (data["DATE_PAIEMENT_COMMISSION"] < data["DATE_ENCAISSEMENT_COMMISSION"])
    )

    return data


def get_previous_period_data(
    data: pd.DataFrame,
    current_start: pd.Timestamp,
    current_end: pd.Timestamp,
) -> pd.DataFrame:
    """Retourne une période antérieure de même durée."""
    duration = current_end - current_start
    previous_end = current_start - pd.Timedelta(days=1)
    previous_start = previous_end - duration

    return data[
        data["DATE_EMISSION"].between(previous_start, previous_end)
    ].copy()


def percentage_delta(current: float, previous: float) -> str | None:
    if previous is None or pd.isna(previous) or previous == 0:
        return None
    return f"{((current - previous) / previous) * 100:+.1f} %".replace(".", ",")


def safe_median(series: pd.Series) -> float:
    valid = series.dropna()
    valid = valid[valid >= 0]
    return valid.median() if not valid.empty else np.nan


# -------------------------------------------------------------------
# CHARGEMENT
# -------------------------------------------------------------------
if DATA_PATH is None:
    st.error("Le fichier Excel anonymisé est introuvable dans le dépôt.")
    st.write("Emplacements recherchés :")
    for path in POSSIBLE_DATA_PATHS:
        st.code(str(path))

    st.write("Fichiers actuellement disponibles :")
    files_available = sorted(
        str(path.relative_to(APP_DIR))
        for path in APP_DIR.rglob("*")
        if path.is_file()
    )
    st.dataframe(
        pd.DataFrame({"FICHIER": files_available}),
        use_container_width=True,
        hide_index=True,
    )
    st.stop()

with st.spinner("Chargement des données…"):
    df = load_data(DATA_PATH)

min_date = df["DATE_EMISSION"].min()
max_date = df["DATE_EMISSION"].max()


# -------------------------------------------------------------------
# SIDEBAR
# -------------------------------------------------------------------
with st.sidebar:
    st.markdown("## Insurance Analytics")
    st.caption("Filtres du tableau de bord")
    st.markdown("---")

    date_selection = st.date_input(
        "Période d’émission",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
        key="date_selection",
    )

    if isinstance(date_selection, tuple) and len(date_selection) == 2:
        selected_start = pd.Timestamp(date_selection[0])
        selected_end = pd.Timestamp(date_selection[1])
    else:
        selected_start = min_date
        selected_end = max_date

    company_values = sorted(df["COMPAGNIE_ID"].dropna().astype(str).unique())
    selected_companies = st.multiselect(
        "Compagnies",
        company_values,
        default=company_values,
        key="companies",
    )

    branch_values = sorted(df["BRANCHE"].dropna().astype(str).unique())
    selected_branches = st.multiselect(
        "Branches",
        branch_values,
        default=branch_values,
        key="branches",
    )

    client_type_values = sorted(df["TYPE_CLIENT"].dropna().astype(str).unique())
    selected_client_types = st.multiselect(
        "Types de clients",
        client_type_values,
        default=client_type_values,
        key="client_types",
    )

    avenant_values = sorted(df["TYPE_AVENANT"].dropna().astype(str).unique())
    selected_avenants = st.multiselect(
        "Types d’opérations",
        avenant_values,
        default=avenant_values,
        key="avenants",
    )

    st.markdown("---")
    st.caption(
        "Données anonymisées et transformées. "
        "Les résultats sont destinés à la démonstration."
    )


# -------------------------------------------------------------------
# FILTRAGE
# -------------------------------------------------------------------
filtered = df[
    df["DATE_EMISSION"].between(selected_start, selected_end)
    & df["COMPAGNIE_ID"].isin(selected_companies)
    & df["BRANCHE"].isin(selected_branches)
    & df["TYPE_CLIENT"].isin(selected_client_types)
    & df["TYPE_AVENANT"].isin(selected_avenants)
].copy()

if filtered.empty:
    st.warning(
        "Aucune observation ne correspond aux filtres sélectionnés. "
        "Élargis la période ou réactive certaines modalités."
    )
    st.stop()

previous_period = get_previous_period_data(
    df[
        df["COMPAGNIE_ID"].isin(selected_companies)
        & df["BRANCHE"].isin(selected_branches)
        & df["TYPE_CLIENT"].isin(selected_client_types)
        & df["TYPE_AVENANT"].isin(selected_avenants)
    ],
    selected_start,
    selected_end,
)


# -------------------------------------------------------------------
# EN-TÊTE
# -------------------------------------------------------------------
st.markdown(
    f"""
    <div class="hero">
        <div class="hero-eyebrow">Portfolio Data Analytics</div>
        <div class="hero-title">Tableau de bord de performance commerciale</div>
        <div class="hero-subtitle">
            Analyse du portefeuille, des primes, des commissions et des délais
            de paiement dans une activité de courtage en assurance.
        </div>
        <div class="hero-meta">
            Période analysée : {selected_start.strftime("%d/%m/%Y")}
            — {selected_end.strftime("%d/%m/%Y")}
            &nbsp;•&nbsp; {format_integer(len(filtered))} opérations
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------------------------
# KPIs
# -------------------------------------------------------------------
total_premium = filtered["PRIME_TTC_FCFA"].sum()
total_commission = filtered["COMMISSION_FCFA"].sum()
observed_rate = (
    total_commission / total_premium * 100
    if total_premium > 0
    else np.nan
)
median_delay = safe_median(filtered["DELAI_RECALCULE"])

previous_premium = previous_period["PRIME_TTC_FCFA"].sum()
previous_commission = previous_period["COMMISSION_FCFA"].sum()
previous_operations = len(previous_period)
previous_clients = previous_period["CLIENT_ID"].nunique()

kpi_cols = st.columns(6)

kpi_cols[0].metric(
    "Primes TTC",
    format_fcfa(total_premium, compact=True),
    percentage_delta(total_premium, previous_premium),
)
kpi_cols[1].metric(
    "Commissions",
    format_fcfa(total_commission, compact=True),
    percentage_delta(total_commission, previous_commission),
)
kpi_cols[2].metric(
    "Taux observé",
    f"{observed_rate:.2f} %".replace(".", ",") if not pd.isna(observed_rate) else "—",
)
kpi_cols[3].metric(
    "Clients",
    format_integer(filtered["CLIENT_ID"].nunique()),
    percentage_delta(filtered["CLIENT_ID"].nunique(), previous_clients),
)
kpi_cols[4].metric(
    "Opérations",
    format_integer(len(filtered)),
    percentage_delta(len(filtered), previous_operations),
)
kpi_cols[5].metric(
    "Délai médian",
    f"{median_delay:.0f} jours" if not pd.isna(median_delay) else "—",
)


# -------------------------------------------------------------------
# ONGLETS
# -------------------------------------------------------------------
overview_tab, performance_tab, recovery_tab, quality_tab, data_tab = st.tabs(
    [
        "Vue d’ensemble",
        "Performance commerciale",
        "Recouvrement",
        "Qualité des données",
        "Explorateur",
    ]
)


# ===================================================================
# ONGLET 1 : VUE D'ENSEMBLE
# ===================================================================
with overview_tab:
    st.markdown(
        '<div class="section-title">Dynamique mensuelle du portefeuille</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">'
        "Les primes sont représentées en barres et les commissions sur l’axe secondaire."
        "</div>",
        unsafe_allow_html=True,
    )

    monthly = (
        filtered.dropna(subset=["MOIS_EMISSION_DATE"])
        .groupby("MOIS_EMISSION_DATE", as_index=False)
        .agg(
            PRIME_TTC_FCFA=("PRIME_TTC_FCFA", "sum"),
            COMMISSION_FCFA=("COMMISSION_FCFA", "sum"),
            NOMBRE_OPERATIONS=("ID_OPERATION", "count"),
        )
    )

    trend_fig = make_subplots(specs=[[{"secondary_y": True}]])
    trend_fig.add_trace(
        go.Bar(
            x=monthly["MOIS_EMISSION_DATE"],
            y=monthly["PRIME_TTC_FCFA"],
            name="Primes TTC",
            marker_color=COLORS["blue"],
            opacity=0.86,
            hovertemplate="<b>%{x|%b %Y}</b><br>Primes : %{y:,.0f} FCFA<extra></extra>",
        ),
        secondary_y=False,
    )
    trend_fig.add_trace(
        go.Scatter(
            x=monthly["MOIS_EMISSION_DATE"],
            y=monthly["COMMISSION_FCFA"],
            name="Commissions",
            mode="lines+markers",
            line={"color": COLORS["teal"], "width": 3},
            marker={"size": 7},
            hovertemplate="<b>%{x|%b %Y}</b><br>Commissions : %{y:,.0f} FCFA<extra></extra>",
        ),
        secondary_y=True,
    )
    trend_fig.update_yaxes(title_text="Primes TTC (FCFA)", secondary_y=False)
    trend_fig.update_yaxes(title_text="Commissions (FCFA)", secondary_y=True)
    trend_fig = style_figure(
        trend_fig,
        title="Évolution mensuelle des primes et commissions",
        height=470,
    )
    st.plotly_chart(trend_fig, use_container_width=True)

    left, right = st.columns([1.05, 0.95])

    with left:
        branch_perf = (
            filtered.groupby("BRANCHE", as_index=False)
            .agg(PRIME_TTC_FCFA=("PRIME_TTC_FCFA", "sum"))
            .nlargest(10, "PRIME_TTC_FCFA")
            .sort_values("PRIME_TTC_FCFA")
        )

        branch_fig = px.bar(
            branch_perf,
            x="PRIME_TTC_FCFA",
            y="BRANCHE",
            orientation="h",
            color_discrete_sequence=[COLORS["blue"]],
        )
        branch_fig.update_traces(
            hovertemplate="<b>%{y}</b><br>Primes : %{x:,.0f} FCFA<extra></extra>"
        )
        branch_fig = style_figure(
            branch_fig,
            title="Principales branches",
            height=430,
            show_legend=False,
        )
        branch_fig.update_xaxes(title="Primes TTC (FCFA)")
        branch_fig.update_yaxes(title="")
        st.plotly_chart(branch_fig, use_container_width=True)

    with right:
        operation_mix = (
            filtered.groupby("TYPE_AVENANT", as_index=False)
            .agg(NOMBRE_OPERATIONS=("ID_OPERATION", "count"))
            .sort_values("NOMBRE_OPERATIONS", ascending=False)
        )

        donut_fig = px.pie(
            operation_mix,
            names="TYPE_AVENANT",
            values="NOMBRE_OPERATIONS",
            hole=0.58,
            color_discrete_sequence=PLOTLY_COLORS,
        )
        donut_fig.update_traces(
            textposition="inside",
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>Opérations : %{value}<br>Part : %{percent}<extra></extra>",
        )
        donut_fig.add_annotation(
            text=f"<b>{len(filtered):,}</b><br>opérations".replace(",", " "),
            x=0.5,
            y=0.5,
            showarrow=False,
            font={"size": 18, "color": COLORS["navy"]},
        )
        donut_fig = style_figure(
            donut_fig,
            title="Structure des opérations",
            height=430,
        )
        st.plotly_chart(donut_fig, use_container_width=True)

    top_branch = branch_perf.iloc[-1] if not branch_perf.empty else None
    top_company_name = (
        filtered.groupby("COMPAGNIE_ID")["PRIME_TTC_FCFA"]
        .sum()
        .idxmax()
    )

    if top_branch is not None:
        st.markdown(
            f"""
            <div class="insight-box">
                <b>Lecture rapide.</b> La branche dominante sur la période est
                <b>{top_branch["BRANCHE"]}</b>, tandis que
                <b>{top_company_name}</b> concentre le volume de primes le plus élevé.
                Ces résultats doivent être interprétés avec le risque de concentration
                du portefeuille.
            </div>
            """,
            unsafe_allow_html=True,
        )


# ===================================================================
# ONGLET 2 : PERFORMANCE
# ===================================================================
with performance_tab:
    st.markdown(
        '<div class="section-title">Comparaison des partenaires et des clients</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">'
        "Analyse de la contribution aux primes, aux commissions et à la concentration commerciale."
        "</div>",
        unsafe_allow_html=True,
    )

    company_perf = (
        filtered.groupby("COMPAGNIE_ID", as_index=False)
        .agg(
            OPERATIONS=("ID_OPERATION", "count"),
            CLIENTS=("CLIENT_ID", "nunique"),
            PRIMES=("PRIME_TTC_FCFA", "sum"),
            COMMISSIONS=("COMMISSION_FCFA", "sum"),
            DELAI_MEDIAN=("DELAI_RECALCULE", safe_median),
        )
        .sort_values("PRIMES", ascending=False)
    )

    company_fig = go.Figure()
    company_fig.add_trace(
        go.Bar(
            x=company_perf["COMPAGNIE_ID"],
            y=company_perf["PRIMES"],
            name="Primes TTC",
            marker_color=COLORS["blue"],
            hovertemplate="<b>%{x}</b><br>Primes : %{y:,.0f} FCFA<extra></extra>",
        )
    )
    company_fig.add_trace(
        go.Bar(
            x=company_perf["COMPAGNIE_ID"],
            y=company_perf["COMMISSIONS"],
            name="Commissions",
            marker_color=COLORS["teal"],
            hovertemplate="<b>%{x}</b><br>Commissions : %{y:,.0f} FCFA<extra></extra>",
        )
    )
    company_fig.update_layout(barmode="group")
    company_fig = style_figure(
        company_fig,
        title="Primes et commissions par compagnie",
        height=470,
    )
    company_fig.update_xaxes(title="")
    company_fig.update_yaxes(title="Montants (FCFA)")
    st.plotly_chart(company_fig, use_container_width=True)

    left, right = st.columns([1.05, 0.95])

    with left:
        client_value = (
            filtered.groupby("CLIENT_ID", as_index=False)
            .agg(
                PRIMES=("PRIME_TTC_FCFA", "sum"),
                COMMISSIONS=("COMMISSION_FCFA", "sum"),
                OPERATIONS=("ID_OPERATION", "count"),
            )
            .sort_values("PRIMES", ascending=False)
            .reset_index(drop=True)
        )

        client_value["RANG"] = np.arange(1, len(client_value) + 1)
        client_value["PART_CUMULEE"] = (
            client_value["PRIMES"].cumsum()
            / client_value["PRIMES"].sum()
            * 100
        )

        pareto_fig = make_subplots(specs=[[{"secondary_y": True}]])
        pareto_fig.add_trace(
            go.Bar(
                x=client_value["RANG"],
                y=client_value["PRIMES"],
                name="Primes",
                marker_color=COLORS["blue"],
                hovertemplate="Rang %{x}<br>Primes : %{y:,.0f} FCFA<extra></extra>",
            ),
            secondary_y=False,
        )
        pareto_fig.add_trace(
            go.Scatter(
                x=client_value["RANG"],
                y=client_value["PART_CUMULEE"],
                name="Part cumulée",
                line={"color": COLORS["orange"], "width": 3},
                hovertemplate="Rang %{x}<br>Part cumulée : %{y:.1f} %<extra></extra>",
            ),
            secondary_y=True,
        )
        pareto_fig.add_hline(
            y=80,
            line_dash="dash",
            line_color=COLORS["red"],
            secondary_y=True,
        )
        pareto_fig.update_yaxes(title_text="Primes TTC (FCFA)", secondary_y=False)
        pareto_fig.update_yaxes(
            title_text="Part cumulée (%)",
            range=[0, 105],
            secondary_y=True,
        )
        pareto_fig = style_figure(
            pareto_fig,
            title="Concentration du portefeuille clients",
            height=445,
        )
        pareto_fig.update_xaxes(title="Rang des clients")
        st.plotly_chart(pareto_fig, use_container_width=True)

    with right:
        top_clients = client_value.head(10).sort_values("PRIMES")

        client_fig = px.bar(
            top_clients,
            x="PRIMES",
            y="CLIENT_ID",
            orientation="h",
            color="PRIMES",
            color_continuous_scale=["#DBEAFE", COLORS["blue"], COLORS["navy"]],
        )
        client_fig.update_layout(coloraxis_showscale=False)
        client_fig.update_traces(
            hovertemplate="<b>%{y}</b><br>Primes : %{x:,.0f} FCFA<extra></extra>"
        )
        client_fig = style_figure(
            client_fig,
            title="Top 10 clients selon les primes",
            height=445,
            show_legend=False,
        )
        client_fig.update_xaxes(title="Primes TTC (FCFA)")
        client_fig.update_yaxes(title="")
        st.plotly_chart(client_fig, use_container_width=True)

    top_5_share = (
        client_value.head(5)["PRIMES"].sum()
        / client_value["PRIMES"].sum()
        * 100
        if client_value["PRIMES"].sum() > 0
        else np.nan
    )

    st.markdown(
        f"""
        <div class="insight-box">
            <b>Risque de concentration.</b> Les cinq premiers clients représentent
            <b>{top_5_share:.1f} %</b> des primes de la sélection.
            Une part élevée signale une dépendance commerciale qui mérite un suivi.
        </div>
        """.replace(".", ",", 1),
        unsafe_allow_html=True,
    )

    st.dataframe(
        company_perf.rename(
            columns={
                "COMPAGNIE_ID": "Compagnie",
                "OPERATIONS": "Opérations",
                "CLIENTS": "Clients",
                "PRIMES": "Primes TTC (FCFA)",
                "COMMISSIONS": "Commissions (FCFA)",
                "DELAI_MEDIAN": "Délai médian (jours)",
            }
        ),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Primes TTC (FCFA)": st.column_config.NumberColumn(format="%.0f"),
            "Commissions (FCFA)": st.column_config.NumberColumn(format="%.0f"),
            "Délai médian (jours)": st.column_config.NumberColumn(format="%.0f"),
        },
    )


# ===================================================================
# ONGLET 3 : RECOUVREMENT
# ===================================================================
with recovery_tab:
    st.markdown(
        '<div class="section-title">Suivi des délais de paiement des commissions</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">'
        "Seules les opérations disposant de dates exploitables sont incluses dans le calcul des délais."
        "</div>",
        unsafe_allow_html=True,
    )

    valid_delay = filtered[
        filtered["DELAI_RECALCULE"].notna()
        & (filtered["DELAI_RECALCULE"] >= 0)
    ].copy()

    missing_payment = filtered["DATE_PAIEMENT_COMMISSION"].isna()
    missing_payment_count = int(missing_payment.sum())
    missing_payment_commission = filtered.loc[
        missing_payment,
        "COMMISSION_FCFA",
    ].sum()

    late_60_count = int((valid_delay["DELAI_RECALCULE"] > 60).sum())
    late_60_rate = (
        late_60_count / len(valid_delay) * 100
        if len(valid_delay) > 0
        else np.nan
    )
    late_90_rate = (
        (valid_delay["DELAI_RECALCULE"] > 90).mean() * 100
        if len(valid_delay) > 0
        else np.nan
    )

    rec_cols = st.columns(4)
    rec_cols[0].metric(
        "Paiements renseignés",
        format_integer(len(valid_delay)),
    )
    rec_cols[1].metric(
        "Délai médian",
        f"{safe_median(valid_delay['DELAI_RECALCULE']):.0f} jours"
        if not valid_delay.empty
        else "—",
    )
    rec_cols[2].metric(
        "Retards > 60 jours",
        f"{late_60_rate:.1f} %".replace(".", ",")
        if not pd.isna(late_60_rate)
        else "—",
    )
    rec_cols[3].metric(
        "Retards > 90 jours",
        f"{late_90_rate:.1f} %".replace(".", ",")
        if not pd.isna(late_90_rate)
        else "—",
    )

    if not valid_delay.empty:
        valid_delay["CLASSE_DELAI"] = pd.cut(
            valid_delay["DELAI_RECALCULE"],
            bins=[-0.1, 30, 60, 90, np.inf],
            labels=[
                "0 à 30 jours",
                "31 à 60 jours",
                "61 à 90 jours",
                "Plus de 90 jours",
            ],
        )

        delay_summary = (
            valid_delay.groupby("CLASSE_DELAI", observed=False)
            .agg(
                OPERATIONS=("ID_OPERATION", "count"),
                COMMISSIONS=("COMMISSION_FCFA", "sum"),
            )
            .reset_index()
        )

        left, right = st.columns([0.9, 1.1])

        with left:
            delay_fig = px.bar(
                delay_summary,
                x="CLASSE_DELAI",
                y="OPERATIONS",
                color="CLASSE_DELAI",
                color_discrete_map={
                    "0 à 30 jours": COLORS["green"],
                    "31 à 60 jours": COLORS["blue"],
                    "61 à 90 jours": COLORS["orange"],
                    "Plus de 90 jours": COLORS["red"],
                },
            )
            delay_fig.update_traces(
                hovertemplate="<b>%{x}</b><br>Opérations : %{y}<extra></extra>"
            )
            delay_fig = style_figure(
                delay_fig,
                title="Répartition par classe de délai",
                height=430,
                show_legend=False,
            )
            delay_fig.update_xaxes(title="")
            delay_fig.update_yaxes(title="Nombre d’opérations")
            st.plotly_chart(delay_fig, use_container_width=True)

        with right:
            company_delay = (
                valid_delay.groupby("COMPAGNIE_ID", as_index=False)
                .agg(
                    DELAI_MEDIAN=("DELAI_RECALCULE", "median"),
                    OPERATIONS=("ID_OPERATION", "count"),
                )
                .sort_values("DELAI_MEDIAN")
            )

            company_delay_fig = px.bar(
                company_delay,
                x="DELAI_MEDIAN",
                y="COMPAGNIE_ID",
                orientation="h",
                color="DELAI_MEDIAN",
                color_continuous_scale=[
                    COLORS["blue_light"],
                    COLORS["orange"],
                    COLORS["red"],
                ],
            )
            company_delay_fig.update_layout(coloraxis_showscale=False)
            company_delay_fig.update_traces(
                hovertemplate="<b>%{y}</b><br>Délai médian : %{x:.0f} jours<extra></extra>"
            )
            company_delay_fig = style_figure(
                company_delay_fig,
                title="Délai médian par compagnie",
                height=430,
                show_legend=False,
            )
            company_delay_fig.update_xaxes(title="Délai médian (jours)")
            company_delay_fig.update_yaxes(title="")
            st.plotly_chart(company_delay_fig, use_container_width=True)

    st.markdown(
        f"""
        <div class="warning-box">
            <b>Point de vigilance.</b> {format_integer(missing_payment_count)} opérations
            n’ont pas de date de paiement renseignée, pour
            <b>{format_fcfa(missing_payment_commission)}</b> de commissions associées.
            Cette absence ne permet pas de conclure automatiquement à un impayé :
            une vérification opérationnelle reste nécessaire.
        </div>
        """,
        unsafe_allow_html=True,
    )


# ===================================================================
# ONGLET 4 : QUALITÉ
# ===================================================================
with quality_tab:
    st.markdown(
        '<div class="section-title">Audit de qualité des données</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">'
        "Les anomalies sont conservées afin de documenter la qualité de la source et les limites de l’analyse."
        "</div>",
        unsafe_allow_html=True,
    )

    duplicate_ids = int(filtered["ID_OPERATION"].duplicated().sum())
    effect_after_expiry = int(filtered["FLAG_EFFET_APRES_ECHEANCE"].sum())
    payment_before_collection = int(
        filtered["FLAG_PAIEMENT_AVANT_ENCAISSEMENT"].sum()
    )
    missing_dates = int(filtered["DATE_PAIEMENT_COMMISSION"].isna().sum())

    quality_cols = st.columns(4)
    quality_cols[0].metric("ID dupliqués", format_integer(duplicate_ids))
    quality_cols[1].metric(
        "Effet après échéance",
        format_integer(effect_after_expiry),
    )
    quality_cols[2].metric(
        "Paiement avant encaissement",
        format_integer(payment_before_collection),
    )
    quality_cols[3].metric(
        "Dates de paiement absentes",
        format_integer(missing_dates),
    )

    missing_report = (
        filtered.isna()
        .mean()
        .mul(100)
        .sort_values(ascending=False)
        .rename("TAUX_MANQUANT_PCT")
        .reset_index()
        .rename(columns={"index": "VARIABLE"})
    )
    missing_report = missing_report[missing_report["TAUX_MANQUANT_PCT"] > 0].head(15)

    left, right = st.columns([1.05, 0.95])

    with left:
        missing_fig = px.bar(
            missing_report.sort_values("TAUX_MANQUANT_PCT"),
            x="TAUX_MANQUANT_PCT",
            y="VARIABLE",
            orientation="h",
            color_discrete_sequence=[COLORS["orange"]],
        )
        missing_fig.update_traces(
            hovertemplate="<b>%{y}</b><br>Valeurs manquantes : %{x:.1f} %<extra></extra>"
        )
        missing_fig = style_figure(
            missing_fig,
            title="Variables les plus incomplètes",
            height=470,
            show_legend=False,
        )
        missing_fig.update_xaxes(title="Taux de valeurs manquantes (%)")
        missing_fig.update_yaxes(title="")
        st.plotly_chart(missing_fig, use_container_width=True)

    with right:
        quality_summary = pd.DataFrame(
            {
                "Contrôle": [
                    "Identifiants d’opération dupliqués",
                    "Date d’effet après date d’échéance",
                    "Paiement avant encaissement",
                    "Date de paiement non renseignée",
                ],
                "Nombre": [
                    duplicate_ids,
                    effect_after_expiry,
                    payment_before_collection,
                    missing_dates,
                ],
            }
        )
        quality_summary["Part (%)"] = (
            quality_summary["Nombre"] / len(filtered) * 100
        ).round(2)

        st.markdown("#### Synthèse des contrôles")
        st.dataframe(
            quality_summary,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Part (%)": st.column_config.ProgressColumn(
                    min_value=0,
                    max_value=100,
                    format="%.2f %%",
                )
            },
        )

        st.markdown(
            """
            <div class="insight-box">
                <b>Bonne pratique.</b> Les lignes anormales ne doivent pas être
                supprimées sans justification. Il est préférable de créer des
                indicateurs de contrôle, puis d’exclure uniquement les cas non
                exploitables pour une analyse précise.
            </div>
            """,
            unsafe_allow_html=True,
        )


# ===================================================================
# ONGLET 5 : EXPLORATEUR
# ===================================================================
with data_tab:
    st.markdown(
        '<div class="section-title">Explorateur de données</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">'
        "Recherche, tri et export des données correspondant aux filtres du tableau de bord."
        "</div>",
        unsafe_allow_html=True,
    )

    search_term = st.text_input(
        "Rechercher un identifiant, une branche ou un type d’opération",
        placeholder="Ex. CLIENT_0025, AUTOMOBILE, RENOUVELLEMENT…",
    )

    explorer = filtered.copy()

    if search_term:
        searchable_columns = [
            "ID_OPERATION",
            "CLIENT_ID",
            "COMPAGNIE_ID",
            "POLICE_ID",
            "BRANCHE",
            "CATEGORIE",
            "TYPE_AVENANT",
        ]

        search_mask = pd.Series(False, index=explorer.index)
        for column in searchable_columns:
            search_mask = search_mask | (
                explorer[column]
                .astype(str)
                .str.contains(search_term, case=False, na=False)
            )

        explorer = explorer[search_mask]

    display_columns = [
        "ID_OPERATION",
        "CLIENT_ID",
        "TYPE_CLIENT",
        "COMPAGNIE_ID",
        "POLICE_ID",
        "BRANCHE",
        "CATEGORIE",
        "TYPE_AVENANT",
        "DATE_EMISSION",
        "PRIME_TTC_FCFA",
        "COMMISSION_FCFA",
        "DELAI_RECALCULE",
        "CONTROLE_DATES",
    ]

    available_display_columns = [
        column for column in display_columns if column in explorer.columns
    ]

    st.dataframe(
        explorer[available_display_columns],
        use_container_width=True,
        hide_index=True,
        height=520,
        column_config={
            "DATE_EMISSION": st.column_config.DateColumn(
                "Date d’émission",
                format="DD/MM/YYYY",
            ),
            "PRIME_TTC_FCFA": st.column_config.NumberColumn(
                "Prime TTC",
                format="%.0f FCFA",
            ),
            "COMMISSION_FCFA": st.column_config.NumberColumn(
                "Commission",
                format="%.0f FCFA",
            ),
            "DELAI_RECALCULE": st.column_config.NumberColumn(
                "Délai (jours)",
                format="%.0f",
            ),
        },
    )

    csv_data = explorer.to_csv(
        index=False,
        encoding="utf-8-sig",
    ).encode("utf-8-sig")

    st.download_button(
        label="Télécharger les données filtrées",
        data=csv_data,
        file_name="donnees_commerciales_filtrees.csv",
        mime="text/csv",
        use_container_width=False,
    )


# -------------------------------------------------------------------
# FOOTER
# -------------------------------------------------------------------
st.markdown(
    f"""
    <div class="footer">
        Insurance Brokerage Analytics • Projet portfolio Python & Streamlit<br>
        Source : données commerciales anonymisées et transformées •
        Dernière date d’émission disponible : {max_date.strftime("%d/%m/%Y")}
    </div>
    """,
    unsafe_allow_html=True,
)
