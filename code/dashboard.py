#!/usr/bin/env python3
"""
Streamlit Dashboard — Delphi W1 Completion Monitoring
======================================================
Real-time monitoring of expert response completion rates.

Usage:
    streamlit run code/dashboard.py
    
    Or with custom port:
    streamlit run code/dashboard.py --server.port 8502    
    With expert exclusions:
    streamlit run code/dashboard.py -- --exclude-experts 001PM,001XX

Note: Streamlit arguments go before '--', dashboard arguments after."""

import streamlit as st
import pandas as pd
import sys
import os
from datetime import datetime
import time
from pathlib import Path
import re

try:
    import altair as alt
except ImportError:
    alt = None

# Add code directory to path to import from aggregate_results
sys.path.insert(0, os.path.dirname(__file__))

# Import necessary functions from aggregate_results
try:
    import aggregate_results as aggregate_module
    from aggregate_results import (
        load_config, fetch_all, build_wide, 
        detect_intervention_codes, normalise_columns
    )
except ImportError as e:
    st.error(f"Failed to import from aggregate_results.py: {e}")
    st.stop()

# Load config
cfg = load_config()
TOPIC = cfg.get("TOPIC_CODE", "malaria")
DEFAULT_KOBO_SERVER = cfg.get("KOBO_SERVER", "https://eu.kobotoolbox.org")
DEFAULT_KOBO_TOKEN = cfg.get("KOBO_TOKEN", "")

# Initialize session state for auto-refresh
if "auto_refresh_enabled" not in st.session_state:
    st.session_state.auto_refresh_enabled = False
if "last_auto_refresh_time" not in st.session_state:
    st.session_state.last_auto_refresh_time = 0


def resolve_kobo_credentials(token_override="", server_override=""):
    """Resolve Kobo credentials with priority: UI override > secrets > env > config.env."""
    server = (
        (server_override or "").strip()
        or str(st.secrets.get("KOBO_SERVER", "")).strip()
        or os.getenv("KOBO_SERVER", "").strip()
        or DEFAULT_KOBO_SERVER
    )
    token = (
        (token_override or "").strip()
        or str(st.secrets.get("KOBO_TOKEN", "")).strip()
        or os.getenv("KOBO_TOKEN", "").strip()
        or DEFAULT_KOBO_TOKEN
    )
    return server, token


def resolve_kobo_asset_entries():
    """Collect SUBFORM_ASSET_* entries from config, env, and secrets."""
    assets = {}

    for key, value in cfg.items():
        if key.upper().startswith("SUBFORM_ASSET_") and str(value).strip():
            assets[key] = str(value).strip()

    for key, value in os.environ.items():
        if key.upper().startswith("SUBFORM_ASSET_") and str(value).strip():
            assets[key] = str(value).strip()

    for key in st.secrets.keys():
        if key.upper().startswith("SUBFORM_ASSET_"):
            value = str(st.secrets.get(key, "")).strip()
            if value:
                assets[key] = value

    return assets


def sync_kobo_runtime_config(server, token, asset_entries=()):
    """Propagate resolved credentials/assets into aggregate_results runtime globals."""
    aggregate_module.KOBO_SERVER = server
    aggregate_module.KOBO_TOKEN = token

    for key in [k for k in list(aggregate_module.cfg.keys()) if k.upper().startswith("SUBFORM_ASSET_")]:
        del aggregate_module.cfg[key]

    for key, value in asset_entries:
        aggregate_module.cfg[key] = value


# ═══════════════════════════════════════════════════════════════
# Helper function for code splitting
# ═══════════════════════════════════════════════════════════════

def _split_codes(value):
    """Split comma/semicolon/whitespace-separated codes into a set."""
    if not value or not isinstance(value, str):
        return set()
    codes = re.split(r'[,;\s]+', value.strip())
    return {c.strip().lower() for c in codes if c.strip()}


def load_expected_experts(experts_file=None):
    """Load unique expert entries from experts.txt (ignoring comments/empty lines)."""
    if experts_file is None:
        experts_file = Path(__file__).resolve().parents[1] / "experts.txt"

    experts = set()
    try:
        with open(experts_file, "r", encoding="utf-8") as f:
            for line in f:
                raw_line = line.strip()
                lower_line = raw_line.lower()
                if (
                    not raw_line
                    or raw_line.startswith("#")
                    or "# test" in lower_line
                    or "# ignore" in lower_line
                ):
                    continue
                # Support optional inline comments after '#'
                entry = raw_line.split("#", 1)[0].strip().lower()
                if entry:
                    experts.add(entry)
    except FileNotFoundError:
        return []

    return sorted(experts)


EXPECTED_EXPERTS = load_expected_experts()
N_EXPECTED_EXPERTS = len(EXPECTED_EXPERTS)

# Page configuration
st.set_page_config(
    page_title=f"Delphi W1 — {TOPIC.title()} | Painel de Monitoramento",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


def apply_report_theme():
    """Apply the same font stack and color scheme used in the HTML report."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');

        :root {
            --ink: #f5f2ed;
            --paper: #0f1923;
            --surface: #16212e;
            --accent: #c0392b;
            --accent2: #1a5276;
            --gold: #b7860b;
            --muted: #c3ccd6;
            --border: #2c3a4a;
            --sim: #1a6b3a;
            --sim-bg: #e8f5ee;
            --poss: #7a5c00;
            --poss-bg: #fef9e7;
        }

        html, body, [class*="css"] {
            font-family: 'DM Sans', sans-serif;
            color: var(--ink);
        }

        .stApp {
            background: var(--paper);
            color: var(--ink);
        }

        /* Force readable text on dark background in Streamlit containers */
        .stApp,
        .stApp p,
        .stApp span,
        .stApp div,
        .stApp label,
        .stApp li,
        .stApp small,
        .stMarkdown,
        .stMarkdown p,
        .stCaption,
        [data-testid="stSidebar"] * {
            color: var(--ink);
        }

        h1, h2, h3 {
            font-family: 'DM Serif Display', serif;
            color: var(--ink);
            letter-spacing: 0.01em;
        }

        .stSidebar {
            background: var(--surface);
            border-right: 1px solid var(--border);
        }

        div[data-testid="stMetric"] {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 6px;
            padding: 10px 12px;
        }

        div[data-testid="stMetric"] label {
            color: var(--muted);
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
        }

        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            color: var(--accent);
            font-family: 'DM Serif Display', serif;
        }

        .stButton > button {
            background: var(--accent2);
            color: #ffffff;
            border: 1px solid var(--accent2);
        }

        .stButton > button:hover {
            background: #243246;
            border-color: #243246;
            color: #ffffff;
        }

        .stButton > button,
        .stButton > button * {
            color: #ffffff !important;
        }

        [data-testid="stDataFrame"] {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 6px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════
# Data Fetching & Processing (with caching)
# ═══════════════════════════════════════════════════════════════

@st.cache_data(ttl=300, show_spinner=False)  # Cache for 5 minutes
def fetch_and_process_data(
    force_refresh=False,
    excluded_experts=None,
    kobo_server="",
    kobo_token="",
    asset_entries=(),
):
    """
    Fetch submissions from Kobo API and build coverage matrices.
    Cached to avoid redundant API calls.
    
    Parameters:
        excluded_experts: frozenset of expert codes (lowercase) to filter out
    """
    try:
        import requests
    except ImportError:
        st.error("requests library not installed. Run: pip install requests")
        return None
    
    if not kobo_token:
        st.error("KOBO_TOKEN não configurado (Secrets, variável de ambiente, input manual ou config.env)")
        return None

    sync_kobo_runtime_config(kobo_server, kobo_token, asset_entries=asset_entries)
    
    # Fetch data
    raw = fetch_all(asset_filter=None)
    raw = raw.fillna("")
    
    if raw.empty:
        return {
            "timestamp": datetime.now(),
            "raw": raw,
            "wide": pd.DataFrame(),
            "experts": [],
            "interventions": [],
            "groups": [],
            "n_submissions": 0,
            "n_experts": 0,
            "n_submissions_before_exclusion": 0,
        }
    
    n_before = len(raw)
    
    # Apply expert exclusions if provided
    if excluded_experts:
        excluded_lc = {str(c).strip().lower() for c in excluded_experts}
        raw = raw[
            ~raw["expert_code"].fillna("").astype(str).str.strip().str.lower().isin(excluded_lc)
        ].copy()
    
    # Build wide table
    wide = build_wide(raw)
    
    # Extract unique values
    experts = sorted(raw["expert_code"].dropna().unique())
    groups = sorted(raw["_group"].dropna().unique())
    interventions = sorted(wide["intervention"].dropna().unique()) if not wide.empty else []
    
    return {
        "timestamp": datetime.now(),
        "raw": raw,
        "wide": wide,
        "experts": experts,
        "interventions": interventions,
        "groups": groups,
        "n_submissions": len(raw),
        "n_experts": len(experts),
        "n_submissions_before_exclusion": n_before,
    }


def build_coverage_matrix(wide, experts, interventions):
    """Build expert × intervention completion matrix."""
    if wide.empty or not experts or not interventions:
        return pd.DataFrame()
    
    matrix = []
    for expert in experts:
        row = {"Expert": expert}
        expert_data = wide[wide["expert_code"] == expert]
        for intv in interventions:
            intv_rows = expert_data[expert_data["intervention"] == intv]
            if len(intv_rows) > 0:
                gate = intv_rows.iloc[0].get("gate", "")
                row[intv] = "✓" if gate and str(gate).strip() else "—"
            else:
                row[intv] = "—"
        matrix.append(row)
    
    return pd.DataFrame(matrix)


def build_group_coverage(raw, experts, groups):
    """Build expert × group submission matrix."""
    if raw.empty or not experts or not groups:
        return pd.DataFrame()
    
    matrix = []
    counts = raw.groupby(["expert_code", "_group"]).size().to_dict()
    for expert in experts:
        row = {"Expert": expert}
        for group in groups:
            n = counts.get((expert, group), 0)
            row[group] = "✓" if n == 1 else (f"×{n}" if n > 1 else "—")
        matrix.append(row)
    
    return pd.DataFrame(matrix)


def compute_stats(data):
    """Compute summary statistics."""
    if not data or data["wide"].empty:
        return {
            "response_rate": 0,
            "completed_interventions": 0,
            "total_possible": 0,
            "by_intervention": pd.DataFrame(),
            "by_expert": pd.DataFrame(),
        }
    
    wide = data["wide"]
    n_experts_observed = data["n_experts"]
    n_experts_expected = N_EXPECTED_EXPERTS if N_EXPECTED_EXPERTS > 0 else n_experts_observed
    n_interventions = len(data["interventions"])
    
    # Overall completion
    total_possible = n_experts_expected * n_interventions
    completed = len(wide[wide["gate"].apply(lambda x: bool(x and str(x).strip()))])
    response_rate = (completed / total_possible * 100) if total_possible > 0 else 0
    
    # By intervention
    by_intv = []
    for intv in data["interventions"]:
        intv_data = wide[wide["intervention"] == intv]
        answered = len(intv_data[intv_data["gate"].apply(lambda x: bool(x and str(x).strip()))])
        pct = (answered / n_experts_expected * 100) if n_experts_expected > 0 else 0
        by_intv.append({
            "Intervenção": intv,
            "Respondentes": answered,
            "Total Especialistas": n_experts_expected,
            "Taxa (%)": round(pct, 1)
        })
    
    # By expert
    by_exp = []
    for expert in data["experts"]:
        expert_data = wide[wide["expert_code"] == expert]
        answered = len(expert_data[expert_data["gate"].apply(lambda x: bool(x and str(x).strip()))])
        pct = (answered / n_interventions * 100) if n_interventions > 0 else 0
        by_exp.append({
            "Especialista": expert,
            "Respondidas": answered,
            "Total Intervenções": n_interventions,
            "Taxa (%)": round(pct, 1)
        })
    
    return {
        "response_rate": round(response_rate, 1),
        "completed_interventions": completed,
        "total_possible": total_possible,
        "n_experts_expected": n_experts_expected,
        "n_experts_observed": n_experts_observed,
        "by_intervention": pd.DataFrame(by_intv),
        "by_expert": pd.DataFrame(by_exp),
    }


# ═══════════════════════════════════════════════════════════════
# UI Components
# ═══════════════════════════════════════════════════════════════

def render_header():
    """Render page header."""
    st.title(f"📊 Painel de Monitoramento — Delphi W1 ({TOPIC.title()})")
    st.markdown("Acompanhamento em tempo real das respostas dos especialistas")
    st.divider()


def render_overview_cards(stats):
    """Render top-level overview cards."""
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label="Taxa Global de Resposta",
            value=f"{stats['response_rate']}%",
            delta=None
        )
    
    with col2:
        st.metric(
            label="Respostas Completas (vs esperadas)",
            value=f"{stats['completed_interventions']} / {stats['total_possible']}",
            delta=None
        )
    
    with col3:
        completion_pct = (stats['completed_interventions'] / stats['total_possible'] * 100) if stats['total_possible'] > 0 else 0
        st.metric(
            label="Progresso Global (esperado)",
            value=f"{round(completion_pct, 1)}%",
            delta=None
        )

    st.caption(
        f"Denominador esperado: {stats['n_experts_expected']} especialistas de experts.txt "
        f"(observados na API: {stats['n_experts_observed']})."
    )


def render_coverage_heatmap(data):
    """Render expert × intervention coverage heatmap."""
    st.header("🗂️ Matriz de Cobertura: Especialista × Intervenção")
    
    if data["wide"].empty:
        st.info("Sem dados para mostrar")
        return
    
    coverage = build_coverage_matrix(data["wide"], data["experts"], data["interventions"])
    
    if coverage.empty:
        st.info("Sem dados para mostrar")
        return
    
    # Style the dataframe
    def highlight_cells(val):
        if val == "✓":
            return "background-color: #e8f5ee; color: #1a6b3a; font-weight: 700"
        elif val == "—":
            return "background-color: #f3f4f6; color: #6b7280"
        return ""
    
    styled = coverage.style.applymap(highlight_cells, subset=coverage.columns[1:])
    st.dataframe(styled, use_container_width=True, height=400)
    
    # Summary stats
    total_cells = len(data["experts"]) * len(data["interventions"])
    completed_cells = (coverage.iloc[:, 1:] == "✓").sum().sum()
    st.caption(f"Células preenchidas: {completed_cells} / {total_cells} ({round(completed_cells/total_cells*100, 1)}%)")


def render_group_coverage(data):
    """Render expert × group coverage matrix."""
    st.header("📋 Cobertura por Grupo de Formulário")
    
    if data["raw"].empty:
        st.info("Sem dados para mostrar")
        return
    
    coverage = build_group_coverage(data["raw"], data["experts"], data["groups"])
    
    if coverage.empty:
        st.info("Sem dados para mostrar")
        return
    
    # Style the dataframe
    def highlight_cells(val):
        if val == "✓":
            return "background-color: #e8f5ee; color: #1a6b3a; font-weight: 700"
        elif val == "—":
            return "background-color: #f3f4f6; color: #6b7280"
        elif str(val).startswith("×"):
            return "background-color: #fef9e7; color: #7a5c00; font-weight: 700"
        return ""
    
    styled = coverage.style.applymap(highlight_cells, subset=coverage.columns[1:])
    st.dataframe(styled, use_container_width=True, height=300)


def render_response_rates(stats):
    """Render response rate charts."""
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Taxa de Resposta por Intervenção")
        if not stats["by_intervention"].empty:
            chart_data = stats["by_intervention"].sort_values("Taxa (%)", ascending=False)
            if alt is not None:
                chart = (
                    alt.Chart(chart_data)
                    .mark_bar(color="#1a5276")
                    .encode(
                        x=alt.X(
                            "Intervenção:N",
                            sort="-y",
                            axis=alt.Axis(labelAngle=-45, title="Código da Intervenção"),
                        ),
                        y=alt.Y("Taxa (%):Q", scale=alt.Scale(domain=[0, 100]), title="Taxa (%)"),
                        tooltip=["Intervenção", "Respondentes", "Total Especialistas", "Taxa (%)"],
                    )
                    .properties(height=400)
                    .configure_axis(labelColor="#f5f2ed", titleColor="#f5f2ed")
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.bar_chart(chart_data.set_index("Intervenção")["Taxa (%)"], height=400)
        else:
            st.info("Sem dados")
    
    with col2:
        st.subheader("👥 Taxa de Resposta por Especialista")
        if not stats["by_expert"].empty:
            chart_data = stats["by_expert"].sort_values("Taxa (%)", ascending=False)
            if alt is not None:
                chart = (
                    alt.Chart(chart_data)
                    .mark_bar(color="#c0392b")
                    .encode(
                        x=alt.X(
                            "Especialista:N",
                            sort="-y",
                            axis=alt.Axis(labelAngle=-45, title="Código do Especialista"),
                        ),
                        y=alt.Y("Taxa (%):Q", scale=alt.Scale(domain=[0, 100]), title="Taxa (%)"),
                        tooltip=["Especialista", "Respondidas", "Total Intervenções", "Taxa (%)"],
                    )
                    .properties(height=400)
                    .configure_axis(labelColor="#f5f2ed", titleColor="#f5f2ed")
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.bar_chart(chart_data.set_index("Especialista")["Taxa (%)"], height=400)
        else:
            st.info("Sem dados")


def _extract_submission_timestamps(raw):
    """Extract submission timestamps using known Kobo columns."""
    if raw.empty:
        return pd.Series(dtype="datetime64[ns, UTC]")

    for col in ("_submitted_at", "_submission_time", "end", "start"):
        if col in raw.columns:
            ts = pd.to_datetime(raw[col], errors="coerce", utc=True)
            ts = ts.dropna()
            if not ts.empty:
                return ts
    return pd.Series(dtype="datetime64[ns, UTC]")


def render_submission_timeline(data):
    """Render histogram of submissions over time."""
    st.header("⏱️ Ritmo de Submissões ao Longo do Tempo")

    if data["raw"].empty:
        st.info("Sem dados para mostrar")
        return

    ts = _extract_submission_timestamps(data["raw"])
    if ts.empty:
        st.info("Sem carimbos de data/hora disponíveis nas submissões")
        return

    ts_df = pd.DataFrame({"submitted_at": ts})

    if alt is not None:
        # Choose a practical temporal bin based on span.
        span_days = (ts_df["submitted_at"].max() - ts_df["submitted_at"].min()).total_seconds() / 86400
        if span_days <= 2:
            time_unit = "yearmonthdatehours"
            x_title = "Hora"
        elif span_days <= 31:
            time_unit = "yearmonthdate"
            x_title = "Dia"
        else:
            time_unit = "yearmonth"
            x_title = "Mês"

        hist = (
            alt.Chart(ts_df)
            .mark_bar(color="#b7860b")
            .encode(
                x=alt.X(f"{time_unit}(submitted_at):T", title=x_title),
                y=alt.Y("count():Q", title="Número de Submissões"),
                tooltip=[alt.Tooltip("count():Q", title="Submissões")],
            )
            .properties(height=260)
            .configure_axis(labelColor="#f5f2ed", titleColor="#f5f2ed")
        )
        st.altair_chart(hist, use_container_width=True)
    else:
        # Fallback: daily count line/bar friendly for environments without Altair.
        daily = ts_df.assign(day=ts_df["submitted_at"].dt.date).groupby("day").size().rename("Submissões")
        st.bar_chart(daily, height=260)


def render_detailed_tables(stats):
    """Render detailed response rate tables."""
    st.header("📋 Detalhes de Resposta")
    
    tab1, tab2 = st.tabs(["Por Intervenção", "Por Especialista"])
    
    with tab1:
        if not stats["by_intervention"].empty:
            st.dataframe(
                stats["by_intervention"].sort_values("Taxa (%)", ascending=False),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Sem dados")
    
    with tab2:
        if not stats["by_expert"].empty:
            st.dataframe(
                stats["by_expert"].sort_values("Taxa (%)", ascending=False),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Sem dados")


# ═══════════════════════════════════════════════════════════════
# Main App
# ═══════════════════════════════════════════════════════════════

def main():
    # Check if auto-refresh is due (before rendering)
    if st.session_state.auto_refresh_enabled:
        current_time = time.time()
        last_refresh = st.session_state.last_auto_refresh_time
        if last_refresh == 0:
            # First time enabled, set initial timestamp
            st.session_state.last_auto_refresh_time = current_time
        elif current_time - last_refresh >= 300:  # 300 seconds = 5 minutes
            st.session_state.last_auto_refresh_time = current_time
            st.cache_data.clear()
            st.rerun()
    
    # Apply typography and palette from the report
    apply_report_theme()
    
    # If auto-refresh is enabled, inject client-side JavaScript for polling
    if st.session_state.auto_refresh_enabled:
        st.markdown("""
        <script>
        // Auto-reload page every 5 minutes (300000 ms) when auto-refresh is enabled
        setTimeout(function() {
            location.reload();
        }, 300000);
        </script>
        """, unsafe_allow_html=True)

    # Render header
    render_header()
    
    # Create a placeholder for the data fetch to happen after sidebar is rendered
    # We need to capture sidebar input first
    with st.sidebar:
        st.header("⚙️ Controles")

        st.subheader("🔐 Kobo API")
        use_manual_token = st.toggle(
            "Inserir token manualmente (sessão actual)",
            value=False,
            help="No Streamlit Cloud, prefira guardar KOBO_TOKEN em Secrets."
        )
        manual_token = ""
        manual_server = ""
        if use_manual_token:
            manual_token = st.text_input(
                "KOBO_TOKEN",
                value="",
                type="password",
                help="Token usado apenas nesta sessão do dashboard."
            )
            manual_server = st.text_input(
                "KOBO_SERVER (opcional)",
                value=DEFAULT_KOBO_SERVER,
                help="Ex: https://eu.kobotoolbox.org"
            )

        resolved_server, resolved_token = resolve_kobo_credentials(
            token_override=manual_token,
            server_override=manual_server,
        )
        resolved_assets = resolve_kobo_asset_entries()

        token_source = "manual" if manual_token.strip() else (
            "secrets/env/config" if resolved_token else "não configurado"
        )
        st.caption(f"Fonte do token: {token_source}")
        st.caption(f"Servidor Kobo: {resolved_server}")
        st.caption(f"Assets configurados (SUBFORM_ASSET_*): {len(resolved_assets)}")
        st.divider()
        
        # Manual refresh button
        if st.button("🔄 Actualizar Agora", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        
        # Auto-refresh toggle (non-blocking with session state)
        st.session_state.auto_refresh_enabled = st.toggle(
            "Auto-actualizar (5 min)", 
            value=st.session_state.auto_refresh_enabled
        )
        
        # Show status when auto-refresh is active
        if st.session_state.auto_refresh_enabled:
            st.caption("✅ Auto-refresh ativo — a página actualiza a cada 5 minutos")
        
        st.divider()
        
        # Exclusion input
        st.subheader("⛔ Filtrar Especialistas")
        exclusion_input = st.text_input(
            "Códigos a excluir (separados por vírgula):",
            value="",
            placeholder="Ex: 001PM, 001XX",
            help="Insira os códigos de especialistas para remover da análise"
        )
        excluded_experts_from_ui = _split_codes(exclusion_input)
    
    # Fetch data with loading indicator using the exclusions from UI
    with st.spinner("A carregar dados da API do KoboToolbox..."):
        data = fetch_and_process_data(
            excluded_experts=frozenset(excluded_experts_from_ui) if excluded_experts_from_ui else None,
            kobo_server=resolved_server,
            kobo_token=resolved_token,
            asset_entries=tuple(sorted(resolved_assets.items())),
        )
    
    if not data:
        st.error("Falha ao carregar dados. Verifique a configuração do KOBO_TOKEN.")
        st.stop()
    
    # Compute statistics
    stats = compute_stats(data)
    
    # Render rest of sidebar with stats
    with st.sidebar:
        if excluded_experts_from_ui:
            st.info(
                f"🚫 **{len(excluded_experts_from_ui)} especialista(s) excluído(s)**\n\n"
                f"Códigos: {', '.join(sorted(excluded_experts_from_ui))}"
            )
            if data:
                n_removed = data.get("n_submissions_before_exclusion", 0) - data.get("n_submissions", 0)
                st.caption(f"Submissões removidas: {n_removed}")
        
        st.divider()
        
        # Summary stats
        st.header("📈 Resumo")
        if data and not data["wide"].empty:
            st.metric("Especialistas (esperados)", stats["n_experts_expected"])
            st.metric("Especialistas (observados)", stats["n_experts_observed"])
            st.metric("Submissões", data["n_submissions"])
            st.metric("Taxa de Resposta", f"{stats['response_rate']}%")
            st.metric("Intervenções", len(data["interventions"]))
            st.metric("Grupos", len(data["groups"]))
            
            st.divider()
            st.caption(f"Última actualização: {data['timestamp'].strftime('%d/%m/%Y %H:%M:%S')}")
        else:
            st.warning("Sem dados disponíveis")
    
    # Main content
    if data["wide"].empty:
        st.warning("⚠️ Nenhuma submissão encontrada ainda.")
        st.info("O painel será actualizado automaticamente quando houver dados disponíveis.")
        st.stop()
    
    # Overview cards
    render_overview_cards(stats)
    st.divider()
    
    # Coverage heatmap
    render_coverage_heatmap(data)
    st.divider()
    
    # Group coverage
    render_group_coverage(data)
    st.divider()
    
    # Response rate charts
    render_response_rates(stats)
    st.divider()

    # Submission timeline
    render_submission_timeline(data)
    st.divider()
    
    # Detailed tables
    render_detailed_tables(stats)


if __name__ == "__main__":
    main()
