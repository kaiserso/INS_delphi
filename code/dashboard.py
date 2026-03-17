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
        load_config, fetch_all,
    )
except ImportError as e:
    st.error(f"Failed to import from aggregate_results.py: {e}")
    st.stop()

# Load config
cfg = load_config()
deployed_cfg = load_config("deployed_forms.env")


def _init_secret_get(key, default=""):
    """Read Streamlit secret safely during module initialization."""
    try:
        return str(st.secrets.get(key, default)).strip()
    except Exception:
        return default


TOPIC = (
    _init_secret_get("TOPIC_CODE")
    or os.getenv("TOPIC_CODE", "").strip()
    or cfg.get("TOPIC_CODE", "")
    or deployed_cfg.get("TOPIC_CODE", "")
    or "malaria"
)
DEFAULT_KOBO_SERVER = cfg.get("KOBO_SERVER", "https://eu.kobotoolbox.org")
DEFAULT_KOBO_TOKEN = cfg.get("KOBO_TOKEN", "")

# Initialize session state for auto-refresh
if "auto_refresh_enabled" not in st.session_state:
    st.session_state.auto_refresh_enabled = False
if "last_auto_refresh_time" not in st.session_state:
    st.session_state.last_auto_refresh_time = 0


def _secrets_get(key, default=""):
    """Safe wrapper around st.secrets — returns default when no secrets file exists."""
    try:
        return str(st.secrets.get(key, default)).strip()
    except Exception:
        return default


def _secrets_keys():
    """Safe wrapper around st.secrets.keys() — returns empty list when unavailable."""
    try:
        return list(st.secrets.keys())
    except Exception:
        return []


def resolve_kobo_credentials(token_override="", server_override=""):
    """Resolve Kobo credentials with priority: UI override > secrets > env > config.env."""
    server = (
        (server_override or "").strip()
        or _secrets_get("KOBO_SERVER")
        or os.getenv("KOBO_SERVER", "").strip()
        or DEFAULT_KOBO_SERVER
    )
    token = (
        (token_override or "").strip()
        or _secrets_get("KOBO_TOKEN")
        or os.getenv("KOBO_TOKEN", "").strip()
        or DEFAULT_KOBO_TOKEN
    )
    return server, token


def resolve_kobo_asset_entries():
    """Collect SUBFORM_ASSET_* entries from deployed_forms.env, config.env, env vars, and secrets.

    Priority (highest last, so later sources overwrite earlier ones):
      deployed_forms.env < config.env < environment variables < st.secrets
    """
    assets = {}

    # deployed_forms.env — written by deploy_kobo_forms.py; primary local source
    for key, value in deployed_cfg.items():
        if key.upper().startswith("SUBFORM_ASSET_") and str(value).strip():
            assets[key] = str(value).strip()

    for key, value in cfg.items():
        if key.upper().startswith("SUBFORM_ASSET_") and str(value).strip():
            assets[key] = str(value).strip()

    for key, value in os.environ.items():
        if key.upper().startswith("SUBFORM_ASSET_") and str(value).strip():
            assets[key] = str(value).strip()

    for key in _secrets_keys():
        if key.upper().startswith("SUBFORM_ASSET_"):
            value = _secrets_get(key)
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


@st.cache_data(show_spinner=False)
def load_team_mapping(topic=TOPIC):
    """Load intervention → team mapping from the dictionary Excel file."""
    dict_path = Path(__file__).resolve().parents[1] / "dict" / f"dicionario_delphi_w1_{topic}.xlsx"
    if not dict_path.exists():
        return {}
    try:
        xl = pd.ExcelFile(dict_path)
        # Find the Catalogo sheet (named e.g. "Catalogo_HIV_SIDA")
        sheet = next((s for s in xl.sheet_names if s.lower().startswith("catalogo")), xl.sheet_names[0])
        df = pd.read_excel(dict_path, sheet_name=sheet, header=1)
        if "Código" not in df.columns or "Team" not in df.columns:
            return {}
        mapping = (
            df[["Código", "Team"]]
            .dropna(subset=["Código"])
            .set_index("Código")["Team"]
            .astype(str)
            .str.strip()
            .to_dict()
        )
        return mapping
    except Exception:
        return {}


def _slugify(text):
    """Replicate generate_kobo_and_pages.py slugify for label lookups."""
    import unicodedata
    text = unicodedata.normalize("NFKD", str(text))
    text = text.encode("ascii", "ignore").decode("ascii")
    text = text.lower().strip()
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", "_", text)
    return text


@st.cache_data(show_spinner=False)
def load_group_label_mapping(topic=TOPIC):
    """Return {_group_slug: display_label} using the Programa (or Grupo) column from the dictionary."""
    dict_path = Path(__file__).resolve().parents[1] / "dict" / f"dicionario_delphi_w1_{topic}.xlsx"
    if not dict_path.exists():
        return {}
    try:
        xl = pd.ExcelFile(dict_path)
        sheet = next((s for s in xl.sheet_names if s.lower().startswith("catalogo")), xl.sheet_names[0])
        df = pd.read_excel(dict_path, sheet_name=sheet, header=1)

        group_by = cfg.get("SUBFORM_GROUP_BY", "programa").lower().strip()
        col = {"grupo": "Grupo", "programa": "Programa", "team": "Team"}.get(group_by, "Programa")
        if col not in df.columns:
            col = next((c for c in ("Programa", "Grupo") if c in df.columns), None)
        if col is None:
            return {}

        # Build {base_slug: display_label}
        base_map = {_slugify(v): str(v) for v in df[col].dropna().unique()}
        return base_map
    except Exception:
        return {}


def format_group_label(group_slug, label_map):
    """Convert a _group slug (e.g. 'ct_adulto_2') to a human display label (e.g. 'C&T ADULTO (2)')."""
    if not label_map:
        return group_slug
    if group_slug in label_map:
        return label_map[group_slug]
    # Strip trailing _N chunk suffix
    m = re.match(r"^(.+)_(\d+)$", group_slug)
    if m:
        base, part = m.group(1), m.group(2)
        if base in label_map:
            return f"{label_map[base]} ({part})"
    # Fallback: humanise the slug
    return group_slug.replace("_", " ").title()


@st.cache_data(show_spinner=False)
def load_group_team_mapping(topic=TOPIC):
    """Return {group_slug: team} by matching dictionary Grupo names to _group slugs."""
    dict_path = Path(__file__).resolve().parents[1] / "dict" / f"dicionario_delphi_w1_{topic}.xlsx"
    if not dict_path.exists():
        return {}
    try:
        xl = pd.ExcelFile(dict_path)
        sheet = next((s for s in xl.sheet_names if s.lower().startswith("catalogo")), xl.sheet_names[0])
        df = pd.read_excel(dict_path, sheet_name=sheet, header=1)
        if "Grupo" not in df.columns or "Team" not in df.columns:
            return {}
        # Dominant team per Grupo name
        grupo_team = (
            df[["Grupo", "Team"]].dropna()
            .groupby("Grupo")["Team"]
            .agg(lambda x: x.mode().iloc[0])
            .to_dict()
        )
        # Normalise both keys for matching: strip non-alphanumeric, lowercase
        def _norm(s):
            return re.sub(r"[^a-z0-9]", "", str(s).lower())
        return {_norm(g): team for g, team in grupo_team.items()}
    except Exception:
        return {}


def _group_to_team(group_slug, norm_map):
    """Map a _group slug to a team letter using the normalised dictionary map."""
    key = re.sub(r"[^a-z0-9]", "", str(group_slug).lower())
    return norm_map.get(key)


EXPECTED_EXPERTS = load_expected_experts()
N_EXPECTED_EXPERTS = len(EXPECTED_EXPERTS)
# When experts.txt is not present (e.g. Streamlit Cloud deployment),
# fall back to the count stored in Streamlit secrets as N_EXPERTS_EXPECTED.
_N_EXPERTS_SOURCE = "experts.txt"
if N_EXPECTED_EXPERTS == 0:
    _secret_n = _secrets_get("N_EXPERTS_EXPECTED", "")
    if _secret_n.isdigit() and int(_secret_n) > 0:
        N_EXPECTED_EXPERTS = int(_secret_n)
        _N_EXPERTS_SOURCE = "secrets (N_EXPERTS_EXPECTED)"

# Per-team expected expert counts (from secrets: N_EXPERTS_TEAM_A … _D)
N_EXPERTS_PER_TEAM: dict = {}
for _team in ("A", "B", "C", "D"):
    _val = _secrets_get(f"N_EXPERTS_TEAM_{_team}", "")
    if _val.isdigit() and int(_val) > 0:
        N_EXPERTS_PER_TEAM[_team] = int(_val)

# Page configuration
st.set_page_config(
    page_title=f"Delphi W1 — {TOPIC.title()} | Painel de Monitoramento",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


def apply_report_theme():
    """Apply light, professional styling matching the HTML report design."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');

        :root {
            --ink: #1a1a1a;
            --paper: #f5f7fa;
            --surface: #ffffff;
            --accent: #c0392b;
            --accent2: #1a5276;
            --gold: #b7860b;
            --muted: #78909C;
            --border: #ECEFF1;
            --sim: #2e7d52;
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
            border-radius: 8px;
            padding: 14px 16px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
        }

        div[data-testid="stMetric"] label {
            color: var(--muted);
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 600;
        }

        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            color: var(--accent2);
            font-family: 'DM Serif Display', serif;
            font-size: 1.8rem;
            font-weight: 600;
        }

        .stButton > button {
            background: var(--accent2);
            color: #ffffff;
            border: none;
            border-radius: 6px;
            font-weight: 500;
            box-shadow: 0 1px 3px rgba(26, 82, 118, 0.15);
        }

        .stButton > button:hover {
            background: #154470;
            box-shadow: 0 2px 6px rgba(26, 82, 118, 0.25);
            color: #ffffff;
        }

        .stButton > button,
        .stButton > button * {
            color: #ffffff !important;
        }

        [data-testid="stDataFrame"] {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 8px;
        }

        /* Divider styling */
        .element-container .stHorizontalBlock hr {
            border-color: var(--border);
            margin: 20px 0;
        }

        /* Toggle and checkbox styling */
        [data-testid="stCheckbox"] {
            color: var(--ink);
        }

        [data-testid="stCheckbox"] input {
            accent-color: var(--accent2);
        }

        /* Text input styling */
        input, textarea {
            background: var(--surface) !important;
            color: var(--ink) !important;
            border: 1px solid var(--border) !important;
            border-radius: 6px !important;
        }

        input::placeholder, textarea::placeholder {
            color: var(--muted) !important;
        }

        /* Tab styling */
        [data-testid="stTabs"] button {
            color: var(--muted);
            font-weight: 500;
        }

        [data-testid="stTabs"] button[aria-selected="true"] {
            color: var(--accent2);
            border-bottom: 2px solid var(--accent2);
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
    Fetch submissions from Kobo API at questionnaire (form group) level.
    Only needs expert_code, _group, and _submitted_at — no build_wide.
    Cached to avoid redundant API calls.
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

    # Fetch data — keep only the columns we need for questionnaire-level tracking
    raw = fetch_all(asset_filter=None)
    raw = raw.fillna("")

    if raw.empty:
        return {
            "timestamp": datetime.now(),
            "raw": raw,
            "experts": [],
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

    experts = sorted(raw["expert_code"].dropna().unique())
    groups = sorted(raw["_group"].dropna().unique())
    label_map = load_group_label_mapping()
    group_labels = {g: format_group_label(g, label_map) for g in groups}

    return {
        "timestamp": datetime.now(),
        "raw": raw,
        "experts": experts,
        "groups": groups,
        "group_labels": group_labels,
        "n_submissions": len(raw),
        "n_experts": len(experts),
        "n_submissions_before_exclusion": n_before,
    }


def build_group_coverage(raw, experts, groups, group_labels=None):
    """Build expert × group submission matrix using display labels as column names."""
    if raw.empty or not experts or not groups:
        return pd.DataFrame()

    counts = raw.groupby(["expert_code", "_group"]).size().to_dict()
    matrix = []
    for expert in experts:
        row = {"Expert": expert}
        for group in groups:
            label = (group_labels or {}).get(group, group)
            n = counts.get((expert, group), 0)
            row[label] = "✓" if n == 1 else (f"×{n}" if n > 1 else "—")
        matrix.append(row)
    return pd.DataFrame(matrix)


def compute_stats(data):
    """Compute questionnaire-level summary statistics."""
    raw = data.get("raw", pd.DataFrame())
    n_experts_observed = data["n_experts"]
    n_experts_expected = N_EXPECTED_EXPERTS if N_EXPECTED_EXPERTS > 0 else n_experts_observed
    groups = data["groups"]
    n_groups = len(groups)

    if raw.empty or not groups:
        return {
            "response_rate": 0,
            "completed_submissions": 0,
            "total_possible": 0,
            "n_experts_expected": n_experts_expected,
            "n_experts_observed": n_experts_observed,
            "by_group": pd.DataFrame(),
            "by_expert": pd.DataFrame(),
        }

    submitted = set(zip(
        raw["expert_code"].astype(str).str.strip(),
        raw["_group"].astype(str).str.strip(),
    ))

    total_possible = n_experts_expected * n_groups
    completed = sum(1 for (e, g) in submitted if e and g)
    response_rate = (completed / total_possible * 100) if total_possible > 0 else 0

    # By questionnaire group
    by_group = []
    group_labels = data.get("group_labels", {})
    group_counts = raw.groupby("_group")["expert_code"].nunique().to_dict()
    for group in groups:
        answered = group_counts.get(group, 0)
        pct = (answered / n_experts_expected * 100) if n_experts_expected > 0 else 0
        by_group.append({
            "Questionário": group_labels.get(group, group),
            "Submetido por": answered,
            "Total Especialistas": n_experts_expected,
            "Taxa (%)": round(pct, 1),
        })

    # By expert
    expert_counts = raw.groupby("expert_code")["_group"].nunique().to_dict()
    by_exp = []
    for expert in data["experts"]:
        answered = expert_counts.get(expert, 0)
        pct = (answered / n_groups * 100) if n_groups > 0 else 0
        by_exp.append({
            "Especialista": expert,
            "Questionários submetidos": answered,
            "Total Questionários": n_groups,
            "Taxa (%)": round(pct, 1),
        })

    return {
        "response_rate": round(response_rate, 1),
        "completed_submissions": completed,
        "total_possible": total_possible,
        "n_experts_expected": n_experts_expected,
        "n_experts_observed": n_experts_observed,
        "by_group": pd.DataFrame(by_group),
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


def render_overview_cards(stats, data):
    """Render top-level overview cards including per-team completion rates."""
    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(label="Taxa Global de Submissão", value=f"{stats['response_rate']}%")

    with col2:
        st.metric(
            label="Submissões Recebidas (vs esperadas)",
            value=f"{stats['completed_submissions']} / {stats['total_possible']}",
        )

    with col3:
        st.metric(label="Especialistas Observados", value=stats['n_experts_observed'])

    st.caption(
        f"Denominador esperado: {stats['n_experts_expected']} especialistas de {_N_EXPERTS_SOURCE} "
        f"(observados na API: {stats['n_experts_observed']})."
    )

    # Per-team completion rates (only shown when N_EXPERTS_PER_TEAM secrets are configured)
    if N_EXPERTS_PER_TEAM and not data["raw"].empty:
        norm_map = load_group_team_mapping()
        if norm_map:
            # Count unique submitting experts per group
            group_experts = (
                data["raw"]
                .groupby("_group")["expert_code"]
                .nunique()
                .to_dict()
            )
            # Aggregate by team
            team_submitted: dict = {}
            team_n_groups: dict = {}
            for group in data["groups"]:
                team = _group_to_team(group, norm_map)
                if team:
                    team_submitted[team] = team_submitted.get(team, 0) + group_experts.get(group, 0)
                    team_n_groups[team] = team_n_groups.get(team, 0) + 1

            team_colors = {"A": "#1a5c8a", "B": "#2e7d52", "C": "#b45309", "D": "#6b3fa0"}
            teams_to_show = sorted(set(N_EXPERTS_PER_TEAM) | set(team_submitted))
            cols = st.columns(len(teams_to_show))
            for col, team in zip(cols, teams_to_show):
                n_exp = N_EXPERTS_PER_TEAM.get(team, stats['n_experts_expected'])
                n_grp = team_n_groups.get(team, 0)
                submitted = team_submitted.get(team, 0)
                total = n_exp * n_grp
                pct = round(submitted / total * 100, 1) if total > 0 else 0
                color = team_colors.get(team, "#78909C")
                col.markdown(
                    f"<div style='background:#ffffff;border:1px solid #ECEFF1;border-radius:8px;"
                    f"padding:12px 16px;text-align:center;border-top:3px solid {color}'>"
                    f"<div style='font-size:0.7rem;text-transform:uppercase;letter-spacing:0.08em;"
                    f"color:#78909C;font-weight:600'>Equipa {team}</div>"
                    f"<div style='font-size:1.6rem;font-weight:700;color:{color}'>{pct}%</div>"
                    f"<div style='font-size:0.75rem;color:#78909C'>{submitted}/{total} submissões</div>"
                    f"</div>",
                    unsafe_allow_html=True,
                )


def render_coverage_heatmap(data):
    """Render expert × group submission matrix."""
    st.header("🗂️ Matriz de Cobertura: Especialista × Questionário")

    if data["raw"].empty:
        st.info("Sem dados para mostrar")
        return

    coverage = build_group_coverage(data["raw"], data["experts"], data["groups"], data.get("group_labels"))

    if coverage.empty:
        st.info("Sem dados para mostrar")
        return

    def highlight_cells(val):
        if val == "✓":
            return "background-color: #e8f5ee; color: #2e7d52; font-weight: 700"
        elif val == "—":
            return "background-color: #f5f5f5; color: #90A4AE"
        elif str(val).startswith("×"):
            return "background-color: #fef9e7; color: #7a5c00; font-weight: 700"
        return ""

    styled = coverage.style.applymap(highlight_cells, subset=coverage.columns[1:])
    st.dataframe(styled, use_container_width=True, height=400)

    total_cells = len(data["experts"]) * len(data["groups"])
    completed_cells = (coverage.iloc[:, 1:] == "✓").sum().sum()
    st.caption(f"Submissões recebidas: {completed_cells} / {total_cells} ({round(completed_cells/total_cells*100, 1)}%)")


def _make_chart(chart_data, x_field, x_title, color, tooltip_fields, height=380, bottom_pad=80):
    """Build a standard bar chart with explicit light background so labels show on Streamlit Cloud."""
    return (
        alt.Chart(chart_data)
        .mark_bar(color=color)
        .encode(
            x=alt.X(
                x_field,
                sort="-y",
                axis=alt.Axis(
                    labelAngle=-45,
                    title=x_title,
                    labelColor="#1a1a1a",
                    titleColor="#1a1a1a",
                    labelOverlap=False,
                    labelLimit=0,
                ),
            ),
            y=alt.Y(
                "Taxa (%):Q",
                scale=alt.Scale(domain=[0, 100]),
                title="Taxa (%)",
                axis=alt.Axis(labelColor="#1a1a1a", titleColor="#1a1a1a"),
            ),
            tooltip=tooltip_fields,
        )
        .properties(height=height, padding={"left": 5, "right": 5, "top": 5, "bottom": bottom_pad})
        .configure_view(fill="#f5f7fa", stroke=None)
        .configure(background="#f5f7fa")
    )


def render_response_rates(stats):
    """Render questionnaire-level response rate charts."""
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📊 Taxa de Submissão por Questionário")
        if not stats["by_group"].empty:
            chart_data = stats["by_group"].sort_values("Taxa (%)", ascending=False)
            if alt is not None:
                chart = _make_chart(
                    chart_data,
                    x_field="Questionário:N",
                    x_title="Questionário",
                    color="#1a5276",
                    tooltip_fields=["Questionário", "Submetido por", "Total Especialistas", "Taxa (%)"],
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.bar_chart(chart_data.set_index("Questionário")["Taxa (%)"], height=380)
        else:
            st.info("Sem dados")

    with col2:
        st.subheader("👥 Taxa de Submissão por Especialista")
        if not stats["by_expert"].empty:
            chart_data = stats["by_expert"].sort_values("Taxa (%)", ascending=False)
            if alt is not None:
                chart = _make_chart(
                    chart_data,
                    x_field="Especialista:N",
                    x_title="Código do Especialista",
                    color="#c0392b",
                    tooltip_fields=["Especialista", "Questionários submetidos", "Total Questionários", "Taxa (%)"],
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.bar_chart(chart_data.set_index("Especialista")["Taxa (%)"], height=380)
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
                x=alt.X(f"{time_unit}(submitted_at):T", title=x_title, axis=alt.Axis(labelColor="#1a1a1a", titleColor="#1a1a1a", labelOverlap=False)),
                y=alt.Y("count():Q", title="Número de Submissões", axis=alt.Axis(labelColor="#1a1a1a", titleColor="#1a1a1a")),
                tooltip=[alt.Tooltip("count():Q", title="Submissões")],
            )
            .properties(height=260, padding={"left": 5, "right": 5, "top": 5, "bottom": 40})
            .configure_view(fill="#f5f7fa", stroke=None)
            .configure(background="#f5f7fa")
        )
        st.altair_chart(hist, use_container_width=True)
    else:
        # Fallback: daily count line/bar friendly for environments without Altair.
        daily = ts_df.assign(day=ts_df["submitted_at"].dt.date).groupby("day").size().rename("Submissões")
        st.bar_chart(daily, height=260)


def render_detailed_tables(stats):
    """Render detailed response rate tables."""
    st.header("📋 Detalhes de Submissão")

    tab1, tab2 = st.tabs(["Por Questionário", "Por Especialista"])

    with tab1:
        if not stats["by_group"].empty:
            st.dataframe(
                stats["by_group"].sort_values("Taxa (%)", ascending=False),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Sem dados")

    with tab2:
        if not stats["by_expert"].empty:
            st.dataframe(
                stats["by_expert"].sort_values("Taxa (%)", ascending=False),
                use_container_width=True,
                hide_index=True,
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
        if data and data["n_submissions"] > 0:
            st.metric("Especialistas (esperados)", stats["n_experts_expected"])
            st.metric("Especialistas (observados)", stats["n_experts_observed"])
            st.metric("Submissões recebidas", data["n_submissions"])
            st.metric("Taxa de Submissão", f"{stats['response_rate']}%")
            st.metric("Questionários", len(data["groups"]))

            st.divider()
            st.caption(f"Última actualização: {data['timestamp'].strftime('%d/%m/%Y %H:%M:%S')}")
        else:
            st.warning("Sem dados disponíveis")

    # Main content
    if data["n_submissions"] == 0:
        st.warning("⚠️ Nenhuma submissão encontrada ainda.")
        st.info("O painel será actualizado automaticamente quando houver dados disponíveis.")
        st.stop()

    # Overview cards (includes team completion)
    render_overview_cards(stats, data)
    st.divider()

    # Coverage heatmap (expert × questionnaire)
    render_coverage_heatmap(data)
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
