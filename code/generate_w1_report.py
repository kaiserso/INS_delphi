#!/usr/bin/env python3
"""
Gerador de Relatório Delphi W1
================================
Uso:
    python gerar_relatorio_delphi_w1.py <resultados.xlsx> [dicionario.xlsx] [--output-dir DIR]

Argumentos:
    resultados.xlsx   Ficheiro com os resultados agregados (obrigatório).
                      Folha usada: 'Responses' (tall-skinny, uma linha por
                      expert_code × intervention).
    dicionario.xlsx   Ficheiro dicionário com os metadados das intervenções
                      (opcional mas recomendado).
                      Folha usada: primeira cujo nome começa por 'Catalogo'.
    --output-dir DIR  Directório de saída (opcional, padrão: ../reports/).

Metadados das intervenções (código, nome, componente, URL):
    1. Dicionário externo (se fornecido) — folha Catalogo_*
    2. Folha Catalogo_* dentro do próprio ficheiro de resultados
    3. Colunas url_* na folha Submissions do ficheiro de resultados
       (nome e componente ficam em branco neste último caso)
    Se nenhuma fonte estiver disponível, o script usa os códigos brutos.

Colunas obrigatórias na folha Responses:
    expert_code, intervention, gate
Colunas opcionais:
    dup, which_dup, which_dup_other, intg, which_intg, which_intg_other,
    res, oth, oth_reason, impact, cmt, exp, modality, group

Saída:
    delphi_w1_relatorio_<timestamp>.html  no directório de trabalho actual
"""

import sys
import os
from collections import defaultdict, OrderedDict
from datetime import datetime
import statistics
from pathlib import Path

try:
    import pandas as pd
    import openpyxl
except ImportError:
    sys.exit("Erro: dependências em falta. Execute: pip install pandas openpyxl")

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACÇÃO DE METADADOS
# ─────────────────────────────────────────────────────────────────────────────

def _find_catalogo_sheet(wb):
    """Return the first sheet whose name starts with 'Catalogo' (case-insensitive)."""
    for name in wb.sheetnames:
        if name.lower().startswith("catalogo"):
            return wb[name]
    return None

def _parse_catalogo_sheet(ws):
    """
    Parse a Catalogo_* sheet.
    Expects row 1 as section headers (ignored), row 2 as column names,
    data from row 3 onward.
    Returns list of dicts with keys: code, label, component, url
    """
    # Find header row: look for a row containing 'Código'
    header_row = None
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if any(str(c).strip() == "Código" for c in row if c):
            header_row = i
            headers = [str(c).strip() if c else "" for c in row]
            break
    if header_row is None:
        return []

    col = {h: i for i, h in enumerate(headers)}
    interventions = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        code = row[col.get("Código", -1)] if col.get("Código") is not None else None
        if not code or not str(code).strip():
            continue
        code = str(code).strip()
        label     = row[col["Intervenção"]]     if "Intervenção"  in col else None
        component = row[col["Componente"]]      if "Componente"   in col else None
        url       = row[col["URL da Ficha"]]    if "URL da Ficha" in col else None
        interventions.append({
            "code":      code,
            "label":     str(label).strip()     if label     else code,
            "component": str(component).strip() if component else "",
            "url":       str(url).strip()        if url       else "",
        })
    return interventions

def _parse_submissions_urls(wb):
    """
    Fallback: extract URLs from url_* columns in the Submissions sheet.
    Returns list of dicts with code and url only (label/component empty).
    """
    if "Submissions" not in wb.sheetnames:
        return []
    ws = wb["Submissions"]
    headers = [c.value for c in ws[1]]
    url_cols = [(i, h) for i, h in enumerate(headers)
                if h and str(h).lower().startswith("url_")]
    seen = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        for i, h in url_cols:
            code = h[4:]  # strip "url_"
            if code not in seen and row[i]:
                seen[code] = str(row[i]).strip()
    return [{"code": c, "label": c, "component": "", "url": u}
            for c, u in sorted(seen.items())]

def _parse_responses_codes(df):
    """Last-resort fallback: just use the codes found in the data."""
    codes = sorted(df["intervention"].dropna().unique())
    return [{"code": c, "label": c, "component": "", "url": ""} for c in codes]

def load_metadata(results_path, dict_path=None):
    """
    Load intervention metadata with the following priority:
      1. External dictionary file (dict_path), Catalogo_* sheet
      2. Catalogo_* sheet inside results file
      3. url_* columns in Submissions sheet of results file
      4. Raw codes from Responses sheet
    Returns ordered list of dicts: code, label, component, url
    """
    interventions = []

    # 1. External dictionary
    if dict_path and os.path.exists(dict_path):
        wb = openpyxl.load_workbook(dict_path, data_only=True)
        ws = _find_catalogo_sheet(wb)
        if ws:
            interventions = _parse_catalogo_sheet(ws)
            if interventions:
                print(f"  Metadados: dicionário externo '{os.path.basename(dict_path)}' "
                      f"({len(interventions)} intervenções)")
                return interventions

    # 2. Catalogo_* in results file
    wb_res = openpyxl.load_workbook(results_path, data_only=True)
    ws = _find_catalogo_sheet(wb_res)
    if ws:
        interventions = _parse_catalogo_sheet(ws)
        if interventions:
            print(f"  Metadados: folha Catalogo_* no ficheiro de resultados "
                  f"({len(interventions)} intervenções)")
            return interventions

    # 3. url_* columns in Submissions
    interventions = _parse_submissions_urls(wb_res)
    if interventions:
        print(f"  Metadados: colunas url_* na folha Submissions "
              f"({len(interventions)} intervenções, sem nomes)")
        return interventions

    # 4. Raw codes
    return None  # signal to caller to build from df after loading

# ─────────────────────────────────────────────────────────────────────────────
# NORMALIZAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

GATE_NORM = {
    "sim_def": "sim_def", "sim definitivamente": "sim_def",
    "sim,definitivamente": "sim_def", "sim": "sim_def",
    "possivelmente": "possivelmente",
    "nao": "nao", "não": "nao", "no": "nao",
}
YN_NORM = {"sim": "sim", "yes": "sim", "nao": "nao", "não": "nao", "no": "nao"}

def norm_gate(v):
    if pd.isna(v): return None
    return GATE_NORM.get(str(v).strip().lower(), None)

def norm_yn(v):
    if pd.isna(v): return None
    return YN_NORM.get(str(v).strip().lower(), None)

def parse_multi(v):
    if pd.isna(v) or str(v).strip() == "": return []
    return [x.strip() for x in str(v).replace(",", " ").split() if x.strip()]

def safe_int(v):
    try: return int(float(str(v)))
    except: return None

# ─────────────────────────────────────────────────────────────────────────────
# CARREGAMENTO DE DADOS
# ─────────────────────────────────────────────────────────────────────────────

def load_data(path, sheet="Responses"):
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    missing = {"expert_code", "intervention", "gate"} - set(df.columns)
    if missing:
        sys.exit(f"Erro: colunas obrigatórias em falta: {missing}")
    return df


def load_expected_experts(experts_file=None):
    """Load unique expected experts from experts.txt, ignoring comments/blank lines."""
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

# ─────────────────────────────────────────────────────────────────────────────
# AGREGAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

def aggregate(df, interventions):
    """Aggregate responses per intervention. Returns dict keyed by code."""
    inv_codes  = [r["code"] for r in interventions]
    inv_labels = {r["code"]: r["label"] for r in interventions}
    results = {}

    for inv in interventions:
        code = inv["code"]
        sub  = df[df["intervention"] == code].copy()

        # Per expert: prefer row with gate filled
        deduped = []
        for _, grp in sub.groupby("expert_code", sort=False):
            filled = grp[grp["gate"].notna() & (grp["gate"].str.strip() != "")]
            deduped.append(filled.iloc[0] if len(filled) > 0 else grp.iloc[0])

        n_sim = n_poss = n_nao = n_missing = 0
        impacts = []
        dup_yes = intg_yes = res_yes = 0
        which_dup_counts  = defaultdict(int)
        which_intg_counts = defaultdict(int)
        comments = []

        for row in deduped:
            g = norm_gate(row.get("gate", None))
            if   g == "sim_def":       n_sim  += 1
            elif g == "possivelmente": n_poss += 1
            elif g == "nao":           n_nao  += 1
            else:                      n_missing += 1

            if g in ("sim_def", "possivelmente"):
                imp = safe_int(row.get("impact", None))
                if imp and 1 <= imp <= 3:
                    impacts.append(imp)

                if norm_yn(row.get("dup", None)) == "sim":
                    dup_yes += 1
                    for t in parse_multi(row.get("which_dup", None)):
                        if t in inv_codes and t != code:
                            which_dup_counts[t] += 1

                if norm_yn(row.get("intg", None)) == "sim":
                    intg_yes += 1
                    for t in parse_multi(row.get("which_intg", None)):
                        if t in inv_codes and t != code:
                            which_intg_counts[t] += 1

                if norm_yn(row.get("res", None)) == "sim":
                    res_yes += 1

                cmt = row.get("cmt", None)
                if cmt and not pd.isna(cmt) and str(cmt).strip() not in ("", ".", "None"):
                    comments.append(str(cmt).strip())

        n_resp = n_sim + n_poss + n_nao

        top_dup  = sorted(which_dup_counts,  key=which_dup_counts.get,  reverse=True)[:3]
        top_intg = sorted(which_intg_counts, key=which_intg_counts.get, reverse=True)[:3]

        avg_impact = round(sum(impacts) / len(impacts), 2) if impacts else 0

        results[code] = {
            "code":      code,
            "label":     inv["label"],
            "url":       inv["url"],
            "component": inv["component"],
            "n_total":   n_resp,
            "n_missing": n_missing,
            "n_sim":     n_sim,
            "n_poss":    n_poss,
            "n_nao":     n_nao,
            "pct_optimizable": round((n_sim + n_poss) / n_resp * 100) if n_resp else 0,
            "pct_definitely":  round(n_sim / n_resp * 100)            if n_resp else 0,
            "avg_impact": avg_impact,
            "dup_pct":   round(dup_yes  / n_resp * 100) if n_resp else 0,
            "intg_pct":  round(intg_yes / n_resp * 100) if n_resp else 0,
            "res_pct":   round(res_yes  / n_resp * 100) if n_resp else 0,
            "top_dup":   top_dup,
            "top_intg":  top_intg,
            "comments":  comments,
            "composite": round((n_sim / n_resp) * avg_impact, 3) if n_resp and avg_impact else 0,
        }

    return results

# ─────────────────────────────────────────────────────────────────────────────
# ESTATÍSTICAS SUMÁRIAS
# ─────────────────────────────────────────────────────────────────────────────

def summary_stats(results, df, n_experts_expected=None):
    n_experts_observed = df["expert_code"].nunique()
    n_experts = n_experts_expected if n_experts_expected and n_experts_expected > 0 else n_experts_observed
    n_inv_80    = sum(1 for r in results.values() if r["pct_optimizable"] >= 80)
    n_imp_high  = sum(1 for r in results.values() if r["avg_impact"] >= 2.5)
    n_unanimous = sum(1 for r in results.values() if r["pct_optimizable"] == 100)

    # Response rate per intervention: n_total / n_experts (% of experts who answered)
    n_experts_int = n_experts if n_experts > 0 else 1
    rr = [r["n_total"] / n_experts_int * 100 for r in results.values() if r["n_total"] > 0]
    rr_median = round(statistics.median(rr), 1) if rr else 0
    rr_min    = round(min(rr), 1)               if rr else 0
    rr_max    = round(max(rr), 1)               if rr else 0

    # Per-intervention breakdown for the response rate table
    rr_detail = sorted(
        [{"label": r["label"], "n": r["n_total"], "pct": round(r["n_total"] / n_experts_int * 100, 1)}
         for r in results.values()],
        key=lambda x: x["pct"], reverse=True
    )

    return {
        "n_experts":   n_experts,
      "n_experts_observed": n_experts_observed,
        "n_inv":       len(results),
        "n_inv_80":    n_inv_80,
        "n_imp_high":  n_imp_high,
        "n_unanimous": n_unanimous,
        "rr_median":   rr_median,
        "rr_min":      rr_min,
        "rr_max":      rr_max,
        "rr_detail":   rr_detail,
    }

# ─────────────────────────────────────────────────────────────────────────────
# GERAÇÃO DE HTML
# ─────────────────────────────────────────────────────────────────────────────

def esc(s):
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"','&quot;')

def render_html(results, stats, interventions, source_file):
    inv_label = {r["code"]: r["label"] for r in interventions}
    sorted_inv = sorted(results.values(), key=lambda r: r["composite"], reverse=True)
    now = datetime.now().strftime("%d/%m/%Y às %H:%M")

    def tier(r):
        if r["pct_definitely"] >= 60 and r["avg_impact"] >= 2.5: return "alta"
        if r["pct_optimizable"] >= 65 and r["avg_impact"] >= 1.8: return "media"
        return "baixa"

    alta  = [r["label"] for r in sorted_inv if tier(r) == "alta"]
    media = [r["label"] for r in sorted_inv if tier(r) == "media"]
    baixa = [r["label"] for r in sorted_inv if tier(r) == "baixa"]

    # ── Response rate section ──────────────────────────────────────────────
    rr_rows = ""
    n_experts = stats["n_experts"]
    n_experts_observed = stats.get("n_experts_observed", n_experts)
    for item in stats["rr_detail"]:
        bar_w = round(item["pct"])
        color = "#1a6b3a" if item["pct"] >= 80 else ("#d4a017" if item["pct"] >= 60 else "#e57373")
        rr_rows += f"""<tr>
          <td style="max-width:320px;font-size:12px">{esc(item["label"])}</td>
          <td style="text-align:center;font-weight:600">{item["n"]}/{n_experts}</td>
          <td style="width:200px">
            <div style="display:flex;align-items:center;gap:8px">
              <div style="flex:1;height:8px;background:#e5e7eb;border-radius:2px;overflow:hidden">
                <div style="width:{bar_w}%;height:100%;background:{color}"></div>
              </div>
              <span style="font-size:11px;font-weight:600;color:{color};white-space:nowrap">{item["pct"]}%</span>
            </div>
          </td>
        </tr>"""

    # ── Ranking table rows ──────────────────────────────────────────────────
    rank_rows = ""
    for i, r in enumerate(sorted_inv, 1):
        sim_w  = round(r["n_sim"]  / r["n_total"] * 100) if r["n_total"] else 0
        poss_w = round(r["n_poss"] / r["n_total"] * 100) if r["n_total"] else 0
        nao_w  = round(r["n_nao"]  / r["n_total"] * 100) if r["n_total"] else 0
        imp    = r["avg_impact"]
        imp_cls = "imp-high" if imp >= 2.5 else ("imp-med" if imp >= 1.8 else "imp-low")
        imp_lbl = "Alto" if imp >= 2.5 else ("Médio" if imp >= 1.8 else "Baixo")
        name_cell = (f'<a href="{esc(r["url"])}" target="_blank" class="inv-link">{esc(r["label"])}</a>'
                     if r["url"] else esc(r["label"]))
        rank_rows += f"""<tr>
          <td class="rank-num">{i}</td>
          <td style="font-weight:500">{name_cell}</td>
          <td><span class="comp-tag">{esc(r["component"])}</span></td>
          <td>
            <div class="mini-bar-wrap">
              <div class="mini-bar-bg">
                <div style="display:flex;height:100%">
                  <div style="width:{sim_w}%;background:var(--sim)"></div>
                  <div style="width:{poss_w}%;background:#d4a017"></div>
                  <div style="width:{nao_w}%;background:#e5e7eb"></div>
                </div>
              </div>
              <span class="bar-counts">{r["n_sim"]}S·{r["n_poss"]}P·{r["n_nao"]}N</span>
            </div>
          </td>
          <td><strong style="font-size:14px">{r["pct_optimizable"]}%</strong></td>
          <td><span class="impact-badge {imp_cls}">{imp:.1f} — {imp_lbl}</span></td>
          <td style="color:var(--muted)">{r["dup_pct"]}%</td>
          <td style="color:var(--muted)">{r["intg_pct"]}%</td>
        </tr>"""

    # ── Detail cards grouped by component ──────────────────────────────────
    by_comp = OrderedDict()
    for inv in interventions:
        comp = inv["component"] or "Outros"
        if comp not in by_comp:
            by_comp[comp] = []
        if inv["code"] in results:
            by_comp[comp].append(results[inv["code"]])

    cards_html = ""
    for comp, items in by_comp.items():
        if not items:
            continue
        cards_html += f'<div class="comp-label">{esc(comp)}</div>\n<div class="priority-grid">\n'
        for r in items:
            sim_p  = round(r["n_sim"]  / r["n_total"] * 100) if r["n_total"] else 0
            poss_p = round(r["n_poss"] / r["n_total"] * 100) if r["n_total"] else 0
            nao_p  = round(r["n_nao"]  / r["n_total"] * 100) if r["n_total"] else 0
            imp    = r["avg_impact"]
            imp_cls = "mc-impact-high" if imp >= 2.5 else ("mc-impact-med" if imp >= 1.8 else "mc-impact-low")

            def make_tag(code, cls):
                lbl = inv_label.get(code, code)
                return f'<span class="{cls}">{esc(lbl)}</span>'

            dup_tags  = "".join(make_tag(t, "dup-tag")  for t in r["top_dup"])
            intg_tags = "".join(make_tag(t, "intg-tag") for t in r["top_intg"])
            dup_row  = (f'<div class="tag-row"><span class="tag-label dup-label">Duplicação com:</span>'
                        f'{dup_tags}</div>') if r["top_dup"]  else ""
            intg_row = (f'<div class="tag-row"><span class="tag-label">Integrar com:</span>'
                        f'{intg_tags}</div>') if r["top_intg"] else ""

            cmt_html = ""
            if r["comments"]:
                cmt_html = '<div class="comments-section"><div class="cmt-label">Sugestões dos especialistas (amostra)</div>'
                for c in r["comments"][:2]:
                    cmt_html += f'<div class="cmt-item">"{esc(c)}"</div>'
                if len(r["comments"]) > 2:
                    cmt_html += f'<div class="cmt-more">+{len(r["comments"])-2} comentário(s) adicionais</div>'
                cmt_html += '</div>'

            missing_note = (f'<div class="missing-note">⚠ {r["n_missing"]} resposta(s) em falta / '
                            f'não aplicável</div>') if r["n_missing"] else ""

            name_cell = (f'<a href="{esc(r["url"])}" target="_blank" class="inv-link">{esc(r["label"])}</a>'
                         if r["url"] else esc(r["label"]))

            cards_html += f"""<div class="inv-card">
              <div class="card-top">
                <span class="comp-tag">{esc(r["component"])}</span>
              </div>
              <h3>{name_cell}</h3>
              <div class="gate-label-sm">Precisa de optimização? (n={r["n_total"]})</div>
              <div class="gate-bar-wrap">
                <div class="gate-seg seg-sim"  style="width:{sim_p}%"></div>
                <div class="gate-seg seg-poss" style="width:{poss_p}%"></div>
                <div class="gate-seg seg-nao"  style="width:{nao_p}%"></div>
              </div>
              <div class="gate-counts">
                <div class="gc"><div class="dot" style="background:var(--sim)"></div><span class="val">{r["n_sim"]}</span><span class="gc-lbl">Sim def. ({sim_p}%)</span></div>
                <div class="gc"><div class="dot" style="background:#d4a017"></div><span class="val">{r["n_poss"]}</span><span class="gc-lbl">Possiv. ({poss_p}%)</span></div>
                <div class="gc"><div class="dot" style="background:#e0e0e0"></div><span class="val">{r["n_nao"]}</span><span class="gc-lbl">Não ({nao_p}%)</span></div>
              </div>
              <div class="metrics-row">
                <div class="metric-chip"><div class="mc-val {imp_cls}">{imp:.1f}<span class="mc-denom">/3</span></div><div class="mc-lbl">Impacto médio</div></div>
                <div class="metric-chip"><div class="mc-val" style="color:var(--ink)">{r["dup_pct"]}%</div><div class="mc-lbl">Duplicação</div></div>
                <div class="metric-chip"><div class="mc-val" style="color:var(--accent2)">{r["intg_pct"]}%</div><div class="mc-lbl">Integração</div></div>
                <div class="metric-chip"><div class="mc-val" style="color:var(--muted)">{r["res_pct"]}%</div><div class="mc-lbl">↓ Recursos</div></div>
              </div>
              {dup_row}{intg_row}
              {cmt_html}
              {missing_note}
            </div>\n"""
        cards_html += "</div>\n"

    alta_str  = " · ".join(alta)  or "—"
    media_str = " · ".join(media) or "—"
    baixa_str = " · ".join(baixa) or "—"

    html = f"""<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Delphi W1 — Resultados Agregados</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root {{
  --ink:#0f1923; --paper:#f5f2ed; --accent:#c0392b; --accent2:#1a5276;
  --gold:#b7860b; --muted:#6b7280; --border:#d5cfc6;
  --sim:#1a6b3a; --sim-bg:#e8f5ee; --poss:#7a5c00; --poss-bg:#fef9e7;
  --card-bg:#ffffff;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'DM Sans',sans-serif;background:var(--paper);color:var(--ink);line-height:1.5;font-size:14px}}
.site-header{{background:var(--ink);color:white;padding:0 40px;display:flex;align-items:stretch;border-bottom:4px solid var(--accent)}}
.header-brand{{padding:24px 0;border-right:1px solid rgba(255,255,255,.12);padding-right:32px;margin-right:32px}}
.header-brand .eyebrow{{font-size:10px;letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,.5);margin-bottom:4px}}
.header-brand h1{{font-family:'DM Serif Display',serif;font-size:22px;font-weight:400;color:white;line-height:1.2}}
.header-meta{{padding:24px 0;display:flex;flex-direction:column;justify-content:center;gap:4px}}
.wave-badge{{display:inline-flex;align-items:center;gap:8px;background:var(--accent);color:white;font-size:11px;font-weight:600;letter-spacing:.08em;padding:3px 10px;border-radius:2px;width:fit-content;margin-bottom:6px}}
.header-meta p{{font-size:12px;color:rgba(255,255,255,.6)}}
.header-meta strong{{color:rgba(255,255,255,.9)}}
.intro-band{{background:var(--accent2);color:white;padding:20px 40px;display:flex;align-items:center;gap:20px}}
.intro-band .icon{{font-size:28px;flex-shrink:0}}
.intro-band p{{font-size:13px;line-height:1.6;color:rgba(255,255,255,.88)}}
.intro-band strong{{color:white}}
.summary-bar{{background:white;border-bottom:1px solid var(--border);padding:20px 40px;display:flex;gap:32px;align-items:center;flex-wrap:wrap}}
.stat-pill{{display:flex;flex-direction:column;align-items:center;gap:2px}}
.stat-pill .num{{font-family:'DM Serif Display',serif;font-size:32px;line-height:1;color:var(--accent)}}
.stat-pill .lbl{{font-size:10px;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);text-align:center;max-width:80px}}
.stat-divider{{width:1px;height:40px;background:var(--border)}}
.legend-box{{margin-left:auto;display:flex;gap:16px;align-items:center;flex-wrap:wrap}}
.legend-item{{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--muted)}}
.legend-dot{{width:10px;height:10px;border-radius:50%}}
.dot-sim{{background:var(--sim)}}.dot-poss{{background:#d4a017}}.dot-nao{{background:#d1d5db}}
.main{{padding:32px 40px;max-width:1300px;margin:0 auto}}
.section-header{{display:flex;align-items:baseline;gap:12px;margin-bottom:20px;padding-bottom:10px;border-bottom:2px solid var(--ink)}}
.section-header h2{{font-family:'DM Serif Display',serif;font-size:20px;font-weight:400}}
.section-header .subtitle{{font-size:12px;color:var(--muted);letter-spacing:.04em}}
.note-box{{background:#fffbeb;border:1px solid #fcd34d;border-left:4px solid var(--gold);border-radius:3px;padding:14px 18px;margin-bottom:28px;font-size:12px;color:#78350f;line-height:1.6}}
.note-box strong{{color:#92400e}}
/* ── Response rate section ── */
.rr-panel{{background:white;border:1px solid var(--border);border-radius:4px;margin-bottom:40px;overflow:hidden}}
.rr-summary{{display:flex;gap:0;border-bottom:1px solid var(--border)}}
.rr-stat{{flex:1;padding:16px 20px;text-align:center;border-right:1px solid var(--border)}}
.rr-stat:last-child{{border-right:none}}
.rr-stat .rr-num{{font-family:'DM Serif Display',serif;font-size:28px;line-height:1;color:var(--accent2)}}
.rr-stat .rr-lbl{{font-size:10px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-top:2px}}
.rr-table{{width:100%;border-collapse:collapse;font-size:12px}}
.rr-table td{{padding:7px 16px;border-bottom:1px solid #f0ede8;vertical-align:middle}}
.rr-table tr:last-child td{{border-bottom:none}}
.rr-table tr:nth-child(even) td{{background:#fafaf8}}
.rr-details{{margin-top:12px}}
.rr-details summary{{cursor:pointer;padding:10px 16px;background:#f7f9fb;border-top:1px solid var(--border);font-size:12px;font-weight:600;color:var(--accent2);user-select:none;display:flex;align-items:center;gap:8px}}
.rr-details summary:hover{{background:#eef2f7}}
.rr-details summary::marker{{content:''}}
.rr-details summary::before{{content:'▶';font-size:10px;transition:transform 0.2s}}
.rr-details[open] summary::before{{transform:rotate(90deg)}}
/* ── Ranking table ── */
.rank-table{{width:100%;border-collapse:collapse;margin-bottom:40px;background:white;border:1px solid var(--border);border-radius:4px;overflow:hidden;font-size:12px}}
.rank-table thead tr{{background:var(--ink);color:white}}
.rank-table th{{padding:10px 14px;font-size:10px;text-transform:uppercase;letter-spacing:.1em;font-weight:600;text-align:left}}
.rank-table td{{padding:9px 14px;border-bottom:1px solid var(--border);vertical-align:middle}}
.rank-table tr:last-child td{{border-bottom:none}}
.rank-table tr:nth-child(even) td{{background:#fafafa}}
.rank-num{{font-family:'DM Serif Display',serif;font-size:18px;color:var(--muted);width:32px}}
.comp-tag{{font-size:10px;background:var(--paper);border:1px solid var(--border);color:var(--muted);padding:2px 6px;border-radius:2px;white-space:nowrap}}
.inv-link{{color:inherit;text-decoration:none;border-bottom:1px dotted var(--muted)}}
.inv-link:hover{{color:var(--accent2);border-bottom-color:var(--accent2)}}
.mini-bar-wrap{{display:flex;align-items:center;gap:8px}}
.mini-bar-bg{{flex:1;height:8px;background:#e5e7eb;border-radius:2px;overflow:hidden}}
.bar-counts{{font-size:10px;color:var(--muted);white-space:nowrap}}
.impact-badge{{display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:600;padding:3px 8px;border-radius:2px}}
.imp-high{{background:var(--sim-bg);color:var(--sim)}}
.imp-med{{background:var(--poss-bg);color:var(--poss)}}
.imp-low{{background:#f3f4f6;color:var(--muted)}}
/* ── Cards ── */
.comp-label{{font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--accent2);padding:6px 0;border-bottom:1px solid var(--border);margin-bottom:14px;margin-top:28px}}
.priority-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}}
.inv-card{{background:var(--card-bg);border:1px solid var(--border);border-radius:4px;padding:16px 18px}}
.inv-card:hover{{box-shadow:0 4px 16px rgba(0,0,0,.08)}}
.card-top{{display:flex;align-items:flex-start;justify-content:flex-end;gap:8px;margin-bottom:10px}}
.inv-card h3{{font-size:13px;font-weight:600;line-height:1.3;margin-bottom:12px}}
.gate-label-sm{{font-size:10px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:5px}}
.gate-bar-wrap{{display:flex;height:12px;border-radius:2px;overflow:hidden;gap:1px;margin-bottom:4px}}
.gate-seg{{height:100%}}
.seg-sim{{background:var(--sim)}}.seg-poss{{background:#d4a017}}.seg-nao{{background:#e0e0e0}}
.gate-counts{{display:flex;gap:12px;margin-bottom:2px}}
.gc{{display:flex;align-items:center;gap:4px;font-size:10px}}
.gc .dot{{width:7px;height:7px;border-radius:50%}}
.gc .val{{font-weight:600}}
.gc-lbl{{color:var(--muted)}}
.metrics-row{{display:flex;gap:8px;margin-top:12px;padding-top:12px;border-top:1px solid var(--border)}}
.metric-chip{{flex:1;background:var(--paper);border-radius:3px;padding:6px 8px;text-align:center}}
.mc-val{{font-family:'DM Serif Display',serif;font-size:16px;line-height:1;margin-bottom:2px}}
.mc-denom{{font-size:10px;font-family:'DM Sans',sans-serif}}
.mc-lbl{{font-size:9px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted)}}
.mc-impact-high{{color:var(--sim)}}.mc-impact-med{{color:var(--poss)}}.mc-impact-low{{color:#9ca3af}}
.tag-row{{margin-top:8px;display:flex;flex-wrap:wrap;gap:4px;align-items:center}}
.tag-label{{font-size:9px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-right:2px}}
.dup-label{{color:#9b2226}}
.intg-tag{{font-size:10px;font-weight:600;background:#e8f0fe;color:#1a3a8f;padding:2px 7px;border-radius:2px;white-space:normal;line-height:1.4}}
.dup-tag{{font-size:10px;font-weight:600;background:#fde8e8;color:#9b2226;padding:2px 7px;border-radius:2px;white-space:normal;line-height:1.4}}
.comments-section{{margin-top:10px;padding-top:10px;border-top:1px solid var(--border)}}
.cmt-label{{font-size:9px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:5px}}
.cmt-item{{font-size:11px;color:var(--ink);font-style:italic;line-height:1.5;margin-bottom:4px;padding-left:8px;border-left:2px solid var(--border)}}
.cmt-more{{font-size:10px;color:var(--muted);margin-top:3px}}
.missing-note{{font-size:10px;color:#b45309;margin-top:6px;background:#fffbeb;padding:3px 6px;border-radius:2px}}
/* ── Next steps ── */
.next-box{{background:white;border:1px solid var(--border);border-radius:4px;padding:24px;margin-bottom:40px}}
.next-box p{{font-size:13px;line-height:1.8}}
.tier-panels{{margin-top:16px;padding-top:16px;border-top:1px solid var(--border);display:flex;gap:16px;flex-wrap:wrap}}
.tier-panel{{flex:1;min-width:180px;border-radius:3px;padding:12px 16px}}
.tier-alta{{background:var(--sim-bg)}}.tier-media{{background:var(--poss-bg)}}.tier-baixa{{background:#f3f4f6}}
.tier-heading{{font-size:10px;text-transform:uppercase;letter-spacing:.1em;font-weight:700;margin-bottom:6px}}
.tier-alta .tier-heading{{color:var(--sim)}}.tier-media .tier-heading{{color:var(--poss)}}.tier-baixa .tier-heading{{color:var(--muted)}}
.tier-panel p{{font-size:12px;line-height:1.8}}
.report-footer{{background:var(--ink);color:rgba(255,255,255,.5);padding:20px 40px;font-size:11px;display:flex;justify-content:space-between;align-items:center;margin-top:40px;flex-wrap:wrap;gap:8px}}
.report-footer strong{{color:rgba(255,255,255,.8)}}
@media print{{body{{background:white}}.site-header,.intro-band{{-webkit-print-color-adjust:exact;print-color-adjust:exact}}.inv-card{{break-inside:avoid}}}}
</style>
</head>
<body>

<header class="site-header">
  <div class="header-brand">
    <div class="eyebrow">Programa Nacional de Controlo da Malária</div>
    <h1>Exercício de Optimização<br>de Intervenções</h1>
  </div>
  <div class="header-meta">
    <div class="wave-badge">⬤ &nbsp;ONDA 1 — RESULTADOS AGREGADOS</div>
    <p>Gerado em: <strong>{now}</strong> &nbsp;|&nbsp; Fonte: <strong>{esc(os.path.basename(source_file))}</strong></p>
    <p>Base esperada: <strong>{stats["n_experts"]} especialistas</strong> (experts.txt) &nbsp;|&nbsp; Respondentes observados: <strong>{n_experts_observed}</strong> &nbsp;|&nbsp; Intervenções avaliadas: <strong>{stats["n_inv"]}</strong></p>
    <p>Este documento é <strong>confidencial</strong> — para uso exclusivo dos participantes da oficina Delphi</p>
  </div>
</header>

<div class="intro-band">
  <div class="icon">📋</div>
  <p>Este relatório apresenta os <strong>resultados anónimos e agregados</strong> da primeira onda (W1).
  Os dados reflectem as avaliações colectivas dos especialistas sobre o potencial de optimização de cada intervenção.
  <strong>Nenhuma resposta individual é identificável.</strong>
  Com base nestes resultados, o grupo seleccionará as intervenções prioritárias para a Onda 2 (W2).</p>
</div>

<div class="summary-bar">
  <div class="stat-pill"><span class="num">{stats["n_experts"]}</span><span class="lbl">Especialistas esperados</span></div>
  <div class="stat-divider"></div>
  <div class="stat-pill"><span class="num">{n_experts_observed}</span><span class="lbl">Especialistas respondentes</span></div>
  <div class="stat-divider"></div>
  <div class="stat-pill"><span class="num">{stats["n_inv"]}</span><span class="lbl">Intervenções avaliadas</span></div>
  <div class="stat-divider"></div>
  <div class="stat-pill"><span class="num">{stats["n_inv_80"]}</span><span class="lbl">Com ≥80% a favor de optimização</span></div>
  <div class="stat-divider"></div>
  <div class="stat-pill"><span class="num">{stats["n_imp_high"]}</span><span class="lbl">Com impacto esperado alto (≥2,5)</span></div>
  <div class="stat-divider"></div>
  <div class="stat-pill"><span class="num">{stats["n_unanimous"]}</span><span class="lbl">Com consenso unânime (100%)</span></div>
  <div class="legend-box">
    <div class="legend-item"><div class="legend-dot dot-sim"></div>Sim, definitivamente</div>
    <div class="legend-item"><div class="legend-dot dot-poss"></div>Possivelmente</div>
    <div class="legend-item"><div class="legend-dot dot-nao"></div>Não</div>
  </div>
</div>

<div class="main">

  <div class="note-box">
    <strong>Como ler este relatório:</strong> Para cada intervenção são apresentados:
    (1) a distribuição de respostas à pergunta de triagem ("precisa de optimização?"),
    (2) o impacto médio esperado se optimizada (escala 1–3),
    e (3) as percentagens de especialistas que identificaram duplicação, potencial de integração
    e possibilidade de redução de recursos.
    As intervenções estão ordenadas por <em>pontuação composta</em> (% "Sim definitivamente" × impacto médio).
    S = Sim definitivamente &nbsp;|&nbsp; P = Possivelmente &nbsp;|&nbsp; N = Não.
    Os nomes das intervenções são hiperligações para as respectivas fichas de descrição.
  </div>

  <!-- ═══ RESPONSE RATE ═══════════════════════════════════════════════════ -->
  <div class="section-header">
    <h2>Taxa de Resposta por Intervenção</h2>
    <span class="subtitle">Proporção de especialistas que avaliaram cada intervenção</span>
  </div>

  <div class="rr-panel">
    <div class="rr-summary">
      <div class="rr-stat">
        <div class="rr-num">{stats["rr_median"]}%</div>
        <div class="rr-lbl">Mediana</div>
      </div>
      <div class="rr-stat">
        <div class="rr-num">{stats["rr_min"]}%</div>
        <div class="rr-lbl">Mínimo</div>
      </div>
      <div class="rr-stat">
        <div class="rr-num">{stats["rr_max"]}%</div>
        <div class="rr-lbl">Máximo</div>
      </div>
      <div class="rr-stat" style="flex:2;text-align:left;padding-left:24px">
        <div style="font-size:12px;color:var(--muted);line-height:1.6">
          A taxa de resposta por intervenção varia porque o formulário W1 está dividido
          em grupos de intervenções, e nem todos os especialistas completaram todos os grupos.
          Intervenções com taxa abaixo de 60% devem ser interpretadas com cautela.
        </div>
      </div>
    </div>
    <details class="rr-details">
      <summary>Ver detalhes por intervenção ({stats["n_inv"]} intervenções)</summary>
      <table class="rr-table">
        <tbody>{rr_rows}</tbody>
      </table>
    </details>
  </div>

  <!-- ═══ RANKING TABLE ════════════════════════════════════════════════════ -->
  <div class="section-header">
    <h2>Tabela de Prioridade — Todas as Intervenções</h2>
    <span class="subtitle">Ordenadas por pontuação composta</span>
  </div>

  <table class="rank-table">
    <thead>
      <tr>
        <th>#</th><th>Intervenção</th><th>Componente</th>
        <th style="width:200px">Distribuição de respostas</th>
        <th>% Optimizável</th><th>Impacto esperado</th>
        <th>% Duplicação</th><th>% Integração</th>
      </tr>
    </thead>
    <tbody>{rank_rows}</tbody>
  </table>

  <!-- ═══ DETAIL CARDS ════════════════════════════════════════════════════ -->
  <div class="section-header">
    <h2>Detalhe por Intervenção</h2>
    <span class="subtitle">Agrupado por componente programático</span>
  </div>

  {cards_html}

  <!-- ═══ NEXT STEPS ══════════════════════════════════════════════════════ -->
  <div class="section-header">
    <h2>Próximos Passos — Onda 2</h2>
  </div>
  <div class="next-box">
    <p>Com base nos resultados da W1, o grupo irá <strong>seleccionar colectivamente as intervenções a focar na W2</strong>.
    Recomenda-se priorizar intervenções com elevada percentagem de "Sim, definitivamente" combinada com impacto esperado alto (≥2,5).
    As intervenções seleccionadas serão distribuídas por grupos de trabalho temáticos que irão desenvolver <strong>propostas concretas de optimização</strong>.
    Essas propostas serão submetidas à avaliação anónima colectiva na Onda 3 (W3).</p>
    <div class="tier-panels">
      <div class="tier-panel tier-alta">
        <div class="tier-heading">Candidatas de Alta Prioridade</div>
        <p>{esc(alta_str)}</p>
      </div>
      <div class="tier-panel tier-media">
        <div class="tier-heading">Candidatas de Prioridade Média</div>
        <p>{esc(media_str)}</p>
      </div>
      <div class="tier-panel tier-baixa">
        <div class="tier-heading">Baixa Prioridade / Manter</div>
        <p>{esc(baixa_str)}</p>
      </div>
    </div>
  </div>

</div>

<div class="report-footer">
  <span>Programa Nacional de Controlo da Malária — INS &nbsp;|&nbsp; Exercício Delphi 2026 &nbsp;|&nbsp; <strong>CONFIDENCIAL — Uso interno</strong></span>
  <span>Resultados anónimos. Nenhuma resposta individual é identificável. &nbsp;|&nbsp; Gerado em {now}</span>
</div>

</body>
</html>"""
    return html

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    results_path = sys.argv[1]
    dict_path    = None
    output_dir   = None

    # Parse optional arguments
    for i in range(2, len(sys.argv)):
        if sys.argv[i].startswith("--output-dir") or sys.argv[i].startswith("-o"):
            if sys.argv[i] in ("--output-dir", "-o") and i + 1 < len(sys.argv):
                output_dir = sys.argv[i + 1]
            elif "=" in sys.argv[i]:
                output_dir = sys.argv[i].split("=", 1)[1]
        elif not sys.argv[i].startswith("-"):
            if dict_path is None:
                dict_path = sys.argv[i]

    # Default output directory: ./reports relative to current working directory
    if output_dir is None:
        output_dir = "reports"
    output_dir = os.path.abspath(output_dir)

    if not os.path.exists(results_path):
        sys.exit(f"Erro: ficheiro de resultados não encontrado: {results_path}")
    if dict_path and not os.path.exists(dict_path):
        print(f"Aviso: dicionário não encontrado: {dict_path} — continuando sem ele.")
        dict_path = None

    expected_experts = load_expected_experts()
    n_expected_experts = len(expected_experts)

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    print(f"A carregar resultados: '{results_path}'...")
    df = load_data(results_path, "Responses")
    print(f"  {len(df)} linhas · {df['expert_code'].nunique()} especialistas únicos")
    if n_expected_experts > 0:
      print(f"  Base esperada (experts.txt): {n_expected_experts} especialistas")
    else:
      print("  Aviso: experts.txt não encontrado ou vazio; usando especialistas observados.")

    print("A carregar metadados das intervenções...")
    interventions = load_metadata(results_path, dict_path)
    if interventions is None:
        df_codes = sorted(df["intervention"].dropna().unique())
        interventions = [{"code": c, "label": c, "component": "", "url": ""} for c in df_codes]
        print(f"  Metadados: códigos brutos da folha Responses ({len(interventions)} intervenções)")

    print("A agregar respostas...")
    results = aggregate(df, interventions)
    stats   = summary_stats(results, df, n_expected_experts)

    print(f"  {stats['n_experts']} especialistas · {stats['n_inv']} intervenções")
    print(f"  Taxa de resposta: mediana={stats['rr_median']}% "
          f"intervalo=[{stats['rr_min']}%–{stats['rr_max']}%]")

    print("A gerar relatório HTML...")
    html = render_html(results, stats, interventions, results_path)

    ts       = datetime.now().strftime("%Y%m%d_%H%M")
    out_path = os.path.join(output_dir, f"delphi_w1_relatorio_{ts}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✓ Relatório guardado em: {out_path}")

if __name__ == "__main__":
    main()
