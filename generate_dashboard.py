#!/usr/bin/env python3
"""
OnePageReport Dashboard Generator v3 - Template Filling Approach
Copy the spreadsheet template and populate data tabs using a raw export
from Microsoft Planner / Teams board.

Usage:
    python generate_dashboard.py <input_file.xlsx> [output_file.xlsx] \
        [--template template.xlsx] [--author "Nome - Cargo"]
"""

import sys
import os
import io
import re
import shutil
import argparse
import calendar
import unicodedata
import zipfile
from bisect import bisect_right
from datetime import datetime, timedelta, date
from collections import defaultdict
from functools import lru_cache

# ─── UTF-8 stdout wrapper (Windows cp1252 fix) ────────────────────────────────
if hasattr(sys.stdout, "buffer") and sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import pandas as pd
from openpyxl import load_workbook

# ─────────────────────────────── KEYWORDS ────────────────────────────────────
AREA_KEYWORDS = {
    "Gestao":       ["gp","gestao","gestão","gerente","gerencia","coordena"],
    "Requisitos":   ["requisito","analista de negocio"],
    "Testes":       ["teste","q/a","qa","qualidade","testador"],
    "UX":           ["ux","ui","design","designer"],
    "Devs":         ["dev","frontend","mobile","desktop","backend","programad","desenvolv"],
    "Arquitetura":  ["arquit"],
    "DB":           ["db","banco de dados","dba"],
    "Publicacao":   ["public","deploy","release"],
    "Revisao":      ["revis","review"],
    "Construcao":   ["constr","constru"],
    "Prototipo":    ["proto"],
    "Pesquisa":     ["bench","pesquisa"],
}

# Map internal area keys to template display names
AREA_DISPLAY = {
    "Gestao":      "Gestão",
    "Requisitos":  "Requisitos",
    "Testes":      "Testes",
    "UX":          "UX",
    "Devs":        "Devs",
    "Arquitetura": "Arquitetura",
    "DB":          "DB",
    "Publicacao":  "Publicação",
    "Revisao":     "Revisão",
    "Construcao":  "Construção",
    "Prototipo":   "Prototipo",
    "Pesquisa":    "Pesquisa",
}

TEMPLATE_AREAS = [
    # Primeiras 6 = range do gráfico PERCENTUAIS (Areas!B2:B7)
    "Gestão", "Requisitos", "Testes", "UX", "Devs", "Revisão",
    # Demais (fora do range do gráfico, mas preenchidos na sheet)
    "Arquitetura", "DB", "Publicação", "Construção", "Prototipo", "Pesquisa",
]

NAO_PREVISTO_KW = ["nao previsto","não previsto","nao mapeado","não mapeado","nao_previsto"]
BACKLOG_KW      = ["backlog","incremento"]
EXCLUDE_CAT     = {"cancelado","residual","bugs e ajustes","cms","banco de dados",
                   "nao previsto","não previsto","backlog","incremento"}


# ─────────────────────────────── HELPERS ─────────────────────────────────────
HU_ID_RE        = re.compile(r"(HU\s*\d+)", re.IGNORECASE)
HU_LABEL_RE     = re.compile(r"HU\s*\d+", re.IGNORECASE)
SPRINT_GOAL_RE  = re.compile(r"(?i)^sprint\s+goal\s*[:\-]\s*")
NAO_PREV_RE     = re.compile("|".join(re.escape(k) for k in NAO_PREVISTO_KW), re.IGNORECASE)
BACKLOG_RE      = re.compile("|".join(re.escape(k) for k in BACKLOG_KW), re.IGNORECASE)


def _strip(v):
    if v is None:
        return ""
    return str(v).strip()

def _lower(v):
    return _strip(v).lower()

def _norm_text_key(v):
    """Normalize text for resilient Excel header/sheet matching."""
    t = _strip(v)
    if not t:
        return ""
    t = unicodedata.normalize("NFKD", t)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", t).strip().lower()

def _find_excel_sheet(xl, *names):
    wanted = {_norm_text_key(name) for name in names}
    for sheet in xl.sheet_names:
        if _norm_text_key(sheet) in wanted:
            return sheet
    return None

def _find_column(columns, *aliases):
    wanted = {_norm_text_key(alias) for alias in aliases}
    for col in columns:
        if _norm_text_key(col) in wanted:
            return col
    return None

def _first_non_empty(df, col):
    if col is None or col not in df.columns:
        return ""
    vals = df[col].dropna().astype(str).map(str.strip)
    vals = vals[(vals != "") & (vals.str.lower() != "nan")]
    return vals.iloc[0] if not vals.empty else ""

def _parse_date(val):
    if not val or str(val).strip().lower() in ("", "nan", "none"):
        return None
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None

def _parse_effort(checklist_items_cell, completed_cell):
    if completed_cell and "/" in str(completed_cell):
        try:
            return int(str(completed_cell).split("/")[1])
        except Exception:
            pass
    if checklist_items_cell and str(checklist_items_cell).strip():
        s = str(checklist_items_cell).strip()
        return s.count(";") + 1
    return 0

def _is_nao_previsto(text):
    b = _lower(text)
    for kw in NAO_PREVISTO_KW:
        if kw in b:
            return True
    return False

def _is_backlog(bucket):
    b = _lower(bucket)
    for kw in BACKLOG_KW:
        if kw in b:
            return True
    return False

def _area_for_assignee(assignee):
    a = _norm_text_key(assignee)
    for area, kws in AREA_KEYWORDS.items():
        for kw in kws:
            if _norm_text_key(kw) in a:
                return area
    return "Outros"


def _area_for_label_token(label):
    key = _norm_text_key(label).lstrip(".")
    if not key or key in {"fluxo.continuo", "fluxo continuo"}:
        return None
    if key.startswith("hu"):
        return None

    exact = {
        "gp": "Gestao",
        "gestao": "Gestao",
        "gestao projeto": "Gestao",
        "requisitos": "Requisitos",
        "req": "Requisitos",
        "q/a": "Testes",
        "q/a testes": "Testes",
        "qa": "Testes",
        "testes": "Testes",
        "teste": "Testes",
        "ux": "UX",
        "ui": "UX",
        "ux/ui": "UX",
        "dev": "Devs",
        "frontend": "Devs",
        "front-end": "Devs",
        "backend": "Devs",
        "back-end": "Devs",
        "mobile": "Devs",
        "arquitetura": "Arquitetura",
        "publicacao": "Publicacao",
        "deploy": "Publicacao",
        "revisao": "Revisao",
    }
    if key in exact:
        return exact[key]

    for area, kws in AREA_KEYWORDS.items():
        for kw in kws:
            kw_key = _norm_text_key(kw)
            if not kw_key:
                continue
            pattern = r"(^|[\s._/\-])" + re.escape(kw_key) + r"($|[\s._/\-])"
            if re.search(pattern, key):
                return area
    return None


def _parse_date_series(series):
    """
    Vectorised date parsing with the exact same precedence as _parse_date:
    %d/%m/%Y, %Y-%m-%d, %m/%d/%Y, %d-%m-%Y.
    """
    s = series.fillna("").astype(str).str.strip()
    s = s.mask(s.str.lower().isin({"", "nan", "none"}))

    parsed = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    formats = ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y")
    for fmt in formats:
        missing = parsed.isna() & s.notna()
        if not missing.any():
            break
        parsed.loc[missing] = pd.to_datetime(s[missing], format=fmt, errors="coerce")

    out = parsed.dt.date
    return out.where(parsed.notna(), None)


@lru_cache(maxsize=4096)
def _extract_hu_cached(labels):
    """Extract first HU id-like label (e.g. HU044) from a semicolon label list."""
    if not labels or _lower(labels) in ("", "nan"):
        return ""

    parts = [p.strip() for p in str(labels).split(";") if p.strip()]
    for part in parts:
        if part.upper().startswith("HU"):
            m = HU_ID_RE.match(part)
            if m:
                return m.group(1).replace(" ", "").upper()
            return part.strip()
    return ""


@lru_cache(maxsize=4096)
def _area_for_labels_or_assignee_cached(labels, assignee):
    labels_text = _strip(labels)
    if labels_text and _lower(labels_text) not in ("", "nan"):
        parts = [p.strip() for p in labels_text.split(";") if p.strip()]
        for part in parts:
            area = _area_for_label_token(part)
            if area:
                return area
    return _area_for_assignee(assignee)


@lru_cache(maxsize=4096)
def _extract_hu_full_label(labels):
    """Return first full HU label token from a semicolon label list."""
    if not labels or _lower(labels) in ("", "nan"):
        return ""

    parts = [p.strip() for p in str(labels).split(";") if p.strip()]
    for part in parts:
        if HU_LABEL_RE.match(part):
            return part
    return ""


@lru_cache(maxsize=4096)
def _area_for_bucket_labels_assignee(bucket, labels, assignee):
    """
    Atribui área primária (mutuamente exclusiva) combinando bucket + labels.
    Prioridade:
      1. Bucket "Gestão"   → Gestao
      2. Label "GP"        → Gestao
      3. Bucket "Em Teste" → Revisao  (tarefas em revisão/teste)
      4. Demais            → baseado em labels/assignee
    """
    b_lower = _norm_text_key(bucket)
    if "gestão" in b_lower or "gestao" in b_lower:
        return "Gestao"
    # GP label check (tasks with GP outside of Gestão bucket)
    for part in [p.strip() for p in str(labels).split(";") if p.strip()]:
        area = _area_for_label_token(part)
        if area == "Gestao":
            return "Gestao"
    if "em teste" in b_lower or "homolog" in b_lower:
        return "Revisao"
    return _area_for_labels_or_assignee_cached(labels, assignee)


@lru_cache(maxsize=8192)
def _categories_from_labels(labels):
    """Return all categories from semicolon label list (including HU labels)."""
    if not labels or _lower(labels) in ("", "nan"):
        return ()

    parts = [p.strip() for p in str(labels).split(";") if p.strip()]
    return tuple(parts)


def _norm_label_key(text):
    """Normalize label for case-insensitive grouping (ignores accents/spaces)."""
    t = _strip(text)
    if not t:
        return ""
    t = unicodedata.normalize("NFKD", t)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = re.sub(r"\s+", " ", t).strip().lower()
    return t


def _row_has_label_key(labels_text, target_key):
    if not labels_text or not target_key:
        return False
    for token in _categories_from_labels(labels_text):
        if _norm_label_key(token) == target_key:
            return True
    return False


def _count_rows_with_label(df, label_text):
    key = _norm_label_key(label_text)
    if not key or "labels" not in df.columns:
        return 0
    return int(
        df["labels"].fillna("").astype(str).apply(lambda s: _row_has_label_key(s, key)).sum()
    )


def _extract_gp_storypoints_by_hu(wb):
    """
    Extrai storypoints por HU a partir da aba gp_Plann_Sprint.
    Soma todas as células numéricas por coluna de HU (ignorando N/A).
    Retorna dict {HUxxx: sp_total}.
    """
    sheet_name = None
    for cand in ("gp_Plann_Sprint", "Gp_Plann_Sprint"):
        if cand in wb.sheetnames:
            sheet_name = cand
            break
    if sheet_name is None:
        return {}

    ws = wb[sheet_name]
    hu_cols = []  # list[(col_idx, hu_short)]
    for c in range(2, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is None or str(v).strip() == "":
            continue
        m = re.search(r"\b(HU\d+)\b", str(v), flags=re.IGNORECASE)
        if m:
            hu_cols.append((c, m.group(1).upper()))

    if not hu_cols:
        return {}

    out = {hu: 0.0 for _, hu in hu_cols}
    for c, hu in hu_cols:
        s = 0.0
        for r in range(2, ws.max_row + 1):
            first = ws.cell(r, 1).value
            if first is None and r > 20:
                break
            raw = ws.cell(r, c).value
            try:
                if isinstance(raw, str) and raw.strip().upper() == "N/A":
                    continue
                if raw is None or str(raw).strip() == "":
                    continue
                s += float(str(raw).replace(",", "."))
            except Exception:
                continue
        out[hu] = round(s, 2)

    return out


def _enrich_kpis_with_hu_storypoints(kpis, df, hu_storypoints):
    """
    Atualiza storypoints e métricas HU/SP dos KPIs com base em HU + storypoints.
    """
    total_sp = round(sum(v for v in hu_storypoints.values() if v), 2)
    kpis["storypoints"] = total_sp

    hu_ct_means = []
    hu_lt_means = []
    weighted_ct_num = 0.0
    weighted_ct_den = 0.0
    weighted_lt_num = 0.0
    weighted_lt_den = 0.0

    for hu, grp in df[df["hu"] != ""].groupby("hu"):
        g_done = grp[grp["done_kpi"] & grp["date_done"].notna()]
        if g_done.empty:
            continue

        # Cycle = conclusão - início
        ct_vals = []
        for _, row in g_done[g_done["date_start"].notna()].iterrows():
            delta = (row["date_done"] - row["date_start"]).days
            if delta >= 0:
                ct_vals.append(delta)
        if ct_vals:
            ct_mean = sum(ct_vals) / len(ct_vals)
            hu_ct_means.append(ct_mean)
            sp = hu_storypoints.get(str(hu).upper(), 0) or 0
            if sp > 0:
                weighted_ct_num += ct_mean * sp
                weighted_ct_den += sp

        # Lead = conclusão - criação (com proteção início<criação)
        lt_vals = []
        for _, row in g_done[g_done["date_criacao"].notna()].iterrows():
            d_done = row["date_done"]
            d_cr = row["date_criacao"]
            d_start = row.get("date_start")
            base = max(d_cr, d_start) if (d_start is not None and not pd.isna(d_start)) else d_cr
            delta = (d_done - base).days
            if delta >= 0:
                lt_vals.append(delta)
        if lt_vals:
            lt_mean = sum(lt_vals) / len(lt_vals)
            hu_lt_means.append(lt_mean)
            sp = hu_storypoints.get(str(hu).upper(), 0) or 0
            if sp > 0:
                weighted_lt_num += lt_mean * sp
                weighted_lt_den += sp

    if hu_ct_means:
        kpis["ct_hu"] = round(sum(hu_ct_means) / len(hu_ct_means), 2)
    if hu_lt_means:
        kpis["lt_hu"] = round(sum(hu_lt_means) / len(hu_lt_means), 2)
    kpis["ct_sp"] = round(weighted_ct_num / weighted_ct_den, 2) if weighted_ct_den > 0 else kpis.get("ct_hu", 0)
    kpis["lt_sp"] = round(weighted_lt_num / weighted_lt_den, 2) if weighted_lt_den > 0 else kpis.get("lt_hu", 0)


# ─────────────────────────────── LOAD ────────────────────────────────────────
def _load_base_legacy_unused(path):
    """Legacy single-sheet loader kept for reference."""
    xl = pd.ExcelFile(path)
    sheet = xl.sheet_names[0]
    df = xl.parse(sheet, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    plan_name      = ""
    export_date_str = ""

    if "Nome do plano" in xl.sheet_names:
        try:
            pn = xl.parse("Nome do plano", dtype=str)
            # Column header at index 1 contains the plan name
            # e.g. columns = ['Nome do plano', 'Projeto Sprint 7 (Março/26)']
            if len(pn.columns) >= 2:
                plan_name = str(pn.columns[1]).strip()
            # Row where first col contains "exporta" has the export date
            mask = pn.iloc[:, 0].astype(str).str.lower().str.contains("exporta", na=False)
            if mask.any():
                export_date_str = str(pn.loc[mask].iloc[0, 1]).strip()
        except Exception:
            pass

    return df, plan_name, export_date_str


def load_base(path):
    """Load Planner/Teams export in legacy and current multi-sheet formats."""
    xl = pd.ExcelFile(path)
    plan_name = ""
    export_date_str = ""

    plan_sheet = _find_excel_sheet(xl, "Plano")
    if plan_sheet:
        try:
            pn = xl.parse(plan_sheet, dtype=str)
            pn.columns = [_strip(c) for c in pn.columns]
            plan_name = _first_non_empty(
                pn, _find_column(pn.columns, "Nome do plano", "Plan name")
            )
            export_date_str = _first_non_empty(
                pn, _find_column(pn.columns, "Data da exportacao", "Export date")
            )
        except Exception:
            pass

    legacy_plan_sheet = _find_excel_sheet(xl, "Nome do plano", "Plan name")
    if legacy_plan_sheet and (not plan_name or not export_date_str):
        try:
            pn = xl.parse(legacy_plan_sheet, dtype=str)
            pn.columns = [_strip(c) for c in pn.columns]
            if not plan_name:
                plan_col = _find_column(pn.columns, "Nome do plano", "Plan name")
                plan_name = _first_non_empty(pn, plan_col)
                if not plan_name and len(pn.columns) >= 2:
                    plan_name = _strip(pn.columns[1])
            if not export_date_str:
                export_col = _find_column(pn.columns, "Data da exportacao", "Export date")
                export_date_str = _first_non_empty(pn, export_col)
                if not export_date_str and len(pn.columns) >= 2:
                    mask = pn.iloc[:, 0].astype(str).map(_norm_text_key).str.contains(
                        "exporta", na=False
                    )
                    if mask.any():
                        export_date_str = _strip(pn.loc[mask].iloc[0, 1])
        except Exception:
            pass

    task_sheet = (
        _find_excel_sheet(xl, "Dados Consolidados")
        or _find_excel_sheet(xl, "Tarefas", "Tasks")
        or xl.sheet_names[0]
    )
    df = xl.parse(task_sheet, dtype=str)
    df.columns = [_strip(c) for c in df.columns]

    task_col = _find_column(df.columns, "Nome da tarefa", "Task name")
    if not task_col:
        raise ValueError(
            "Arquivo nao parece ser um export valido do Planner/Teams: "
            "coluna de tarefas nao encontrada."
        )

    def _replace_from_lookup(source_aliases, sheet_aliases, key_aliases, value_aliases):
        source_col = _find_column(df.columns, *source_aliases)
        lookup_sheet = _find_excel_sheet(xl, *sheet_aliases)
        if not source_col or not lookup_sheet:
            return
        try:
            lookup = xl.parse(lookup_sheet, dtype=str)
            lookup.columns = [_strip(c) for c in lookup.columns]
            key_col = _find_column(lookup.columns, *key_aliases)
            value_col = _find_column(lookup.columns, *value_aliases)
            if not key_col or not value_col:
                return
            mapping = {}
            for _, row in lookup.iterrows():
                key = _strip(row.get(key_col))
                value = _strip(row.get(value_col))
                if key and key.lower() != "nan" and value and value.lower() != "nan":
                    mapping[key] = value
            if not mapping:
                return

            def _map_cell(v):
                s = _strip(v)
                if not s or s.lower() == "nan":
                    return ""
                parts = [p.strip() for p in str(s).split(";") if p.strip()]
                if len(parts) > 1:
                    return "; ".join(mapping.get(p, p) for p in parts)
                return mapping.get(s, s)

            df[source_col] = df[source_col].map(_map_cell)
        except Exception:
            return

    _replace_from_lookup(
        ("Categoria", "Bucket name", "Nome do bucket"),
        ("Buckets",),
        ("ID de Bucket", "Bucket ID"),
        ("Nome do Bucket", "Bucket name"),
    )
    _replace_from_lookup(
        ("Atribuido a", "Assigned to"),
        ("Usuarios", "Users"),
        ("ID do Usuario", "User ID"),
        ("Nome do usuario", "User name"),
    )
    _replace_from_lookup(
        ("Criado por", "Created by"),
        ("Usuarios", "Users"),
        ("ID do Usuario", "User ID"),
        ("Nome do usuario", "User name"),
    )
    _replace_from_lookup(
        ("Concluida por", "Completed by"),
        ("Usuarios", "Users"),
        ("ID do Usuario", "User ID"),
        ("Nome do usuario", "User name"),
    )

    return df, plan_name, export_date_str


# ────────────────────────────── COMPUTE ──────────────────────────────────────
def compute_all(df_raw, plan_name, export_date_str=""):
    """Normalise column names, parse dates and return enriched DataFrame."""
    COL_MAP = {
        "nome da tarefa":                           "tarefa",
        "task name":                                "tarefa",
        "nome do bucket":                           "bucket",
        "bucket name":                              "bucket",
        "categoria":                                "bucket",
        "atribuído a":                              "assignee",
        "assigned to":                              "assignee",
        "concluído em":                             "data_conclusao",
        "date completed":                           "data_conclusao",
        "data de conclusão":                        "data_due",
        "data de vencimento":                       "data_due",
        "due date":                                 "data_due",
        "data de início":                           "data_inicio",
        "start date":                               "data_inicio",
        "data de criação":                          "data_criacao",
        "created date":                             "data_criacao",
        "data de criacao":                          "data_criacao",
        "criado em":                                "data_criacao",
        "progresso":                                "pct_done_raw",
        "status":                                   "pct_done_raw",
        "percentual concluído":                     "pct_done",
        "% concluída":                              "pct_done",
        "percent complete":                         "pct_done",
        "rótulos":                                  "labels",
        "labels":                                   "labels",
        "notas":                                    "notas",
        "descrição":                                "notas",
        "notes":                                    "notas",
        "itens da lista de verificação":            "checklist_items",
        "checklist items":                          "checklist_items",
        "itens concluídos da lista de verificação": "checklist_done",
        "completed checklist items":                "checklist_done",
    }
    cols_lower = {c.lower(): c for c in df_raw.columns}
    rename = {}
    for alias, canon in COL_MAP.items():
        if alias in cols_lower:
            rename[cols_lower[alias]] = canon
    df = df_raw.rename(columns=rename).copy()

    for c in ["tarefa","bucket","assignee","data_conclusao","data_inicio",
              "data_due","pct_done","pct_done_raw","labels","notas",
              "checklist_items","checklist_done","data_criacao"]:
        if c not in df.columns:
            df[c] = ""

    pct_raw = df["pct_done_raw"].fillna("").astype(str).str.strip()
    pct_num = pd.to_numeric(
        df["pct_done"].fillna("").astype(str).str.replace("%", "", regex=False),
        errors="coerce"
    ).fillna(0)

    date_done = _parse_date_series(df["data_conclusao"])
    date_start = _parse_date_series(df["data_inicio"])
    date_due = _parse_date_series(df["data_due"])
    date_criacao = _parse_date_series(df.get("data_criacao", pd.Series([""] * len(df))))

    pct_raw_norm = pct_raw.apply(
        lambda s: "".join(
            ch for ch in unicodedata.normalize("NFKD", s)
            if not unicodedata.combining(ch)
        ).lower().strip()
    )
    done_by_text = pct_raw_norm == "concluida"
    done_by_pct = pct_num >= 100
    done_by_date = date_done.notna()

    # Regra de neg?cio: "Conclu?da" ? definida por status de progresso.
    # Datas de conclus?o s?o usadas para timeline, mas n?o para reclassificar status.
    df["done"] = done_by_text | done_by_pct
    df["done_kpi"] = df["done"]
    df["done_by_date"] = done_by_date
    df["date_done"] = date_done
    df["date_start"] = date_start
    df["date_due"] = date_due
    df["date_criacao"] = date_criacao

    checklist_done = df["checklist_done"].fillna("").astype(str).str.strip()
    effort_from_completed = pd.to_numeric(
        checklist_done.str.extract(r"/\s*(\d+)")[0],
        errors="coerce",
    )
    checklist_items = df["checklist_items"].fillna("").astype(str).str.strip()
    effort_from_items = checklist_items.str.count(";") + checklist_items.ne("").astype(int)
    df["effort"] = effort_from_completed.fillna(effort_from_items).fillna(0).astype(int)

    bucket_norm = df["bucket"].fillna("").astype(str).str.strip()
    bucket_lower = bucket_norm.str.lower()
    labels_text = df["labels"].fillna("").astype(str)
    labels_lower = labels_text.str.lower()

    df["bucket_norm"] = bucket_norm
    df["is_nao_prev"] = (
        bucket_lower.str.contains(NAO_PREV_RE, na=False)
        | labels_lower.str.contains(NAO_PREV_RE, na=False)
    )
    df["is_backlog"] = bucket_lower.str.contains(BACKLOG_RE, na=False)
    df["hu"] = labels_text.map(_extract_hu_cached)

    assignees = df["assignee"].fillna("").astype(str).tolist()
    labels_for_area = labels_text.tolist()
    bucket_list = bucket_norm.tolist()
    df["area"] = [
        _area_for_bucket_labels_assignee(bkt, lbl, assignee)
        for bkt, lbl, assignee in zip(bucket_list, labels_for_area, assignees)
    ]

    all_dates = []
    for col in ["date_done","date_start","date_due"]:
        all_dates += df[col].dropna().tolist()

    sprint_start = min(all_dates) if all_dates else date.today()
    sprint_end   = max(all_dates) if all_dates else date.today()
    # Use export date from Nome do plano sheet if available, else today
    export_date  = _parse_date(export_date_str) or date.today()
    sprint_name  = plan_name or "Sprint"

    # Extract Sprint Goal text without removing the row from the dataset
    sprint_goal = ""
    mask = df["tarefa"].fillna("").astype(str).str.strip().str.lower() == "sprint goal"
    if mask.any():
        raw_goal = _strip(df.loc[mask, "notas"].iloc[0])
        # Strip common prefixes like "Sprint goal: " or "Sprint Goal: "
        sprint_goal = SPRINT_GOAL_RE.sub("", raw_goal).strip()

    return df, sprint_name, sprint_start, sprint_end, export_date, sprint_goal


def _parse_ignore_labels(ignore_labels):
    """
    Parse the ignore labels configuration.
    Accepts separators ';' or ',' and returns normalized label keys.
    """
    if ignore_labels is None:
        return set()

    if isinstance(ignore_labels, (list, tuple, set)):
        parts = []
        for item in ignore_labels:
            parts.extend(re.split(r"[;,]", str(item)))
    else:
        parts = re.split(r"[;,]", str(ignore_labels))

    return {_norm_label_key(p) for p in parts if str(p).strip()}


def _labels_has_ignored_token(labels_text, ignored_tokens):
    if not ignored_tokens:
        return False
    if labels_text is None:
        return False

    parts = [p.strip() for p in str(labels_text).split(";") if p.strip()]
    for p in parts:
        if _norm_label_key(p) in ignored_tokens:
            return True
    return False


def _apply_ignore_labels(df, ignore_labels=None):
    """
    Remove rows whose labels contain any ignored token.
    Returns (filtered_df, removed_count).
    """
    ignored_tokens = _parse_ignore_labels(ignore_labels)
    if not ignored_tokens:
        return df, 0

    mask_ignored = df["labels"].fillna("").astype(str).apply(
        lambda s: _labels_has_ignored_token(s, ignored_tokens)
    )
    removed = int(mask_ignored.sum())
    if removed <= 0:
        return df, 0

    return df.loc[~mask_ignored].copy(), removed


def _recompute_sprint_bounds(df, fallback_start, fallback_end):
    all_dates = []
    for col in ("date_done", "date_start", "date_due"):
        if col in df.columns:
            all_dates += df[col].dropna().tolist()

    if all_dates:
        return min(all_dates), max(all_dates)
    return fallback_start, fallback_end


# ─────────────────────────────── KPIs ────────────────────────────────────────
def build_kpis(df):
    # KPIs usam TODAS as tarefas (conforme Prompt Mestre: "Total = total de registros")
    total         = len(df)
    done          = int(df["done_kpi"].sum())
    pending       = total - done
    nao_prev      = int(df["is_nao_prev"].sum())
    backlog_total = int(df["is_backlog"].sum())
    sem_hu        = int((df["hu"] == "").sum())

    # Tarefas concluídas com datas válidas (para métricas de tempo)
    df_done = df[df["done_kpi"] & df["date_done"].notna()].copy()

    # CYCLE TIME TASK = média de (conclusão − início) para tarefas concluídas
    # Usa date_start (Data de Início)
    ct_task = None
    df_ct = df_done[df_done["date_start"].notna()].copy()
    if len(df_ct) > 0:
        delta_ct = [(d - s).days for d, s in zip(df_ct["date_done"].tolist(), df_ct["date_start"].tolist()) if d and s and (d - s).days >= 0]
        ct_task = round(sum(delta_ct) / len(delta_ct), 2) if delta_ct else None

    # LEAD TIME TASK = média de (conclusão − criação) para tarefas concluídas
    lt_task = None
    df_lt = df_done[df_done["date_criacao"].notna()].copy()
    if len(df_lt) > 0:
        delta_lt = []
        for d, c, s in zip(df_lt["date_done"].tolist(), df_lt["date_criacao"].tolist(), df_lt["date_start"].fillna(pd.NaT).tolist()):
            if d and c:
                # Se data início < criação, usa criação como base
                base = max(c, s) if (s and not pd.isna(s)) else c
                delta = (d - base).days
                if delta >= 0:
                    delta_lt.append(delta)
        lt_task = round(sum(delta_lt) / len(delta_lt), 2) if delta_lt else None

    # CYCLE TIME HU = média dos cycle times por HU, depois média das médias
    ct_hu = None
    lt_hu = None
    hu_groups = [(hu, grp) for hu, grp in df[df["hu"] != ""].groupby("hu")]
    if hu_groups:
        hu_ct_means = []
        hu_lt_means = []
        for _, grp in hu_groups:
            g_done = grp[grp["done_kpi"] & grp["date_done"].notna()]
            # CT por HU
            g_ct = g_done[g_done["date_start"].notna()]
            if len(g_ct) > 0:
                deltas = [(d - s).days for d, s in zip(g_ct["date_done"].tolist(), g_ct["date_start"].tolist()) if d and s and (d-s).days >= 0]
                if deltas:
                    hu_ct_means.append(sum(deltas) / len(deltas))
            # LT por HU
            g_lt = g_done[g_done["date_criacao"].notna()]
            if len(g_lt) > 0:
                deltas = []
                for d, c, s in zip(g_lt["date_done"].tolist(), g_lt["date_criacao"].tolist(), g_lt["date_start"].fillna(pd.NaT).tolist()):
                    if d and c:
                        base = max(c, s) if (s and not pd.isna(s)) else c
                        delta = (d - base).days
                        if delta >= 0:
                            deltas.append(delta)
                if deltas:
                    hu_lt_means.append(sum(deltas) / len(deltas))
        ct_hu = round(sum(hu_ct_means) / len(hu_ct_means), 2) if hu_ct_means else None
        lt_hu = round(sum(hu_lt_means) / len(hu_lt_means), 2) if hu_lt_means else None

    # Stakeholders = contagem de responsáveis distintos
    stakeholders = int(df["assignee"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique())

    pct_entrega = round(done / total, 4) if total else 0

    return {
        "total":         int(total),
        "done":          done,
        "pending":       pending,
        "nao_prev":      nao_prev,
        "sem_hu":        sem_hu,
        "backlog_total": backlog_total,
        "ct_task":       ct_task,
        "lt_task":       lt_task,
        "ct_hu":         ct_hu,
        "lt_hu":         lt_hu,
        "pct_entrega":   pct_entrega,
        "stakeholders":  stakeholders,
    }


# ─────────────────────────────── HU LIST + NAMES ─────────────────────────────
def build_hu_list(df):
    """Returns [(hu_id, total, done)] sorted by hu_id."""
    dfw = df[df["hu"] != ""].copy()
    rows = []
    for hu, grp in dfw.groupby("hu"):
        rows.append((hu, len(grp), int(grp["done"].sum())))
    rows.sort(key=lambda x: x[0])
    return rows


def build_hu_full_names(df):
    """
    Map HU short IDs (e.g. 'HU044') to full label strings
    (e.g. 'HU044 - Tela de equipe SquadsRS') as found in the Planner export labels.
    """
    hu_names = {}
    df_hu = df[df["hu"] != ""]

    for hu_id, labels in zip(
        df_hu["hu"].tolist(),
        df_hu["labels"].fillna("").astype(str).tolist(),
    ):
        if hu_id in hu_names:
            continue

        full_name = _extract_hu_full_label(labels)
        hu_names[hu_id] = full_name or hu_id

    return hu_names


# ─────────────────────────── POR COLABORADOR ─────────────────────────────────
def build_por_colaborador(df):
    rows = []
    for person, grp in df.groupby("assignee"):
        name = _strip(person) or "Sem atribuição"
        done = int(grp["done"].sum())
        pend = len(grp) - done
        rows.append((name, done, pend))
    rows.sort(key=lambda x: -(x[1]+x[2]))
    return rows


# ─────────────────────────────── AREAS ───────────────────────────────────────
def build_areas(df):
    """Área = rótulos que começam com '.' (ponto), conforme novo prompt."""
    from collections import defaultdict
    area_counts = defaultdict(int)
    for label_text in df["labels"].fillna("").astype(str).tolist():
        for lbl in label_text.split(";"):
            lbl = lbl.strip()
            if lbl.startswith("."):
                area_counts[lbl] += 1
    rows = [(area, count, 0) for area, count in area_counts.items()]
    rows.sort(key=lambda x: -x[1])
    return rows


# ─────────────────────────────── HU IN/OUT ───────────────────────────────────
def build_hu_in_out(df):
    dfw = df[~df["is_backlog"]].copy()
    in_hu  = int((dfw["hu"] != "").sum())
    out_hu = int((dfw["hu"] == "").sum())
    return [("Em HU", in_hu), ("Fora de HU", out_hu)]


# ─────────────────────────────── POR CATEGORIA ───────────────────────────────
def build_por_categoria(df):
    """Returns (cat_rows, bubble_rows)."""
    cat_counts = defaultdict(lambda: {"done": 0, "total": 0, "effort": 0, "label": ""})

    labels = df["labels"].fillna("").astype(str).tolist()
    done_vals = df["done"].astype(int).tolist()
    efforts = df["effort"].fillna(0).astype(int).tolist()

    for label_text, done, effort in zip(labels, done_vals, efforts):
        for cat in _categories_from_labels(label_text):
            key = _norm_label_key(cat)
            if not key:
                continue
            if not cat_counts[key]["label"]:
                cat_counts[key]["label"] = cat.strip()
            cat_counts[key]["total"] += 1
            cat_counts[key]["done"] += done
            cat_counts[key]["effort"] += effort

    PINNED_LAST = {"bugs e ajustes"}  # categorias sempre exibidas por último

    cat_rows = []
    bub_rows = []
    pinned_cat = []
    pinned_bub = []
    for _key, vals in sorted(cat_counts.items(), key=lambda x: -x[1]["total"]):
        cat = vals["label"]
        done = vals["done"]
        total = vals["total"]
        effort = vals["effort"]
        pend = total - done
        if _norm_label_key(cat) in PINNED_LAST:
            pinned_cat.append((cat, done, pend))
            pinned_bub.append((cat, done, total, effort))
        else:
            cat_rows.append((cat, done, pend))
            bub_rows.append((cat, done, total, effort))

    # Itens fixados sempre ao final (ex: BUGs e Ajustes)
    cat_rows.extend(pinned_cat)
    bub_rows.extend(pinned_bub)

    return cat_rows, bub_rows


# ─────────────────────────────── ESFORÇO HU ──────────────────────────────────
def build_esforco_hu_detailed(df):
    """
    Returns [(hu_id, max_effort_per_person, total_effort)] sorted by total desc.
    max_effort_per_person = max effort contributed by any single assignee for this HU.
    """
    rows = []
    for hu, grp in df[df["hu"] != ""].groupby("hu"):
        total_effort = int(grp["effort"].sum())
        per_person   = grp.groupby("assignee")["effort"].sum()
        max_ep       = int(per_person.max()) if not per_person.empty else 0
        rows.append((hu, max_ep, total_effort))
    rows.sort(key=lambda x: -x[2])
    return rows


# ─────────────────────────── BUILD RÓTULOS ───────────────────────────────────
def build_rotulos(df):
    """
    Returns list of (label_display, done, pending, lead_time, cycle_time) sorted by total desc.
    LeadTime = conclusão − início  (mean over done tasks with that label)
    CycleTime = conclusão − criação (mean over done tasks with that label)
    Label display = "Rótulo (done/total)"
    """
    from collections import defaultdict
    label_done   = defaultdict(int)
    label_total  = defaultdict(int)
    label_lt     = defaultdict(list)
    label_ct     = defaultdict(list)

    for _, row in df.iterrows():
        lbls = [l.strip() for l in str(row.get("labels", "")).split(";") if l.strip()]
        for lbl in lbls:
            key = lbl.strip()
            if not key:
                continue
            label_total[key] += 1
            if row.get("done_kpi"):
                label_done[key] += 1
                # Lead time: conclusão - início
                d_done  = row.get("date_done")
                d_start = row.get("date_start")
                d_criac = row.get("date_criacao")
                if d_done and pd.notna(d_done) and d_start and pd.notna(d_start):
                    lt = (d_done - d_start).days
                    if lt >= 0:
                        label_lt[key].append(lt)
                # Cycle time: conclusão - criação
                if d_done and pd.notna(d_done) and d_criac and pd.notna(d_criac):
                    ct = (d_done - d_criac).days
                    if ct >= 0:
                        label_ct[key].append(ct)

    rows = []
    for lbl in label_total:
        done  = label_done[lbl]
        total = label_total[lbl]
        pend  = total - done
        lt    = round(sum(label_lt[lbl]) / len(label_lt[lbl]), 2) if label_lt[lbl] else 0
        ct    = round(sum(label_ct[lbl]) / len(label_ct[lbl]), 2) if label_ct[lbl] else 0
        display = f"{lbl} ({done}/{total})"
        rows.append((display, done, pend, lt, ct))
    rows.sort(key=lambda x: -(x[1]+x[2]))
    return rows


# ─────────────────────── BUILD RESPONSÁVEIS ──────────────────────────────────
def build_responsaveis(df):
    """
    Returns list of (name_display, done_bucket, pending, lead_time, cycle_time) sorted by total desc.
    name_display = "Nome (done/total)"
    """
    rows = []
    for person, grp in df.groupby("assignee"):
        name = _strip(person) or "Sem atribuição"
        done_bucket = int(grp["done_kpi"].sum())
        pend        = len(grp) - done_bucket
        total       = len(grp)

        grp_done = grp[grp["done_kpi"] & grp["date_done"].notna()]
        # Lead time
        lt_vals = []
        for _, row in grp_done.iterrows():
            d_done  = row.get("date_done")
            d_start = row.get("date_start")
            if d_done and pd.notna(d_done) and d_start and pd.notna(d_start):
                lt = (d_done - d_start).days
                if lt >= 0:
                    lt_vals.append(lt)
        # Cycle time
        ct_vals = []
        for _, row in grp_done.iterrows():
            d_done  = row.get("date_done")
            d_criac = row.get("date_criacao")
            if d_done and pd.notna(d_done) and d_criac and pd.notna(d_criac):
                ct = (d_done - d_criac).days
                if ct >= 0:
                    ct_vals.append(ct)

        lt = round(sum(lt_vals) / len(lt_vals), 2) if lt_vals else 0
        ct = round(sum(ct_vals) / len(ct_vals), 2) if ct_vals else 0
        display = f"{name} ({done_bucket}/{total})"
        rows.append((display, done_bucket, pend, lt, ct))
    rows.sort(key=lambda x: -(x[1]+x[2]))
    return rows


# ──────────────────────────── BUILD CFD ──────────────────────────────────────
def build_cfd(df, sprint_start, export_date):
    """
    Cumulative Flow Diagram — daily counts for full month.
    Returns (dates, todo_list, doing_list, done_list)
    - To Do:   tasks where date_start is None or > day
    - Doing:   tasks where date_start <= day < date_done (or export if no date_done)
    - Done:    cumulative tasks with date_done <= day
    Carry-forward after export_date.
    """
    bd_start, bd_end = _month_bounds(sprint_start)
    days = (bd_end - bd_start).days + 1
    cutoff = export_date if isinstance(export_date, date) else export_date

    todo_list, doing_list, done_list = [], [], []
    last_todo = last_doing = last_done = None

    for i in range(days):
        d = bd_start + timedelta(days=i)
        if d <= cutoff:
            todo  = 0
            doing = 0
            done  = 0
            for _, row in df.iterrows():
                d_start = row.get("date_start")
                d_done  = row.get("date_done")
                # Done: cumulative
                if d_done and d_done <= d:
                    done += 1
                elif d_start and d_start <= d:
                    doing += 1
                else:
                    todo += 1
            last_todo, last_doing, last_done = todo, doing, done
        else:
            todo, doing, done = last_todo, last_doing, last_done
        todo_list.append(todo)
        doing_list.append(doing)
        done_list.append(done)

    dates = [datetime(bd_start.year, bd_start.month, (bd_start + timedelta(days=i)).day) for i in range(days)]
    return dates, todo_list, doing_list, done_list


# ──────────────────────────── BUILD WIP ──────────────────────────────────────
def build_wip(df, sprint_start, export_date):
    """
    WIP by bucket — daily counts for full month.
    Returns (dates, bucket_names, bucket_matrix)
    bucket_matrix[bucket_idx][day_idx] = count
    'Concluído' bucket is cumulative (growing). Others use interval [start, done).
    Carry-forward after export_date.
    """
    bd_start, bd_end = _month_bounds(sprint_start)
    days = (bd_end - bd_start).days + 1
    cutoff = export_date if isinstance(export_date, date) else export_date

    # Identifica buckets únicos (exceto backlog)
    all_buckets = df[~df["is_backlog"]]["bucket_norm"].fillna("").unique().tolist()
    # Ordena: Concluído primeiro, depois por frequência decrescente
    bucket_counts = df[~df["is_backlog"]]["bucket_norm"].value_counts().to_dict()
    all_buckets = sorted(all_buckets, key=lambda b: (
        0 if "conclu" in b.lower() else 1,
        -bucket_counts.get(b, 0)
    ))

    bucket_matrix = {b: [0]*days for b in all_buckets}
    last_state = {b: 0 for b in all_buckets}

    for i in range(days):
        d = bd_start + timedelta(days=i)
        if d <= cutoff:
            for bkt in all_buckets:
                count = 0
                is_concluido = "conclu" in bkt.lower()
                for _, row in df[~df["is_backlog"]].iterrows():
                    if row["bucket_norm"] != bkt:
                        continue
                    d_start = row.get("date_start")
                    d_done  = row.get("date_done")
                    if is_concluido:
                        # Cumulative: count if done on or before day
                        if d_done and d_done <= d:
                            count += 1
                    else:
                        # Interval: [start, done)
                        end = d_done if d_done else cutoff
                        if d_start and d_start <= d < end:
                            count += 1
                        elif not d_start and d <= (end or cutoff):
                            pass  # sem data início, não conta
                last_state[bkt] = count
                bucket_matrix[bkt][i] = count
        else:
            for bkt in all_buckets:
                bucket_matrix[bkt][i] = last_state[bkt]

    dates = [datetime(bd_start.year, bd_start.month, (bd_start + timedelta(days=i)).day) for i in range(days)]
    x_total = [sum(bucket_matrix[b][i] for b in all_buckets) for i in range(days)]
    return dates, all_buckets, bucket_matrix, x_total


def build_wip_prompt3(df, sprint_start, export_date):
    """
    WIP diario por rotulo (31 dias), com conservacao de inventario.

    Regras:
    - X_total = soma total de associacoes (rotulos) da base.
    - Colunas dinamicas: Backlog + todos os rotulos unicos + Concluido.
    - Backlog: soma X se Data_Inicio > Data_Ref (ou sem data_inicio).
    - WIP (rotulos): soma 1 por rotulo se
      Data_Inicio <= Data_Ref e (Data_Conclusao > Data_Ref ou nula).
    - Concluido: soma X se Data_Conclusao <= Data_Ref.
    - No dia da conclusao, sai do WIP e entra em Concluido.
    - Carry-forward apos data de exportacao.

    Retorna (dates, columns, matrix, x_total_rows)
      matrix[col_name][day_idx] = count
      x_total_rows[day_idx] = X_total (constante)
    """
    bd_start, bd_end = _month_bounds(sprint_start)
    days = (bd_end - bd_start).days + 1
    cutoff = export_date if isinstance(export_date, date) else bd_end
    if cutoff < bd_start:
        cutoff = bd_start
    if cutoff > bd_end:
        cutoff = bd_end

    # Mapeia todos os rotulos unicos (case-insensitive), preservando display.
    label_display = {}
    for labels_text in df["labels"].fillna("").astype(str).tolist():
        for token in _categories_from_labels(labels_text):
            key = _norm_label_key(token)
            if key and key not in label_display:
                label_display[key] = token.strip()

    label_keys = sorted(label_display.keys(), key=lambda k: label_display[k].lower())
    columns = ["Backlog"] + [label_display[k] for k in label_keys] + ["Concluido"]
    matrix = {c: [0] * days for c in columns}

    # Preprocessa tarefas com X = qtd de associacoes (rotulos unicos da tarefa).
    tasks = []
    for _, row in df.iterrows():
        seen = set()
        row_keys = []
        for token in _categories_from_labels(row.get("labels", "")):
            key = _norm_label_key(token)
            if key and key not in seen:
                seen.add(key)
                row_keys.append(key)

        x = len(row_keys)
        if x <= 0:
            continue

        d_start = row.get("date_start")
        d_done = row.get("date_done")
        if pd.isna(d_start):
            d_start = None
        if pd.isna(d_done):
            d_done = None
        tasks.append((row_keys, x, d_start, d_done))

    x_total_value = sum(x for _, x, _, _ in tasks)
    x_total_rows = [x_total_value] * days

    for i in range(days):
        d = bd_start + timedelta(days=i)

        # Repeticao apos exportacao (carry-forward).
        if i > 0 and d > cutoff:
            for c in columns:
                matrix[c][i] = matrix[c][i - 1]
            continue

        backlog = 0
        done = 0
        per_label = {k: 0 for k in label_keys}

        for row_keys, x, d_start, d_done in tasks:
            done_now = d_done is not None and d_done <= d
            started = d_start is not None and d_start <= d

            if done_now:
                done += x
                continue

            if started:
                # WIP: soma 1 por rotulo da tarefa.
                for k in row_keys:
                    per_label[k] += 1
            else:
                # Backlog: soma X da tarefa inteira.
                backlog += x

        matrix["Backlog"][i] = backlog
        for k in label_keys:
            matrix[label_display[k]][i] = per_label[k]
        matrix["Concluido"][i] = done

        # Conservacao de inventario: soma horizontal deve fechar em X_total.
        row_sum = backlog + done + sum(per_label.values())
        if row_sum != x_total_value:
            matrix["Backlog"][i] += (x_total_value - row_sum)

    dates = [
        datetime(bd_start.year, bd_start.month, (bd_start + timedelta(days=i)).day)
        for i in range(days)
    ]
    return dates, columns, matrix, x_total_rows


# ──────────────────────────── BUILD CTS ──────────────────────────────────────
def build_cts(df, sprint_start):
    """
    Completions per Time & Subject — matrix days × HU.
    Returns (dates, hu_labels, cts_matrix, fora_hu_daily)
    cts_matrix[hu_idx][day_idx] = count of tasks completed that day for that HU
    fora_hu_daily[day_idx] = count of tasks without HU completed that day
    """
    bd_start, bd_end = _month_bounds(sprint_start)
    days = (bd_end - bd_start).days + 1

    def _all_hus(labels_text):
        out = []
        seen = set()
        for token in [p.strip() for p in str(labels_text).split(";") if p.strip()]:
            m = re.search(r"\b(HU\s*0*\d+)\b", token, flags=re.IGNORECASE)
            if not m:
                continue
            num = re.sub(r"\D", "", m.group(1))
            hu = f"HU{num.zfill(3)}"
            if hu not in seen:
                seen.add(hu)
                out.append(hu)
        return out

    hu_set = set()
    hu_full = {}
    for _, row in df.iterrows():
        labels_text = str(row.get("labels", ""))
        full = _extract_hu_full_label(labels_text)
        for hu in _all_hus(labels_text):
            hu_set.add(hu)
            if hu not in hu_full:
                hu_full[hu] = full or hu
    hu_list = sorted(hu_set)

    cts_matrix   = [[0]*days for _ in hu_list]
    fora_hu_daily = [0]*days

    done_rows = df[df["done_kpi"] & df["date_done"].notna()]
    for _, row in done_rows.iterrows():
        d = row["date_done"]
        offset = (d - bd_start).days
        if 0 <= offset < days:
            hu = row["hu"]
            if hu and hu in hu_list:
                idx = hu_list.index(hu)
                cts_matrix[idx][offset] += 1
            else:
                fora_hu_daily[offset] += 1

    dates = [datetime(bd_start.year, bd_start.month, (bd_start + timedelta(days=i)).day) for i in range(days)]
    return dates, [hu_full.get(h, h) for h in hu_list], cts_matrix, fora_hu_daily


# ──────────────────────── BUILD HISTOGRAMA ───────────────────────────────────
def build_histograma(df):
    """
    Histograma de cycle time (conclusão − início) para tarefas concluídas.
    Returns (hist_31, indicativos)
    hist_31 = list of 31 values, index 0 = cycle time of 1 day
    indicativos = list of 31 values (0.5, 0.75, 0.9 at percentile positions, else None)
    """
    # Cycle time = conclusão − início (Lead Time)
    done_rows = df[df["done_kpi"] & df["date_done"].notna() & df["date_start"].notna()]
    ct_vals = []
    for _, row in done_rows.iterrows():
        delta = (row["date_done"] - row["date_start"]).days
        if delta >= 0:
            ct_vals.append(max(delta, 1))  # mínimo 1 dia

    hist_31 = [0] * 31
    for v in ct_vals:
        if 1 <= v <= 31:
            hist_31[v-1] += 1

    # Percentis (50%, 75%, 90%) baseados na distribuição cumulativa
    total = len(ct_vals)
    indicativos = [None] * 31
    if total > 0:
        cumsum = 0
        found = set()
        for i, count in enumerate(hist_31):
            cumsum += count
            pct = cumsum / total
            if 0.5 not in found and pct >= 0.5:
                indicativos[i] = 0.5
                found.add(0.5)
            elif 0.75 not in found and pct >= 0.75:
                indicativos[i] = 0.75
                found.add(0.75)
            elif 0.9 not in found and pct >= 0.9:
                indicativos[i] = 0.9
                found.add(0.9)

    return hist_31, indicativos


# ─────────────────────── BUILD HISTOGRAMA2 ───────────────────────────────────
def build_histograma2(df):
    """
    Estatísticas do histograma de tempos.
    Retorna lista de (descricao, valor) em ordem crescente de valor.
    Lead Time = conclusão − início
    Cycle Time = conclusão − criação
    """
    done_rows = df[df["done_kpi"] & df["date_done"].notna()]

    # Lead time values (conclusão − início)
    lt_vals = []
    for _, row in done_rows[done_rows["date_start"].notna()].iterrows():
        delta = (row["date_done"] - row["date_start"]).days
        if delta >= 0:
            lt_vals.append(max(delta, 1))

    # Cycle time values (conclusão − criação)
    ct_vals = []
    for _, row in done_rows[done_rows["date_criacao"].notna()].iterrows():
        d_criac = row["date_criacao"]
        d_start = row.get("date_start")
        d_done  = row["date_done"]
        # Se início < criação, usa criação
        base = max(d_criac, d_start) if d_start and not pd.isna(d_start) else d_criac
        delta = (d_done - base).days
        if delta >= 0:
            ct_vals.append(max(delta, 1))

    if not lt_vals:
        return []

    lt_sorted = sorted(lt_vals)
    n = len(lt_sorted)

    def median(vals):
        s = sorted(vals); m = len(s)
        return (s[m//2-1] + s[m//2]) / 2 if m % 2 == 0 else s[m//2]

    def percentil(vals, p):
        s = sorted(vals); idx = int(p * len(s))
        return s[min(idx, len(s)-1)]

    from statistics import mode as stat_mode
    try:
        moda = stat_mode(lt_vals)
    except Exception:
        moda = lt_vals[0] if lt_vals else 0

    med  = round(median(lt_sorted), 2)
    lt_m = round(sum(lt_sorted) / n, 2)
    ct_m = round(sum(ct_vals) / len(ct_vals), 2) if ct_vals else 0
    p75  = round(percentil(lt_sorted, 0.75), 2)
    p90  = round(percentil(lt_sorted, 0.90), 2)
    amp  = lt_sorted[-1] - lt_sorted[0] if lt_sorted else 0

    return [
        ("Moda: Tempo de ciclo mais frequente", moda),
        (f"Mediana (50% das entregas): concluídas neste prazo (ritmo padrão)", med),
        (f"Lead Time Médio: Tempo médio geral de ciclo", lt_m),
        (f"Cycle Time Médio: Tempo médio entre criação e fim da atividade", ct_m),
        (f"Percentil 75% (Corte de Maioria): 3/4 das demandas finalizadas neste prazo", p75),
        (f"SLE - Service Level Expectation (90%): Limite de confiança (90% de certeza neste prazo)", p90),
        (f"Amplitude de Desvio (Outlier): Diferença entre mais rápida [{lt_sorted[0]} dias] e a mais lenta [{lt_sorted[-1]} dias]", amp),
    ]


# ─────────────────────────── DISPERSÃO DIÁRIA ────────────────────────────────
def _month_bounds(ref_date):
    """Return (first_day, last_day) of the month of ref_date."""
    first = date(ref_date.year, ref_date.month, 1)
    last_day = calendar.monthrange(ref_date.year, ref_date.month)[1]
    last  = date(ref_date.year, ref_date.month, last_day)
    return first, last


def build_dispersao_daily(df, sprint_start, sprint_end):
    """
    Build daily task completion counts for Dispersao chart.
    X-axis always covers the full month of sprint_end (1st to last day).
    Returns (hu_list, days, hu_matrix, nao_hu_daily).
      hu_list      : sorted list of HU short IDs
      days         : number of days in the sprint month (28-31)
      hu_matrix    : list[hu_idx][day_offset] = simple task count that day
      nao_hu_daily : list[day_offset] = task count for tasks without HU
    """
    # Snap to full month of sprint_end
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1  # always 28-31

    hu_list = sorted(df[df["hu"] != ""]["hu"].unique().tolist())
    hu_index = {hu: idx for idx, hu in enumerate(hu_list)}

    # Valores float: cada tarefa concluída contribui 1.0
    hu_matrix = [[0.0] * days for _ in hu_list]
    nao_hu_daily = [0.0] * days

    done_rows = df[df["done"] & df["date_done"].notna()][["hu", "date_done"]]
    for hu, d in zip(done_rows["hu"].tolist(), done_rows["date_done"].tolist()):
        offset = (d - bd_start).days
        if 0 <= offset < days:
            idx = hu_index.get(hu)
            if idx is None:
                nao_hu_daily[offset] += 1.0
            else:
                hu_matrix[idx][offset] += 1.0

    return hu_list, days, hu_matrix, nao_hu_daily


# ═══════════════════════════════════════════════════════════════════════════════
#                        TEMPLATE FILLING FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def _is_valid_date(v):
    return v is not None and not (hasattr(v, "_typ") or str(v) == "NaT")


def _done_plan_value(row):
    d = row.get("date_due")
    if _is_valid_date(d):
        return d
    return row.get("date_done")


def _wip_phase_from_row(row):
    bucket = _lower(row.get("bucket_norm", ""))
    labels = _lower(row.get("labels", ""))
    tokens = [t.strip() for t in str(labels).split(";") if t.strip()]
    has_hu = any(t.startswith("hu") for t in tokens)
    has_gp = any(t in (".gp", "gp") for t in tokens)
    has_ux = any(t in (".ux", "ux") for t in tokens)
    has_po = any(t in (".po", "po") for t in tokens)
    has_req = any(t in (".requisitos", "requisitos") for t in tokens)

    if "gestao" in bucket or "gestão" in bucket:
        return "Gestão"
    # Item de gestão sem HU (ex.: GP/PO/servant) deve ir para Gestão.
    if has_gp and not has_hu:
        return "Gestão"
    # Itens de UX com PO/REQUISITOS são tratados como refinamento.
    if has_ux and (has_po or has_req):
        return "Em Refinamento"
    if any(t.startswith(".dev") or "dev." in t for t in tokens):
        return "Em Desenvolvimento"
    if any(
        t.startswith(".ux")
        or t.startswith(".desktop")
        or t.startswith(".mobile")
        or t.startswith(".figma")
        for t in tokens
    ):
        return "UX/UI"
    return "Em Refinamento"


def _build_step_plan(total, days, blocks=5):
    import math

    if days <= 0:
        return []
    blocks = max(1, min(blocks, days))

    offsets = []
    for i in range(blocks):
        off = int(math.ceil((i + 1) * days / blocks) - 1)
        off = max(0, min(days - 1, off))
        if not offsets or off != offsets[-1]:
            offsets.append(off)
    if offsets[-1] != days - 1:
        offsets[-1] = days - 1

    targets = []
    for i in range(len(offsets)):
        t = round(total * (1 - (i + 1) / len(offsets)))
        targets.append(max(t, 0))
    targets[-1] = 0

    plan = []
    cur = total
    k = 0
    for i in range(days):
        while k < len(offsets) and i >= offsets[k]:
            cur = targets[k]
            k += 1
        plan.append(cur)
    return plan


def build_kpis(df):
    total = len(df)
    done = int(df["done_kpi"].sum())
    pending = total - done
    nao_prev = int(df["is_nao_prev"].sum())
    backlog_total = int(df["is_backlog"].sum())
    def _has_hu_label(labels_text):
        for token in [p.strip() for p in str(labels_text).split(";") if p.strip()]:
            if re.search(r"\bHU\s*0*\d+\b", token, flags=re.IGNORECASE):
                return True
        return False

    sem_hu = int((~df["labels"].fillna("").astype(str).apply(_has_hu_label)).sum())
    fluxo_continuo = _count_rows_with_label(df, "FLUXO.CONTINUO")

    df_done = df[df["done_kpi"] & df["date_done"].notna()].copy()

    ct_task = None
    df_ct = df_done[df_done["date_start"].notna()].copy()
    if len(df_ct) > 0:
        deltas = [
            (d - s).days
            for d, s in zip(df_ct["date_done"].tolist(), df_ct["date_start"].tolist())
            if _is_valid_date(d) and _is_valid_date(s) and (d - s).days >= 0
        ]
        ct_task = round(sum(deltas) / len(deltas), 2) if deltas else None

    lt_task = None
    df_lt = df_done[df_done["date_criacao"].notna()].copy()
    if len(df_lt) > 0:
        deltas = []
        for d, c in zip(df_lt["date_done"].tolist(), df_lt["date_criacao"].tolist()):
            if _is_valid_date(d) and _is_valid_date(c):
                delta = (d - c).days
                if delta >= 0:
                    deltas.append(delta)
        lt_task = round(sum(deltas) / len(deltas), 2) if deltas else None

    ct_hu = None
    lt_hu = None
    hu_groups = [(hu, grp) for hu, grp in df[df["hu"] != ""].groupby("hu")]
    if hu_groups:
        hu_ct_means = []
        hu_lt_means = []
        for _, grp in hu_groups:
            g_done = grp[grp["done_kpi"] & grp["date_done"].notna()]

            g_ct = g_done[g_done["date_start"].notna()]
            if len(g_ct) > 0:
                deltas = [
                    (d - s).days
                    for d, s in zip(g_ct["date_done"].tolist(), g_ct["date_start"].tolist())
                    if _is_valid_date(d) and _is_valid_date(s) and (d - s).days >= 0
                ]
                if deltas:
                    hu_ct_means.append(sum(deltas) / len(deltas))

            g_lt = g_done[g_done["date_criacao"].notna()]
            if len(g_lt) > 0:
                deltas = []
                for d, c in zip(g_lt["date_done"].tolist(), g_lt["date_criacao"].tolist()):
                    if _is_valid_date(d) and _is_valid_date(c):
                        delta = (d - c).days
                        if delta >= 0:
                            deltas.append(delta)
                if deltas:
                    hu_lt_means.append(sum(deltas) / len(deltas))

        ct_hu = round(sum(hu_ct_means) / len(hu_ct_means), 2) if hu_ct_means else None
        lt_hu = round(sum(hu_lt_means) / len(hu_lt_means), 2) if hu_lt_means else None

    stakeholders = int(
        df["assignee"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()
    )
    pct_entrega = round(done / total, 4) if total else 0

    return {
        "total": int(total),
        "done": done,
        "pending": pending,
        "nao_prev": nao_prev,
        "sem_hu": sem_hu,
        "backlog_total": backlog_total,
        "ct_task": ct_task,
        "lt_task": lt_task,
        "ct_hu": ct_hu,
        "lt_hu": lt_hu,
        "ct_sp": ct_hu,
        "lt_sp": lt_hu,
        "storypoints": 0,
        "fluxo_continuo": fluxo_continuo,
        "pct_entrega": pct_entrega,
        "stakeholders": stakeholders,
    }


def _enrich_kpis_with_hu_storypoints(kpis, df, hu_storypoints):
    total_sp = round(sum(v for v in hu_storypoints.values() if v), 2)
    kpis["storypoints"] = total_sp

    hu_ct_means = []
    hu_lt_means = []
    weighted_ct_num = 0.0
    weighted_ct_den = 0.0
    weighted_lt_num = 0.0
    weighted_lt_den = 0.0

    for hu, grp in df[df["hu"] != ""].groupby("hu"):
        g_done = grp[grp["done_kpi"]].copy()
        if g_done.empty:
            continue

        ct_vals = []
        for _, row in g_done.iterrows():
            d_done = _done_plan_value(row)
            d_start = row.get("date_start")
            if _is_valid_date(d_done) and _is_valid_date(d_start):
                delta = (d_done - d_start).days
                if delta >= 0:
                    ct_vals.append(delta)
        if ct_vals:
            ct_mean = sum(ct_vals) / len(ct_vals)
            hu_ct_means.append(ct_mean)
            sp = hu_storypoints.get(str(hu).upper(), 0) or 0
            if sp > 0:
                weighted_ct_num += ct_mean * sp
                weighted_ct_den += sp

        lt_vals = []
        for _, row in g_done.iterrows():
            d_done = _done_plan_value(row)
            d_cr = row.get("date_criacao")
            if _is_valid_date(d_done) and _is_valid_date(d_cr):
                delta = (d_done - d_cr).days
                if delta >= 0:
                    lt_vals.append(delta)
        if lt_vals:
            lt_mean = sum(lt_vals) / len(lt_vals)
            hu_lt_means.append(lt_mean)
            sp = hu_storypoints.get(str(hu).upper(), 0) or 0
            if sp > 0:
                weighted_lt_num += lt_mean * sp
                weighted_lt_den += sp

    if hu_ct_means:
        kpis["ct_hu"] = round(sum(hu_ct_means) / len(hu_ct_means), 2)
    if hu_lt_means:
        kpis["lt_hu"] = round(sum(hu_lt_means) / len(hu_lt_means), 2)
    kpis["ct_sp"] = round(weighted_ct_num / weighted_ct_den, 2) if weighted_ct_den > 0 else kpis.get("ct_hu", 0)
    kpis["lt_sp"] = round(weighted_lt_num / weighted_lt_den, 2) if weighted_lt_den > 0 else kpis.get("lt_hu", 0)


def build_areas(df):
    """
    Areas para o gráfico de proporção:
    - considera rótulos com estrutura de área (contendo ponto), exclui HU* e FLUXO.CONTINUO
    - normaliza para prefixo com "." no output (ex.: SERVANT.LEADER -> .SERVANT.LEADER)
    """
    from collections import defaultdict

    counts = defaultdict(int)
    for label_text in df["labels"].fillna("").astype(str).tolist():
        for raw in [p.strip() for p in str(label_text).split(";") if p.strip()]:
            key = _norm_label_key(raw)
            if not key:
                continue
            if key.startswith("hu"):
                continue
            if key == "fluxo.continuo":
                continue
            if "." not in key:
                continue

            display = raw.strip()
            if not display.startswith("."):
                display = f".{display}"
            counts[display.upper()] += 1

    rows = [(area, cnt, 0) for area, cnt in counts.items() if cnt > 0]
    rows.sort(key=lambda x: -x[1])
    return rows


def build_rotulos(df):
    from collections import defaultdict

    label_done = defaultdict(int)
    label_total = defaultdict(int)
    label_lt = defaultdict(list)
    label_ct = defaultdict(list)
    label_display = {}

    for _, row in df.iterrows():
        lbls = [l.strip() for l in str(row.get("labels", "")).split(";") if l.strip()]
        for lbl in lbls:
            key = _norm_label_key(lbl)
            if not key:
                continue
            if key not in label_display:
                label_display[key] = lbl.strip()
            label_total[key] += 1
            if row.get("done_kpi"):
                label_done[key] += 1
                d_done = row.get("date_done")
                d_start = row.get("date_start")
                d_criac = row.get("date_criacao")
                if _is_valid_date(d_done) and _is_valid_date(d_criac):
                    lt = (d_done - d_criac).days
                    if lt >= 0:
                        label_lt[key].append(lt)
                if _is_valid_date(d_done) and _is_valid_date(d_start):
                    ct = (d_done - d_start).days
                    if ct >= 0:
                        label_ct[key].append(ct)

    rows = []
    for key in label_total:
        done = label_done[key]
        total = label_total[key]
        pend = total - done
        lt = round(sum(label_lt[key]) / len(label_lt[key]), 2) if label_lt[key] else 0
        ct = round(sum(label_ct[key]) / len(label_ct[key]), 2) if label_ct[key] else 0
        display = f"{label_display[key]} ({done}/{total})"
        rows.append((display, done, pend, lt, ct))
    rows.sort(key=lambda x: -(x[1] + x[2]))
    return rows


def build_responsaveis(df):
    rows = []
    for person, grp in df.groupby("assignee"):
        name = _strip(person) or "Sem atribuição"
        done_bucket = int(grp["done_kpi"].sum())
        pend = len(grp) - done_bucket
        total = len(grp)

        grp_done = grp[grp["done_kpi"] & grp["date_done"].notna()]
        lt_vals = []
        ct_vals = []
        for _, row in grp_done.iterrows():
            d_done = row.get("date_done")
            d_start = row.get("date_start")
            d_criac = row.get("date_criacao")
            if _is_valid_date(d_done) and _is_valid_date(d_criac):
                lt = (d_done - d_criac).days
                if lt >= 0:
                    lt_vals.append(lt)
            if _is_valid_date(d_done) and _is_valid_date(d_start):
                ct = (d_done - d_start).days
                if ct >= 0:
                    ct_vals.append(ct)

        lt = round(sum(lt_vals) / len(lt_vals), 2) if lt_vals else 0
        ct = round(sum(ct_vals) / len(ct_vals), 2) if ct_vals else 0
        display = f"{name} ({done_bucket}/{total})"
        rows.append((display, done_bucket, pend, lt, ct))

    rows.sort(key=lambda x: -(x[1] + x[2]))
    return rows


def build_cfd(df, sprint_start, export_date):
    bd_start, bd_end = _month_bounds(sprint_start)
    days = (bd_end - bd_start).days + 1
    cutoff = export_date if isinstance(export_date, date) else export_date

    todo_list, doing_list, done_list = [], [], []
    last_todo = last_doing = last_done = 0

    for i in range(days):
        d = bd_start + timedelta(days=i)
        if d <= cutoff:
            todo = 0
            doing = 0
            done = 0
            for _, row in df.iterrows():
                d_start = row.get("date_start")
                d_done = _done_plan_value(row)
                if _is_valid_date(d_done) and d_done <= d:
                    done += 1
                elif _is_valid_date(d_start) and d_start <= d:
                    doing += 1
                else:
                    todo += 1
            last_todo, last_doing, last_done = todo, doing, done
        else:
            todo, doing, done = last_todo, last_doing, last_done
        todo_list.append(todo)
        doing_list.append(doing)
        done_list.append(done)

    dates = [
        datetime(bd_start.year, bd_start.month, (bd_start + timedelta(days=i)).day)
        for i in range(days)
    ]
    return dates, todo_list, doing_list, done_list


def build_wip(df, sprint_start, export_date):
    bd_start, bd_end = _month_bounds(sprint_start)
    days = (bd_end - bd_start).days + 1
    cutoff = export_date if isinstance(export_date, date) else export_date

    bucket_names = ["Concluído", "Em Desenvolvimento", "Em Refinamento", "Gestão", "UX/UI"]
    bucket_matrix = {b: [0] * days for b in bucket_names}
    last_state = {b: 0 for b in bucket_names}

    for i in range(days):
        d = bd_start + timedelta(days=i)
        if d <= cutoff:
            done_count = 0
            active_counts = {
                "Em Desenvolvimento": 0,
                "Em Refinamento": 0,
                "Gestão": 0,
                "UX/UI": 0,
            }
            for _, row in df.iterrows():
                d_start = row.get("date_start")
                d_done = _done_plan_value(row)
                if _is_valid_date(d_done) and d_done <= d:
                    done_count += 1
                    continue
                if _is_valid_date(d_start):
                    end = d_done if _is_valid_date(d_done) else cutoff
                    if d_start <= d < end:
                        phase = _wip_phase_from_row(row)
                        if phase in active_counts:
                            active_counts[phase] += 1

            bucket_matrix["Concluído"][i] = done_count
            bucket_matrix["Em Desenvolvimento"][i] = active_counts["Em Desenvolvimento"]
            bucket_matrix["Em Refinamento"][i] = active_counts["Em Refinamento"]
            bucket_matrix["Gestão"][i] = active_counts["Gestão"]
            bucket_matrix["UX/UI"][i] = active_counts["UX/UI"]
            for b in bucket_names:
                last_state[b] = bucket_matrix[b][i]
        else:
            for b in bucket_names:
                bucket_matrix[b][i] = last_state[b]

    dates = [
        datetime(bd_start.year, bd_start.month, (bd_start + timedelta(days=i)).day)
        for i in range(days)
    ]
    return dates, bucket_names, bucket_matrix, None


def build_cts(df, sprint_start):
    bd_start, bd_end = _month_bounds(sprint_start)
    days = (bd_end - bd_start).days + 1

    def _all_hus(labels_text):
        out = []
        seen = set()
        for token in [p.strip() for p in str(labels_text).split(";") if p.strip()]:
            m = re.search(r"\b(HU\s*0*\d+)\b", token, flags=re.IGNORECASE)
            if not m:
                continue
            num = re.sub(r"\D", "", m.group(1))
            hu = f"HU{num.zfill(3)}"
            if hu not in seen:
                seen.add(hu)
                out.append(hu)
        return out

    hu_set = set()
    hu_full = {}
    for _, row in df.iterrows():
        labels_text = str(row.get("labels", ""))
        full = _extract_hu_full_label(labels_text)
        for hu in _all_hus(labels_text):
            hu_set.add(hu)
            if hu not in hu_full:
                hu_full[hu] = full or hu
    hu_list = sorted(hu_set)

    hu_index = {h: i for i, h in enumerate(hu_list)}
    cts_matrix = [[0] * days for _ in hu_list]
    fora_hu_daily = [0] * days

    done_rows = df.copy()
    for _, row in done_rows.iterrows():
        d = _done_plan_value(row)
        if not _is_valid_date(d):
            continue
        offset = (d - bd_start).days
        if 0 <= offset < days:
            hu_keys = _all_hus(row.get("labels", ""))
            if not hu_keys:
                fora_hu_daily[offset] += 1
            else:
                for hu in hu_keys:
                    idx = hu_index.get(hu)
                    if idx is not None:
                        cts_matrix[idx][offset] += 1

    dates = [
        datetime(bd_start.year, bd_start.month, (bd_start + timedelta(days=i)).day)
        for i in range(days)
    ]
    return dates, [hu_full.get(h, h) for h in hu_list], cts_matrix, fora_hu_daily


def build_dispersao_daily(df, sprint_start, sprint_end):
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1

    def _all_hus(labels_text):
        out = []
        seen = set()
        for token in [p.strip() for p in str(labels_text).split(";") if p.strip()]:
            m = re.search(r"\b(HU\s*0*\d+)\b", token, flags=re.IGNORECASE)
            if not m:
                continue
            num = re.sub(r"\D", "", m.group(1))
            hu = f"HU{num.zfill(3)}"
            if hu not in seen:
                seen.add(hu)
                out.append(hu)
        return out

    hu_set = set()
    for labels_text in df["labels"].fillna("").astype(str).tolist():
        for hu in _all_hus(labels_text):
            hu_set.add(hu)
    hu_list = sorted(hu_set)
    hu_index = {hu: idx for idx, hu in enumerate(hu_list)}

    hu_matrix = [[0.0] * days for _ in hu_list]
    nao_hu_daily = [0.0] * days

    done_rows = df.copy()
    for _, row in done_rows.iterrows():
        d = _done_plan_value(row)
        if not _is_valid_date(d):
            continue
        offset = (d - bd_start).days
        if 0 <= offset < days:
            hu_keys = _all_hus(row.get("labels", ""))
            if not hu_keys:
                nao_hu_daily[offset] += 1.0
            else:
                for hu in hu_keys:
                    idx = hu_index.get(hu)
                    if idx is not None:
                        hu_matrix[idx][offset] += 1.0

    return hu_list, days, hu_matrix, nao_hu_daily


def _w(ws, row, col, value):
    """Write a value to a specific cell."""
    ws.cell(row=row, column=col).value = value


def _fill_kpi_sheet(wb, kpis, sprint_name, sprint_goal,
                    projeto="", gerente="", linkedin="", export_date=None,
                    nome_arquivo="", product_goal="",
                    write_cabecalhos=True):
    """
    Preenche os KPIs no template.

    Novo template (xKPI + gp_Cabecalhos):
      xKPI: métricas numéricas (15 linhas)
      gp_Cabecalhos: metadados do projeto (PROJETO, GERENTE, etc.)

    Template legado (Cabecalhos e KPI):
      Estrutura combinada com metadados + KPIs em 13 linhas.
    """
    if export_date and hasattr(export_date, "strftime"):
        export_date_display = export_date.strftime("%d/%m/%Y")
    else:
        export_date_display = str(export_date) if export_date else ""
    pct = round(kpis["done"] / kpis["total"], 4) if kpis["total"] else 0

    if "xKPI" in wb.sheetnames:
        # ── Novo template: xKPI contém apenas métricas ──
        ws = wb["xKPI"]
        xkpi_rows = [
            ("STORYPOINTS",                kpis.get("storypoints", 0)),
            ("TAREFAS",                    kpis["total"]),
            ("TAREFAS CONCLUÍDAS",         kpis["done"]),
            ("PENDENTES",                  kpis["pending"]),
            ("EM BACKLOG",                 kpis.get("backlog_total", 0)),
            ("AUSENTES EM HU",             kpis["sem_hu"]),
            ("NÃO PREVISTAS",              kpis["nao_prev"]),
            ("CYCLE TIME TASK",            round(kpis.get("ct_task") or 0, 2)),
            ("LEAD TIME TASK",             round(kpis.get("lt_task") or 0, 2)),
            ("CYCLE TIME HU",              round(kpis.get("ct_hu") or 0, 2)),
            ("LEAD TIME HU",               round(kpis.get("lt_hu") or 0, 2)),
            ("CYCLE TIME SP.",             round(kpis.get("ct_sp") or 0, 2)),
            ("LEAD TIME SP.",              round(kpis.get("lt_sp") or 0, 2)),
            ("Fluxo Contínuo",             kpis.get("fluxo_continuo", 0)),
            ("% de Entrega:",              pct),
            ("Stakeholders com tarefas",   kpis.get("stakeholders", 0)),
        ]
        for i, (lbl, val) in enumerate(xkpi_rows, start=1):
            _w(ws, i, 1, lbl)
            _w(ws, i, 2, val)

        # ── gp_Cabecalhos: metadados do projeto (só preenche se write_cabecalhos=True) ──
        if write_cabecalhos and "gp_Cabecalhos" in wb.sheetnames:
            wsc = wb["gp_Cabecalhos"]
            cab_rows = [
                ("PROJETO",                                projeto or sprint_name),
                ("NOME DO GERENTE DO PROJETO",             gerente),
                ("link para LINKEDIN DO GERENTE DO PROJETO", linkedin),
                ("PRODUCT GOAL (OBJETIVO DO PROJETO)",     product_goal or sprint_goal or ""),
                ("SPRINT GOAL (OBJETIVO DA SPRINT)",       sprint_goal or ""),
                ("LINK ARQUIVO BASE / PLANO",              nome_arquivo),
                ("DATA DA EXTRACAO DOS DADOS",             export_date_display),
            ]
            for i, (lbl, val) in enumerate(cab_rows, start=1):
                _w(wsc, i, 1, lbl)
                _w(wsc, i, 2, val)

    else:
        # ── Template legado (Cabecalhos e KPI) ──
        ws = wb["Cabecalhos e KPI"] if "Cabecalhos e KPI" in wb.sheetnames else None
        if ws is None:
            return
        labels_col_a = [
            "PROJETO",
            "NOME DO GERENTE DO PROJETO",
            "link para LINKEDIN DO GERENTE DO PROJETO",
            "DATA DA EXTRACAO DOS DADOS",
            "NOME DO ARQUIVO BASE / PLANO",
            "SPRINT GOAL (OBJETIVO)",
            "TAREFAS TOTAL",
            "TAREFAS CONCLUÍDAS",
            "PENDENTES",
            "NÃO PREVISTAS",
            "% ENTREGA",
            "EM BACKLOG",
            "AUSENTES EM HU",
        ]
        for i, lbl in enumerate(labels_col_a, start=1):
            _w(ws, i, 1, lbl)
        values_col_b = [
            projeto or sprint_name,
            gerente,
            linkedin,
            export_date_display,
            nome_arquivo,
            sprint_goal or "",
            kpis["total"],
            kpis["done"],
            kpis["pending"],
            kpis["nao_prev"],
            pct,
            kpis.get("backlog_total", 0),
            kpis["sem_hu"],
        ]
        for i, val in enumerate(values_col_b, start=1):
            _w(ws, i, 2, val)


def _fill_hu_sheet(wb, hu_list, hu_full_names, extras_count=0):
    """
    HUs sheet:  A = full HU label,  B = total task count
    Template has 5 data rows (rows 2-6) + row 7 = EXTRAS (AUSENTE EM HU).
    O gráfico TAREFAS POR HU referencia HUs!B2:B7 (6 linhas).
    """
    sheet_name = "xHUs" if "xHUs" in wb.sheetnames else "HUs"
    ws = wb[sheet_name]
    TEMPLATE_ROWS = 5
    top = sorted(hu_list, key=lambda x: -x[1])[:TEMPLATE_ROWS]

    for i in range(TEMPLATE_ROWS):
        r = i + 2
        if i < len(top):
            hu_id, total, done = top[i]
            full_name = hu_full_names.get(hu_id, hu_id)
            _w(ws, r, 1, full_name)
            _w(ws, r, 2, total)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)

    # Linha 7: EXTRAS (tarefas sem HU) – referenciado pelo gráfico TAREFAS POR HU
    _w(ws, 7, 1, "EXTRAS (AUSENTE EM HU)")
    _w(ws, 7, 2, extras_count)


def _fill_burndown_sheet(wb, df, sprint_start, sprint_end, export_date):
    """
    BurndownTarefas:  A = datetime,  B = Meta,  C = Planejado,  D = A Realizar
    Template has 31 data rows (rows 2-32).
    X-axis always covers the full month of sprint_end (1st to last day).
    """
    sheet_name = "xBurndownTarefas" if "xBurndownTarefas" in wb.sheetnames else "BurndownTarefas"
    ws = wb[sheet_name]
    ROWS = 31
    dfw = df[~df["is_backlog"]].copy()
    total = len(dfw)

    # Snap to full month of sprint_end
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1  # 28-31

    # Planejado: remove tasks only when due date is within month and <= current day.
    due_offsets = [
        (d - bd_start).days
        for d in dfw["date_due"].tolist()
        if d is not None and not (hasattr(d, '_typ') or str(d) == 'NaT')
        and bd_start <= d <= bd_end
    ]
    due_hist = [0] * days
    for off in due_offsets:
        due_hist[off] += 1
    due_prefix = []
    running = 0
    for count in due_hist:
        running += count
        due_prefix.append(running)

    # A Realizar: total - tasks done up to day (includes tasks done before sprint month).
    done_dates = sorted(
        d for d in dfw.loc[dfw["done"] & dfw["date_done"].notna(), "date_done"].tolist()
    )

    for i in range(ROWS):
        r = i + 2
        if i < days:
            d = bd_start + timedelta(days=i)
            dt = datetime(d.year, d.month, d.day)
            meta = round(total * (1 - i / max(days - 1, 1)))
            plan = total - due_prefix[i]

            concluded_until_day = bisect_right(done_dates, d)
            rlz = total - concluded_until_day

            _w(ws, r, 1, dt)
            _w(ws, r, 2, meta)
            _w(ws, r, 3, plan)
            _w(ws, r, 4, rlz)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)


def _fill_burndown_hu_sheet(wb, df, sprint_start, sprint_end, export_date):
    """
    BurndownHU:  A = datetime,  B = Meta,  C = Planejado,  D = A Realizar
    Template has 31 data rows (rows 2-32).
    X-axis always covers the full month of sprint_end (1st to last day).

    Meta   : decresce de forma LINEAR a partir da soma total de tarefas nas HUs.
    Planejado: degraus proporcionais ao tamanho de cada HU.
    A Realizar: soma das tarefas totais das HUs que ainda NÃO estão 100% concluídas.
                (se HU054 tem 36 tarefas e 1 resta, conta 36)
    """
    sheet_name = "xBurndownHU" if "xBurndownHU" in wb.sheetnames else "BurndownHU"
    ws = wb[sheet_name]
    ROWS = 31
    dfw = df[df["hu"] != ""].copy()

    # Snap to full month of sprint_end
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1  # 28-31

    hu_groups = list(dfw.groupby("hu"))

    # Info por HU: (task_count, completion_date_or_None)
    hu_info = []
    for hu_id, grp in hu_groups:
        hu_total = len(grp)
        done_mask = grp["done"].astype(bool)
        if bool(done_mask.all()):
            done_dates = [
                d for d in grp.loc[done_mask & grp["date_done"].notna(), "date_done"].tolist()
                if d is not None
            ]
            completion_date = max(done_dates) if done_dates else (
                export_date if bd_start <= export_date <= bd_end else bd_end
            )
        else:
            completion_date = None  # não concluída na sprint
        hu_info.append((hu_total, completion_date))

    # Meta inicial = soma de tarefas de todas as HUs
    total_hu_tasks = sum(t for t, _ in hu_info)

    # Planejado: quedas proporcionais ao tamanho de cada HU, em degraus 7/5 dias
    # Ordenar HUs por data de conclusão (concluídas primeiro, pendentes por último)
    hu_info_sorted = sorted(
        hu_info,
        key=lambda x: (x[1] is None, x[1] or date.max)
    )
    drop_offsets = []
    if hu_info_sorted:
        anchor = export_date
        if anchor < bd_start:
            anchor = bd_start
        if anchor > bd_end:
            anchor = bd_end
        first_off = (anchor - bd_start).days
        drop_offsets = [first_off]
        next_off = first_off + 7
        for _ in range(len(hu_info_sorted) - 2):
            if next_off < days - 1:
                drop_offsets.append(next_off)
                next_off += 5
        drop_offsets.append(days - 1)
        # Preenche se necessário
        used = set(drop_offsets)
        cand = days - 2
        while len(drop_offsets) < len(hu_info_sorted) and cand >= 0:
            if cand not in used:
                drop_offsets.insert(-1, cand)
                used.add(cand)
            cand -= 1
        drop_offsets = sorted(drop_offsets[:len(hu_info_sorted)])

    # Planejado: soma acumulada das quedas de cada HU no offset correspondente
    # Cria lista de (offset, task_count) para cada "queda" planejada
    plan_drops = []  # (day_offset, task_count_to_drop)
    for idx, (drop_off) in enumerate(drop_offsets):
        if idx < len(hu_info_sorted):
            task_cnt, _ = hu_info_sorted[idx]
            plan_drops.append((drop_off, task_cnt))

    # Monta array de Planejado por dia
    plan_remaining = total_hu_tasks
    plan_by_day = []
    drops_by_day = defaultdict(int)
    for off, cnt in plan_drops:
        drops_by_day[off] += cnt
    running_plan = total_hu_tasks
    for i in range(days):
        running_plan = max(running_plan - drops_by_day.get(i, 0), 0)
        plan_by_day.append(running_plan)
    # Garantir que último dia = 0
    if plan_by_day:
        plan_by_day[-1] = 0

    for i in range(ROWS):
        r = i + 2
        if i < days:
            d = bd_start + timedelta(days=i)
            dt = datetime(d.year, d.month, d.day)
            meta = round(total_hu_tasks * (1 - i / max(days - 1, 1)))

            plan = plan_by_day[i]

            # A Realizar: soma de tarefas de HUs NÃO 100% concluídas até o dia d
            a_realiz = 0
            for hu_total, comp_date in hu_info:
                if comp_date is None or comp_date > d:
                    a_realiz += hu_total
            # Primeiro dia = Meta inicial (conforme especificação)
            if i == 0:
                a_realiz = total_hu_tasks

            _w(ws, r, 1, dt)
            _w(ws, r, 2, meta)
            _w(ws, r, 3, plan)
            _w(ws, r, 4, a_realiz)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)


def _fill_dispersao_sheet(wb, hu_list, days, hu_matrix, nao_hu_daily,
                           bd_start, hu_full_names):
    """
    Dispersao — interleaved per-series X/Y layout:

      Row  1: X values for series 0  (day + offset_0, col B:AF)
      Row  2: name (col A) + Y values for series 0  (band 1 or None)
      Row  3: X values for series 1  (day + offset_1)
      Row  4: name + Y values for series 1  (band 2 or None)
      Row  5: X values for series 2
      Row  6: name + Y values for series 2
      Row  7: X values for series 3
      Row  8: name + Y values for series 3
      Row  9: X values for series 4
      Row 10: name + Y values for series 4
      Row 11: X values for series 5 (Nao em HU)
      Row 12: name + Y values for series 5  (band 6 or None)

    Per-series horizontal jitter offsets spread points around the integer day
    so same-day events from different HUs are not stacked vertically.
    Offsets: -0.30, -0.18, -0.06, +0.06, +0.18, +0.30

    Y values use row-band positioning: series i → Y = i+1 when tasks were
    completed that day, else None (dot hidden).

    Template has 31 date columns (B:AF) and 12 data rows (rows 1-12).
    bd_start: first day of the sprint month.
    """
    sheet_name = "xDispersaoTarefas" if "xDispersaoTarefas" in wb.sheetnames else "Dispersao"
    ws        = wb[sheet_name]
    DATE_COLS = 31   # columns B:AF
    HU_ROWS   = 5    # series 0-4 (top 5 HUs)

    # Per-series X jitter offsets — spread +-0.30 around integer day
    OFFSETS = [-0.30, -0.18, -0.06, +0.06, +0.18, +0.30]

    # Clear any old data in rows 1-12
    for r in range(1, 13):
        for c in range(1, DATE_COLS + 2):
            ws.cell(r, c).value = None

    def _write_series_rows(x_row, y_row, band_y, day_values, name):
        """Write X row (jittered day) and Y row (band or None) for one series."""
        offset = OFFSETS[band_y - 1]
        ws.cell(x_row, 1).value = None          # col A of X row blank
        ws.cell(y_row, 1).value = name           # series name in col A of Y row
        for j in range(DATE_COLS):
            col = j + 2
            if j < days:
                day_num  = (bd_start + timedelta(days=j)).day
                x_val    = round(day_num + offset, 4)
                has_data = day_values[j] > 1e-9
            else:
                x_val    = None
                has_data = False

            ws.cell(x_row, col).value = x_val
            ws.cell(x_row, col).number_format = "0.00"

            ws.cell(y_row, col).value = band_y if has_data else None
            ws.cell(y_row, col).number_format = "0"

    # Series 0-4: individual HU rows
    top_hu = hu_list[:HU_ROWS]
    for i in range(HU_ROWS):
        x_row  = 2 * i + 1   # 1, 3, 5, 7, 9
        y_row  = 2 * i + 2   # 2, 4, 6, 8, 10
        band_y = i + 1        # 1, 2, 3, 4, 5
        if i < len(top_hu):
            hu_id     = top_hu[i]
            full_name = hu_full_names.get(hu_id, hu_id)
            day_vals  = hu_matrix[i]
        else:
            full_name = None
            day_vals  = [0.0] * days
        _write_series_rows(x_row, y_row, band_y, day_vals, full_name)

    # Series 5: Nao em HU -> rows 11 (X) and 12 (Y), band 6
    _write_series_rows(11, 12, 6, nao_hu_daily, "Nao Presente em HU")


def _fill_colaborador_sheet(wb, collab_rows):
    """
    PorColaborador:
      A = 'Nome (done/total)',  B = done count,  C = pending count
    Template has 11 data rows (rows 2-12).
    """
    sheet_name = "xResponsaveis" if "xResponsaveis" in wb.sheetnames else "PorColaborador"
    ws   = wb[sheet_name]
    ROWS = 11
    top  = collab_rows[:ROWS]

    for i in range(ROWS):
        r = i + 2
        if i < len(top):
            name, done, pend = top[i]
            total = done + pend
            label = f"{name} ({done}/{total})"
            _w(ws, r, 1, label)
            _w(ws, r, 2, done)
            _w(ws, r, 3, pend)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)


def _fill_areas_sheet(wb, area_rows):
    """
    xAreas (novo template): linhas dinâmicas, A = área (label com "."), B = total.
    Areas (template antigo): 12 linhas fixas com TEMPLATE_AREAS.
    """
    if "xAreas" in wb.sheetnames:
        ws = wb["xAreas"]
        # New template: dynamic rows from build_areas (labels starting with ".")
        # Clear existing data rows
        for r in range(2, ws.max_row + 2):
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
        for i, (area, done, pend) in enumerate(area_rows):
            _w(ws, i + 2, 1, area)
            _w(ws, i + 2, 2, done + pend)
    elif "Areas" in wb.sheetnames:
        ws = wb["Areas"]
        area_totals = {}
        for area, done, pend in area_rows:
            display = AREA_DISPLAY.get(area, area)
            area_totals[display] = done + pend
        for i, area_name in enumerate(TEMPLATE_AREAS):
            r = i + 2
            _w(ws, r, 1, area_name)
            _w(ws, r, 2, area_totals.get(area_name, 0))


def _fill_hu_in_out_sheet(wb, in_out_rows):
    """
    HU_inOut: 2 rows, A = group name, B = total count
    """
    sheet_name = "xHU_inOut" if "xHU_inOut" in wb.sheetnames else "HU_inOut"
    ws = wb[sheet_name]
    for i, (label, total) in enumerate(in_out_rows):
        r = i + 2
        _w(ws, r, 1, label)
        _w(ws, r, 2, total)


def _fill_categoria_sheet(wb, cat_rows):
    """
    PorCategoria:
      A = 'LABEL (done/total)',  B = done,  C = pending
    Template has 16 data rows (rows 2-17).
    """
    sheet_name = "xRotulos" if "xRotulos" in wb.sheetnames else "PorCategoria"
    ws   = wb[sheet_name]
    ROWS = 16
    top  = cat_rows[:ROWS]

    for i in range(ROWS):
        r = i + 2
        if i < len(top):
            cat, done, pend = top[i]
            total = done + pend
            label = f"{cat} ({done}/{total})"
            _w(ws, r, 1, label)
            _w(ws, r, 2, done)
            _w(ws, r, 3, pend)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)


def _fill_categoria_bubbles_sheet(wb, bub_rows, template_effort=None):
    """
    PorCategoriaBubbles:
      A = 'LABEL (done/total)',  B = done,  C = total,  D = esforço
    Template has 17 data rows (rows 2-18).
    A última linha (row 18) é reservada para "BUGs e Ajustes" (sempre pinado ao final).

    Se template_effort for fornecido (dict {label_norm: effort}), usa os
    valores de esforço planejado do template em vez do esforço de checklist.
    """
    sheet_name = "PorCategoriaBubbles" if "PorCategoriaBubbles" in wb.sheetnames else None
    if sheet_name is None:
        return  # Sheet not in new template — skip silently
    ws   = wb[sheet_name]
    ROWS = 17
    top  = bub_rows[:ROWS]

    for i in range(ROWS):
        r = i + 2
        if i < len(top):
            cat, done, total, effort_checklist = top[i]
            # Usa esforço do template (planejamento) quando disponível
            if template_effort:
                effort = template_effort.get(_norm_label_key(cat), effort_checklist)
            else:
                effort = effort_checklist
            label = f"{cat} ({done}/{total})"
            _w(ws, r, 1, label)
            _w(ws, r, 2, done)
            _w(ws, r, 3, total)
            _w(ws, r, 4, effort)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)


def _fill_rotulos_sheet(wb, rotulos_rows):
    """
    xRotulos: A = Rótulo (done/total), B = done, C = pending, D = LeadTime, E = CycleTime
    """
    sheet_name = "xRotulos" if "xRotulos" in wb.sheetnames else ("PorCategoria" if "PorCategoria" in wb.sheetnames else None)
    if not sheet_name:
        return
    ws = wb[sheet_name]
    # Clear existing data (keep row 1 header)
    for r in range(2, ws.max_row + 1):
        for c in range(1, 6):
            ws.cell(r, c).value = None
    # Write header row 1
    headers = ["Rótulo (qtd concluído / qtd total)", "Qtd. Concluído", "Qtd Total", "Leadtime (em dias)", "Cycletime (em dias)"]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c).value = h
    for i, (display, done, pend, lt, ct) in enumerate(rotulos_rows):
        r = i + 2
        _w(ws, r, 1, display)
        _w(ws, r, 2, done)
        _w(ws, r, 3, pend)
        _w(ws, r, 4, lt)
        _w(ws, r, 5, ct)


def _fill_responsaveis_sheet(wb, resp_rows):
    """
    xResponsaveis: A = Nome (done/total), B = done_bucket, C = pending, D = LeadTime, E = CycleTime
    """
    sheet_name = "xResponsaveis" if "xResponsaveis" in wb.sheetnames else ("PorColaborador" if "PorColaborador" in wb.sheetnames else None)
    if not sheet_name:
        return
    ws = wb[sheet_name]
    for r in range(2, ws.max_row + 1):
        for c in range(1, 6):
            ws.cell(r, c).value = None
    headers = ["Nome (qtd concluído / qtd total)", "Qtd. no Bucket Concluído", "Qtd. Fora do Bucket Concluído", "LeadTime (Dias)", "CycleTime (Dias)"]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c).value = h
    for i, (display, done_bkt, pend, lt, ct) in enumerate(resp_rows):
        r = i + 2
        _w(ws, r, 1, display)
        _w(ws, r, 2, done_bkt)
        _w(ws, r, 3, pend)
        _w(ws, r, 4, lt)
        _w(ws, r, 5, ct)


def _fill_cfd_sheet(wb, dates, todo_list, doing_list, done_list):
    """
    xCFD: A = Data, B = To Do, C = Doing, D = Done
    """
    sheet_name = "xCFD" if "xCFD" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]
    _w(ws, 1, 1, "Data"); _w(ws, 1, 2, "To Do"); _w(ws, 1, 3, "Doing"); _w(ws, 1, 4, "Done")
    for i, (dt, td, do, dn) in enumerate(zip(dates, todo_list, doing_list, done_list)):
        r = i + 2
        _w(ws, r, 1, dt); _w(ws, r, 2, td); _w(ws, r, 3, do); _w(ws, r, 4, dn)


def _fill_wip_sheet(wb, dates, bucket_names, bucket_matrix, x_total=None):
    """
    xWIP: A = Data, B..N = one dynamic column per state/label, last col = Total.
    """
    sheet_name = "xWIP" if "xWIP" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]

    # Limpa valores antigos (preserva estilo/formatação).
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    _w(ws, 1, 1, "Data")
    for j, bkt in enumerate(bucket_names):
        _w(ws, 1, j + 2, bkt)
    total_col = len(bucket_names) + 2
    if x_total is not None:
        _w(ws, 1, total_col, "Total")
    for i, dt in enumerate(dates):
        r = i + 2
        _w(ws, r, 1, dt)
        for j, bkt in enumerate(bucket_names):
            _w(ws, r, j + 2, bucket_matrix[bkt][i])
        if x_total is not None:
            _w(ws, r, total_col, x_total[i])


def _fill_cts_sheet(wb, dates, hu_labels, cts_matrix, fora_hu_daily):
    """
    xCTS: A = Data, then pairs (HU_nome, qtd) for each HU, last pair = Fora de HU
    Row 1: header with HU names and 'qtd' alternating
    """
    sheet_name = "xCTS" if "xCTS" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]
    # Header row
    def _short_hu_label(lbl):
        s = str(lbl)
        s = s.replace("Construção UI/UX Desktop", "Desktop")
        s = s.replace("Construção UI/UX Mobile", "Mobile")
        s = s.replace("Construção DEV", "Construção")
        s = s.replace("Testes DEV", "Testes")
        s = s.replace("Registro e Publicação", "Registro")
        return s

    _w(ws, 1, 1, "Data")
    all_labels = list(hu_labels) + ["Fora de HU"]
    for j, lbl in enumerate(all_labels):
        _w(ws, 1, j * 2 + 2, _short_hu_label(lbl))
        _w(ws, 1, j * 2 + 3, "qtd")
    # Data rows
    for i, dt in enumerate(dates):
        r = i + 2
        _w(ws, r, 1, dt)
        for j, (lbl, vals) in enumerate(zip(hu_labels, cts_matrix)):
            v = vals[i]
            # index col (HU name repeated when there's data)
            _w(ws, r, j * 2 + 2, lbl if v > 0 else None)
            _w(ws, r, j * 2 + 3, v)
        # Fora de HU
        j = len(hu_labels)
        v = fora_hu_daily[i]
        _w(ws, r, j * 2 + 2, "Fora de HU" if v > 0 else None)
        _w(ws, r, j * 2 + 3, v)


def _fill_histograma_sheet(wb, hist_31, indicativos):
    """
    xHistograma: A = DIAS (1-31), B = QTD, C = INDICATIVO
    """
    sheet_name = "xHistograma" if "xHistograma" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]
    _w(ws, 1, 1, "DIAS"); _w(ws, 1, 2, "QTD"); _w(ws, 1, 3, "INDICATIVO")
    for i in range(31):
        r = i + 2
        _w(ws, r, 1, i + 1)
        _w(ws, r, 2, hist_31[i] if i < len(hist_31) else 0)
        _w(ws, r, 3, indicativos[i] if i < len(indicativos) else None)


def _fill_histograma2_sheet(wb, stats_rows):
    """
    xHistograma2: A = descricao, B = valor
    """
    sheet_name = "xHistograma2" if "xHistograma2" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]
    _w(ws, 1, 1, "Descritivo da Métrica (Legenda)"); _w(ws, 1, 2, "Dado numérico de referência (dias)")
    for i, (desc, val) in enumerate(stats_rows):
        r = i + 2
        _w(ws, r, 1, desc)
        _w(ws, r, 2, val)


def _fill_esforco_hu_sheet(wb, hu_full_names, esforco_detailed):
    """
    EsforcoHU:
      A = HU full name,  B = max effort per profile,  C = total effort
    Template has 5 data rows (rows 2-6).
    """
    ws   = wb["EsforcoHU"]
    ROWS = 5
    top  = esforco_detailed[:ROWS]

    for i in range(ROWS):
        r = i + 2
        if i < len(top):
            hu_id, max_ep, total_effort = top[i]
            full_name = hu_full_names.get(hu_id, hu_id)
            _w(ws, r, 1, full_name)
            _w(ws, r, 2, max_ep)
            _w(ws, r, 3, total_effort)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)


def _clear_report_planning_tables(wb):
    """
    Limpa APENAS as células de dados das tabelas Planejamento e Histórico,
    preservando integralmente a estrutura de cabeçalhos do template:
      - Linha 6: "PRIORIDADE" + números 1-9 (Planejamento) / "SPRINT 0..07" (Histórico)
      - Linha 7: nomes das HUs (Planejamento) / datas das sprints (Histórico)
      - Col G (7): labels de perfil do Planejamento
      - Col R (18): fórmulas AVERAGE do Histórico
      - Col U (21): labels de perfil do Histórico
      - Linhas 19-20: fórmulas LARGE/SUM do Planejamento

    O que é limpo (somente valores manuais):
      Planejamento — linhas 8-18, colunas H-L (8-12): storypoints por HU por perfil
      Histórico    — linhas 8-21, colunas V-AC (22-29): valores históricos por sprint
    """
    if "Report" not in wb.sheetnames:
        return
    ws = wb["Report"]

    # Planejamento: limpa textos antigos atrás do gráfico de roadmap (col G)
    # para evitar "sobreposição visual" quando o gráfico é transparente.
    for r in range(8, 21):
        try:
            ws.cell(r, 7).value = None
        except AttributeError:
            pass  # MergedCell — skip

    # Planejamento: só as células de storypoints manuais (rows 8-18, cols H-L = 8-12)
    for r in range(8, 19):
        for c in range(8, 13):
            cell = ws.cell(r, c)
            try:
                cell.value = None
            except AttributeError:
                pass  # MergedCell — skip

    # Histórico: só os valores históricos (rows 8-21, cols V-AC = 22-29)
    for r in range(8, 22):
        for c in range(22, 30):
            cell = ws.cell(r, c)
            try:
                cell.value = None
            except AttributeError:
                pass  # MergedCell — skip


def _clear_report_roadmap_overlay_cells(wb):
    """
    Limpa cÃ©lulas que ficam por trÃ¡s do grÃ¡fico de Roadmap no Report.

    O chart de roadmap (chart9) cobre aproximadamente C11:AC32. Como esse
    grÃ¡fico pode ser transparente, qualquer texto/fÃ³rmula nessa Ã¡rea aparece
    sobreposto (incluindo #DIV/0! de fÃ³rmulas antigas).
    """
    if "Report" not in wb.sheetnames:
        return
    ws = wb["Report"]

    # Bloco visual coberto pelo roadmap.
    for r in range(11, 33):       # 11..32
        for c in range(3, 30):    # C..AC
            try:
                ws.cell(r, c).value = None
            except AttributeError:
                pass  # MergedCell â€” skip

    # Defesa extra para labels/formulas que podem reaparecer no eixo central.
    for r in range(7, 23):        # 7..22
        for c in (7, 18, 21):     # G, R, U
            try:
                ws.cell(r, c).value = None
            except AttributeError:
                pass


def _ensure_report_manual_structure(wb):
    """
    Garante estrutura mínima do bloco manual do Report (Planejamento/Histórico).
    Restaura labels de perfil APENAS se estiverem ausentes (não sobrescreve).

    Colunas corretas conforme template:
      Planejamento: labels na col G (7), dados HU em H-L (8-12)
      Histórico:    AVERAGE em col R (18), labels na col U (21), dados em V-AC (22-29)
    """
    if "Report" not in wb.sheetnames:
        return

    ws = wb["Report"]

    profile_labels_plan = [
        "PO (DOCS, LINKS, LEIS, ETC.)",
        "REQUISITOS",
        "UX/UI  (DESKTOP)",
        "UX/UI  (MOBILE)",
        "DEV FRONT (DESKTOP)",
        "DEV FRONT (MOBILE)",
        "DEV BACK",
        "DEV BANCO DE DADOS",
        "DEV BI",
        "TESTE (possíveis no prazo da sprint)",
        "ARQUITETURA",
        "Mensuração (maior esforço individual)",
        "Mensuracao soma total esforço por HU",
    ]

    profile_labels_hist = profile_labels_plan + [
        "Quantidade de Histórias de Usuário (HU)",
    ]

    def _safe_cell_write(ws, row, col, value):
        """Write value to cell only if it's writable (not a MergedCell slave)."""
        try:
            cell = ws.cell(row, col)
            existing = cell.value
            if not _strip(str(existing) if existing is not None else ""):
                cell.value = value
        except AttributeError:
            pass  # MergedCell slave — skip

    # ── Planejamento (área superior): NÃO repor labels na col G (linhas 8-20) ──
    # Esses textos aparecem "por trás" do gráfico de roadmap (fundo transparente).
    # O bloco de planejamento da sprint que alimenta fórmulas fica em linhas inferiores.

    # ── Histórico: labels na col U (21), linhas 7-21 ──
    _safe_cell_write(ws, 7, 21, "ID ou DESCRIÇÃO DO ITEM")
    for i, label in enumerate(profile_labels_hist, start=8):  # U8:U21
        _safe_cell_write(ws, i, 21, label)

    # ── Fórmulas AVERAGE do Histórico: col R (18), linhas 8-21 ──
    # (=AVERAGE(V{r}:AC{r}) — referencia cols V-AC = dados históricos)
    for r in range(8, 22):
        _safe_cell_write(ws, r, 18, f"=AVERAGE(V{r}:AC{r})")

    # Nota: NÃO preencher zeros nas células vazias do Histórico.
    # As células vazias devem ficar em branco para o usuário preencher manualmente.


def _update_report_metadata(wb, export_date, author):
    """
    Update Report sheet header cells:
      V2 = data da extração,  Y2 = author name
    """
    if "Report" not in wb.sheetnames:
        return
    ws = wb["Report"]
    # Write as formatted string to avoid ####### when column is narrow
    for row, col, val in [
        (2, 22, export_date.strftime("%d/%m/%Y")),  # V2
        (2, 25, author),                             # Y2
    ]:
        try:
            ws.cell(row, col).value = val
        except AttributeError:
            pass  # MergedCell slave — skip


def _clear_sheet_non_formula(ws, r1=1, c1=1, r2=None, c2=None):
    if r2 is None:
        r2 = ws.max_row
    if c2 is None:
        c2 = ws.max_column
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            val = cell.value
            if isinstance(val, str) and val.startswith("="):
                continue
            try:
                cell.value = None
            except Exception:
                pass


def create_clean_template(source_template_path, output_template_path):
    """
    Build a reusable clean template from an existing workbook:
      - clear data tabs (x* and legacy data tabs)
      - clear manual gp_* input values
      - keep formulas/charts/layout
    """
    wb = load_workbook(source_template_path)

    legacy_data_tabs = {
        "KPI", "Roadmap", "HUs", "BurndownTarefas", "BurndownHU",
        "Dispersao", "Colaboradores", "Areas", "HUInOut",
        "Categorias", "PorCategoria", "xHU", "xHU_inOut",
    }

    # 1) Data tabs used by processing pipeline
    for name in wb.sheetnames:
        ws = wb[name]
        if name.startswith("x") or name in legacy_data_tabs:
            _clear_sheet_non_formula(ws)

    # 2) Manual gp_* tabs
    if "gp_Cabecalhos" in wb.sheetnames:
        _clear_sheet_non_formula(wb["gp_Cabecalhos"], r1=1, c1=2, r2=200, c2=2)
    if "gp_Plann_Sprint" in wb.sheetnames:
        ws = wb["gp_Plann_Sprint"]
        _clear_sheet_non_formula(ws, r1=1, c1=2, r2=ws.max_row, c2=ws.max_column)
    if "gp_Plann_Projeto" in wb.sheetnames:
        _clear_sheet_non_formula(wb["gp_Plann_Projeto"])
    if "gp_Disponibilidade" in wb.sheetnames:
        ws = wb["gp_Disponibilidade"]
        _clear_sheet_non_formula(ws, r1=7, c1=1, r2=ws.max_row, c2=max(4, ws.max_column))
    if "gp_Transversalidade" in wb.sheetnames:
        ws = wb["gp_Transversalidade"]
        _clear_sheet_non_formula(ws, r1=2, c1=1, r2=ws.max_row, c2=max(5, ws.max_column))
    if "gp_Roadmap" in wb.sheetnames:
        ws = wb["gp_Roadmap"]
        _clear_sheet_non_formula(ws, r1=2, c1=1, r2=ws.max_row, c2=max(2, ws.max_column))

    # 3) Legacy manual tabs (when present)
    for legacy_name in ("Planejamento", "Historico"):
        if legacy_name in wb.sheetnames:
            _clear_sheet_non_formula(wb[legacy_name])
    if "Roadmap" in wb.sheetnames:
        ws = wb["Roadmap"]
        _clear_sheet_non_formula(ws, r1=2, c1=1, r2=ws.max_row, c2=ws.max_column)

    # 4) Manual table area in Report and metadata
    if "Report" in wb.sheetnames:
        _clear_report_planning_tables(wb)
        _clear_report_roadmap_overlay_cells(wb)
        ws = wb["Report"]
        for row, col in ((2, 22), (2, 25)):  # V2, Y2
            try:
                ws.cell(row, col).value = None
            except Exception:
                pass

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    wb.save(output_template_path)

    # Restore original chart visuals from source template to avoid white backgrounds.
    _postprocess_xlsx(output_template_path, template_path=source_template_path)


# ═══════════════════════════════════════════════════════════════════════════════
#             AUTO-EXTRAÇÃO DA TAREFA GESTÃO E ROADMAP DO PLANNER
# ═══════════════════════════════════════════════════════════════════════════════

_DATE_LINE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}")

def build_gestao_meta(df):
    """
    Extrai PROJETO, GERENTE e LINKEDIN da tarefa 'Gestão' no export do Planner.
    Retorna dict com chaves: projeto, gerente, linkedin.
    """
    mask = df["tarefa"].fillna("").astype(str).str.strip().str.lower() == "gestão"
    if not mask.any():
        mask = df["tarefa"].fillna("").astype(str).str.strip().str.lower() == "gestao"
    if not mask.any():
        task_norm = df["tarefa"].fillna("").astype(str).map(_norm_text_key)
        mask = task_norm.isin({"gestao", "governanca", "governance"})
    if not mask.any() and "labels" in df.columns:
        label_norm = df["labels"].fillna("").astype(str).map(_norm_text_key)
        notes_norm = df["notas"].fillna("").astype(str).map(_norm_text_key)
        mask = label_norm.str.contains(r"(^|;)\.?gp($|;)", regex=True, na=False) & (
            notes_norm.str.contains("projeto", na=False)
        )
    if not mask.any():
        return {"projeto": "", "gerente": "", "linkedin": ""}

    desc = _strip(df.loc[mask, "notas"].iloc[0])
    result = {"projeto": "", "gerente": "", "linkedin": ""}
    for line in desc.split("\n"):
        line = line.strip()
        if ":" not in line:
            continue
        key_raw, _, val = line.partition(":")
        key_norm = _norm_text_key(key_raw).upper()
        val = val.strip()
        if key_norm == "PROJETO":
            result["projeto"] = val
        elif "GERENTE" in key_norm and "LINKEDIN" not in key_norm:
            result["gerente"] = val
        elif "LINKEDIN" in key_norm:
            result["linkedin"] = val
    return result


def build_roadmap_from_df(df):
    """
    Extrai os itens do roadmap da tarefa 'Roadmap' no export do Planner.
    Cada item = uma linha que começa com data DD/MM/YYYY + linha(s) de descrição.
    Retorna lista de (marco_text, posição).
    """
    POSITION_SEQ = [10, -10, 40, 25, 10, -40, -25, -10, 40, 25, 10, -40, -10,
                    40, 25, 10, -40, -25, -10, 40]

    mask = df["tarefa"].fillna("").astype(str).str.strip().str.lower() == "roadmap"
    if not mask.any():
        return []

    desc = _strip(df.loc[mask, "notas"].iloc[0])
    lines = desc.split("\n")

    # Agrupa linhas em blocos: cada bloco começa com uma data
    blocks = []
    current = []
    for line in lines:
        stripped = line.strip()
        if _DATE_LINE_RE.match(stripped):
            if current:
                blocks.append("\n".join(current))
            current = [stripped]
        elif stripped and current:
            current.append(stripped)
    if current:
        blocks.append("\n".join(current))

    result = []
    for i, text in enumerate(blocks):
        pos = POSITION_SEQ[i % len(POSITION_SEQ)]
        result.append((text, pos))
    return result


def _fill_roadmap_sheet(wb, roadmap_items):
    """
    Preenche a aba Roadmap com os itens extraídos do Planner.
    Estrutura esperada: A = Marco (texto completo), B = Posição (int).
    """
    if "Roadmap" not in wb.sheetnames:
        return
    ws = wb["Roadmap"]
    # Limpa dados (mantém header linha 1)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).value = None
        ws.cell(r, 2).value = None
    for i, (marco, pos) in enumerate(roadmap_items):
        r = i + 2
        _w(ws, r, 1, marco)
        _w(ws, r, 2, pos)


def read_template_effort(wb):
    """
    Lê a tabela de planejamento do Report (perfis × HUs) e a EsforcoHU,
    e retorna um dict {label_normalizado: esforco_total} para uso em
    PorCategoriaBubbles.

    Estrutura do Report (nova template v2):
      Linhas 8-18, col 7 = nome do perfil
      Cols 8+ = esforço por HU (valor inteiro ou 'N/A')

    Mapeamento perfil → rótulos (do Prompt Mestre):
      DEV FRONT.*DESKTOP  → DEV, FRONTEND, DESKTOP
      DEV FRONT.*MOBILE   → DEV, FRONTEND, MOBILE
      DEV BACK            → DEV, BACKEND
      DEV BANCO           → DEV, BACKEND
      DEV BI              → DEV
      UX.*DESKTOP         → UX, FRONTEND, DESKTOP
      UX.*MOBILE          → UX, FRONTEND, MOBILE
      TESTE               → Q/A TESTES
    """
    PROFILE_LABEL_MAP = [
        (re.compile(r"dev front.*desktop", re.I), ["DEV", "FRONTEND", "DESKTOP"]),
        (re.compile(r"dev front.*mobile",  re.I), ["DEV", "FRONTEND", "MOBILE"]),
        (re.compile(r"dev back\b",         re.I), ["DEV", "BACKEND"]),
        (re.compile(r"dev banco",          re.I), ["DEV", "BACKEND"]),
        (re.compile(r"dev bi\b",           re.I), ["DEV"]),
        (re.compile(r"ux.*desktop",        re.I), ["UX", "FRONTEND", "DESKTOP"]),
        (re.compile(r"ux.*mobile",         re.I), ["UX", "FRONTEND", "MOBILE"]),
        (re.compile(r"\bux\b",             re.I), ["UX"]),
        (re.compile(r"teste",              re.I), ["Q/A TESTES", "TESTES"]),
    ]

    label_effort = defaultdict(int)

    ws_rep = wb.get("Report") if hasattr(wb, "get") else (
        wb["Report"] if "Report" in wb.sheetnames else None
    )
    if ws_rep:
        for r in range(8, 19):
            profile = ws_rep.cell(r, 7).value
            if not profile:
                continue
            profile_str = str(profile).strip()
            # Soma esforço deste perfil por todos os HUs (cols 8 em diante)
            total = 0
            for c in range(8, 18):
                val = ws_rep.cell(r, c).value
                if val and str(val).strip() not in ("", "N/A", "nan", "None"):
                    try:
                        total += int(float(str(val)))
                    except (ValueError, TypeError):
                        pass
            if total == 0:
                continue
            for pattern, labels in PROFILE_LABEL_MAP:
                if pattern.search(profile_str):
                    for lbl in labels:
                        label_effort[_norm_label_key(lbl)] += total
                    break

    # EsforcoHU → HU labels
    ws_hu = None
    try:
        ws_hu = wb["EsforcoHU"]
    except (KeyError, TypeError):
        pass
    if ws_hu:
        for r in range(2, ws_hu.max_row + 1):
            hu_name = ws_hu.cell(r, 1).value
            total_val = ws_hu.cell(r, 3).value
            if not hu_name or not total_val:
                continue
            try:
                effort = int(float(str(total_val)))
            except (ValueError, TypeError):
                continue
            label_effort[_norm_label_key(str(hu_name).strip())] = effort
            # Also index by short HU ID (e.g. "HU052")
            m = re.match(r"(HU\s*\d+)", str(hu_name), re.I)
            if m:
                label_effort[_norm_label_key(m.group(1))] = effort

    return dict(label_effort)


# ═══════════════════════════════════════════════════════════════════════════════
#                     DADOS MANUAIS  (Opção B – legado, mantido para compat.)
# ═══════════════════════════════════════════════════════════════════════════════

# Mapeamento fixo das linhas de perfil no Report (col 5 = nome do perfil)
# Report rows 8-18 = 11 perfis; rows 19-20 = fórmulas (não sobrescrever)
REPORT_PROFILE_ROWS   = list(range(8, 19))   # rows 8..18 (11 perfis)
REPORT_PLAN_COL_START = 5   # col E: nome do perfil
REPORT_PLAN_COL_END   = 14  # col N: último HU (9 colunas de HU = F..N)
REPORT_HIST_PROFILE_COL  = 19   # col S: nome do perfil (historico)
REPORT_HIST_SPRINT_START = 20   # col T: primeira sprint
REPORT_HIST_SPRINT_END   = 27   # col AA: última sprint (8 sprints)
REPORT_HIST_ROWS         = list(range(8, 22))   # rows 8..21 (14 linhas)

GP_PLANN_PROJ_DEFAULT_LABELS = [
    "ID ou DESCRIÇÃO DO ITEM",
    "PO (DOCS, LINKS, LEIS, ETC.)",
    "REQUISITOS",
    "UX/UI  (DESKTOP)",
    "UX/UI  (MOBILE)",
    "BENCHMARK COMPONENTES - DEV",
    "CONSTRUCAO COMPONENTES - DEV",
    "TESTES- DEV",
    "REGISTRO e PUBLICAÇÃO (Storybook) - DEV",
    "",
    "",
    "",
    "Mensuração (maior esforço individual)",
    "Throughput - Qtd Storypoints",
    "Throughput - Qtd de Tarefas/Cards",
    "Qtd de Histórias de Usuário (HU)",
]

GP_PLANN_SPRINT_DEFAULT_LABELS = [
    "ID ou DESCRIÇÃO DO ITEM",
    "PO (DOCS, LINKS, LEIS, ETC.)",
    "REQUISITOS",
    "UX/UI  (DESKTOP)",
    "UX/UI  (MOBILE)",
    "BENCHMARK COMPONENTES - DEV",
    "CONSTRUCAO COMPONENTES - DEV",
    "TESTES- DEV",
    "REGISTRO e PUBLICAÇÃO (Storybook) - DEV",
    "",
    "",
    "",
    "Mensuração (maior esforço individual)",
    "Mensuracao soma total esforço por HU",
]


def _normalize_gp_plann_projeto_matrix(matrix):
    """
    Normaliza matriz de gp_Plann_Projeto para evitar deslocamentos de coluna.

    Casos tratados:
    1) Coluna extra de "Média..." no início -> remove coluna 1.
    2) Coluna de rótulos ausente (A1 já começa em "Escopo ...") -> recoloca
       a coluna de rótulos padrão na frente.
    """
    if not isinstance(matrix, list) or not matrix:
        return matrix

    safe_rows = [r for r in matrix if isinstance(r, list)]
    if not safe_rows:
        return matrix

    def _norm_txt(s):
        s = str(s or "").strip().lower()
        s = unicodedata.normalize("NFKD", s)
        return "".join(ch for ch in s if not unicodedata.combining(ch))

    def _looks_scope_header(s):
        return bool(re.search(
            r"escopo|sprint|continuo|jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez",
            _norm_txt(s),
        ))

    def _is_profile_label(s):
        return bool(re.search(
            r"po|requis|ux|dev|benchmark|construcao|teste|registro|mensur|throughput|historia",
            _norm_txt(s),
        ))

    first_header = _norm_txt(safe_rows[0][0] if len(safe_rows[0]) > 0 else "")
    second_header = _norm_txt(safe_rows[0][1] if len(safe_rows[0]) > 1 else "")
    looks_media_header = ("media" in first_header)
    looks_scope_header = _looks_scope_header(second_header)

    non_empty = 0
    numeric_first = 0
    label_second = 0
    profile_first = 0
    na_like_first = 0
    max_check = min(len(safe_rows), 24)
    for i in range(1, max_check):
        row = safe_rows[i]
        a = str(row[0]).strip() if len(row) > 0 and row[0] is not None else ""
        b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        if a or b:
            non_empty += 1
        if re.fullmatch(r"-?\d+([.,]\d+)?", a):
            numeric_first += 1
        if _is_profile_label(b):
            label_second += 1
        if _is_profile_label(a):
            profile_first += 1
        if _norm_txt(a) in {"n/a", "na", ""}:
            na_like_first += 1

    # Caso clássico: usuário colou bloco com coluna "Média..." à esquerda.
    # Segurança: só remove coluna quando a coluna "Média..." é explícita.
    looks_shifted_by_media = looks_media_header
    if looks_shifted_by_media:
        return [row[1:] if isinstance(row, list) and len(row) > 1 else ([] if isinstance(row, list) else row)
                for row in matrix]

    # Caso legado: coluna de rótulos já foi perdida (A1 começa em Escopo...).
    has_id_header = bool(re.search(r"\bid\b", first_header) and re.search(r"descr", first_header))
    looks_scope_first = _looks_scope_header(first_header)
    looks_scope_second = _looks_scope_header(second_header)
    looks_missing_label_col = (
        (not has_id_header)
        and looks_scope_first
        and looks_scope_second
        and non_empty >= 4
        and profile_first <= 1
        and (na_like_first + numeric_first) >= int((non_empty * 0.6) + 0.9999)
    )

    if looks_missing_label_col:
        rebuilt = []
        for i, row in enumerate(matrix):
            if not isinstance(row, list):
                rebuilt.append(row)
                continue
            label = GP_PLANN_PROJ_DEFAULT_LABELS[i] if i < len(GP_PLANN_PROJ_DEFAULT_LABELS) else ""
            rebuilt.append([label] + row)
        return rebuilt

    return matrix


def _normalize_gp_plann_projeto_sheet(ws):
    """
    Normaliza a aba gp_Plann_Projeto diretamente na planilha.
    """
    max_rows = 240
    max_cols = 40

    matrix = []
    for r in range(1, max_rows + 1):
        row = [ws.cell(r, c).value for c in range(1, max_cols + 1)]
        while row and (row[-1] is None or str(row[-1]).strip() == ""):
            row.pop()
        matrix.append(row)

    while matrix and not matrix[-1]:
        matrix.pop()

    normalized = _normalize_gp_plann_projeto_matrix(matrix)
    if normalized == matrix:
        return

    for r in range(1, max_rows + 1):
        for c in range(1, max_cols + 1):
            ws.cell(r, c).value = None

    for r_idx, row in enumerate(normalized, start=1):
        if not isinstance(row, list):
            continue
        for c_idx, val in enumerate(row, start=1):
            ws.cell(r_idx, c_idx).value = val


def _normalize_gp_plann_sprint_matrix(matrix):
    """
    Corrige gp_Plann_Sprint quando a coluna A foi "engolida".
    """
    if not isinstance(matrix, list) or not matrix:
        return matrix

    safe_rows = [r for r in matrix if isinstance(r, list)]
    if not safe_rows:
        return matrix

    def _norm_txt(s):
        s = str(s or "").strip().lower()
        s = unicodedata.normalize("NFKD", s)
        return "".join(ch for ch in s if not unicodedata.combining(ch))

    def _is_profile_label(s):
        return bool(re.search(
            r"po|requis|ux|dev|benchmark|construcao|teste|registro|mensur|throughput|historia",
            _norm_txt(s),
        ))

    first_header = _norm_txt(safe_rows[0][0] if len(safe_rows[0]) > 0 else "")
    second_header = _norm_txt(safe_rows[0][1] if len(safe_rows[0]) > 1 else "")
    has_id_header = bool(re.search(r"\bid\b", first_header) and re.search(r"descr", first_header))

    non_empty = 0
    profile_first = 0
    na_like_first = 0
    hu_header_like = 0
    max_check = min(len(safe_rows), 24)
    for i in range(1, max_check):
        row = safe_rows[i]
        a = str(row[0]).strip() if len(row) > 0 and row[0] is not None else ""
        b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        if a or b:
            non_empty += 1
        if _is_profile_label(a):
            profile_first += 1
        if _norm_txt(a) in {"n/a", "na", ""} or re.fullmatch(r"-?\d+([.,]\d+)?", a):
            na_like_first += 1
        if re.match(r"hu\s*\d+", _norm_txt(first_header)) or re.match(r"hu\s*\d+", _norm_txt(second_header)):
            hu_header_like = 1

    looks_missing_label_col = (
        (not has_id_header)
        and hu_header_like == 1
        and non_empty >= 4
        and profile_first <= 1
        and na_like_first >= int((non_empty * 0.6) + 0.9999)
    )
    if not looks_missing_label_col:
        return matrix

    rebuilt = []
    for i, row in enumerate(matrix):
        if not isinstance(row, list):
            rebuilt.append(row)
            continue
        label = GP_PLANN_SPRINT_DEFAULT_LABELS[i] if i < len(GP_PLANN_SPRINT_DEFAULT_LABELS) else ""
        rebuilt.append([label] + row)
    return rebuilt


def _normalize_gp_plann_sprint_sheet(ws):
    max_rows = 240
    max_cols = 40

    matrix = []
    for r in range(1, max_rows + 1):
        row = [ws.cell(r, c).value for c in range(1, max_cols + 1)]
        while row and (row[-1] is None or str(row[-1]).strip() == ""):
            row.pop()
        matrix.append(row)

    while matrix and not matrix[-1]:
        matrix.pop()

    normalized = _normalize_gp_plann_sprint_matrix(matrix)
    if normalized == matrix:
        return

    for r in range(1, max_rows + 1):
        for c in range(1, max_cols + 1):
            ws.cell(r, c).value = None

    for r_idx, row in enumerate(normalized, start=1):
        if not isinstance(row, list):
            continue
        for c_idx, val in enumerate(row, start=1):
            ws.cell(r_idx, c_idx).value = val


def extract_dados_manuais_starter(template_path, keep_formulas=False):
    """
    Extrai as 3 seções manuais do template e retorna um novo Workbook
    com sheets Roadmap, Planejamento e Historico prontos para o usuário editar.
    """
    from openpyxl import Workbook as _WB
    tpl = load_workbook(template_path, data_only=not keep_formulas)
    out = _WB()
    out.remove(out.active)

    # ── 1. Roadmap ────────────────────────────────────────────────────────────
    if "Roadmap" in tpl.sheetnames:
        ws_src = tpl["Roadmap"]
        ws_dst = out.create_sheet("Roadmap")
        for row in ws_src.iter_rows(values_only=True):
            ws_dst.append(list(row))

    # ── 2. Planejamento (Report cols E-N, rows 6-18) ──────────────────────────
    ws_dst = out.create_sheet("Planejamento")
    if "Report" in tpl.sheetnames:
        ws_rep = tpl["Report"]
        # rows 6-7 = cabeçalhos (PRIORIDADE + HU names)
        # rows 8-18 = perfis com storypoints
        for src_row in range(6, 19):
            row_data = []
            for col in range(REPORT_PLAN_COL_START, REPORT_PLAN_COL_END + 1):
                v = ws_rep.cell(src_row, col).value
                # No starter removemos fórmulas para facilitar edição manual.
                # Na opção A (report anterior), mantemos fórmulas quando solicitado.
                if keep_formulas:
                    row_data.append(v)
                else:
                    row_data.append(None if isinstance(v, str) and v.startswith("=") else v)
            ws_dst.append(row_data)

    # ── 3. Historico (Report cols S-AA, rows 6-21) ───────────────────────────
    ws_dst = out.create_sheet("Historico")
    if "Report" in tpl.sheetnames:
        ws_rep = tpl["Report"]
        # rows 6-7 = cabeçalhos (Sprint labels + períodos)
        # rows 8-21 = perfis com pontuação histórica
        for src_row in range(6, 22):
            row_data = []
            # col S (19) = nome do perfil
            for col in range(REPORT_HIST_PROFILE_COL, REPORT_HIST_SPRINT_END + 1):
                v = ws_rep.cell(src_row, col).value
                if keep_formulas:
                    row_data.append(v)
                else:
                    row_data.append(None if isinstance(v, str) and v.startswith("=") else v)
            ws_dst.append(row_data)

    return out


def inject_dados_manuais(wb, dados_manuais_path):
    """
    Injeta as 3 abas do arquivo de dados manuais no workbook de saída.
    Sheets esperadas: Roadmap, Planejamento, Historico.
    """
    dm = load_workbook(dados_manuais_path)

    def _clear_rect(ws, r1, r2, c1, c2):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).value = None

    # ── 1. Roadmap ────────────────────────────────────────────────────────────
    if "Roadmap" in dm.sheetnames and "Roadmap" in wb.sheetnames:
        ws_src = dm["Roadmap"]
        ws_dst = wb["Roadmap"]
        # limpa linhas de dados (mantém linha 1 = cabeçalho do template)
        _clear_rect(ws_dst, 1, ws_dst.max_row, 1, ws_dst.max_column)
        # copia tudo da fonte
        for r_idx, row in enumerate(ws_src.iter_rows(values_only=True), 1):
            for c_idx, val in enumerate(row, 1):
                ws_dst.cell(r_idx, c_idx).value = val

    # ── 2. Planejamento → Report cols E-N, rows 6-18 ─────────────────────────
    if "Planejamento" in dm.sheetnames and "Report" in wb.sheetnames:
        ws_src = dm["Planejamento"]
        ws_dst = wb["Report"]
        _clear_rect(ws_dst, 6, 18, REPORT_PLAN_COL_START, REPORT_PLAN_COL_END)
        # Planejamento row 1 → Report row 6 (PRIORIDADE / HU names header)
        # Planejamento row 2 → Report row 7 (HU descriptions)
        # Planejamento rows 3-13 → Report rows 8-18 (11 perfis)
        for local_r, report_r in enumerate(range(6, 19), 1):
            if local_r > ws_src.max_row:
                break
            for local_c, report_c in enumerate(
                    range(REPORT_PLAN_COL_START, REPORT_PLAN_COL_END + 1), 1):
                val = ws_src.cell(local_r, local_c).value
                ws_dst.cell(report_r, report_c).value = val

    # ── 3. Historico → Report cols S-AA, rows 6-21 ───────────────────────────
    if "Historico" in dm.sheetnames and "Report" in wb.sheetnames:
        ws_src = dm["Historico"]
        ws_dst = wb["Report"]
        _clear_rect(ws_dst, 6, 21, REPORT_HIST_PROFILE_COL, REPORT_HIST_SPRINT_END)
        # Historico row 1 → Report row 6 (Sprint labels)
        # Historico row 2 → Report row 7 (períodos)
        # Historico rows 3-16 → Report rows 8-21 (perfis + totais)
        for local_r, report_r in enumerate(range(6, 22), 1):
            if local_r > ws_src.max_row:
                break
            for local_c, report_c in enumerate(
                    range(REPORT_HIST_PROFILE_COL, REPORT_HIST_SPRINT_END + 1), 1):
                val = ws_src.cell(local_r, local_c).value
                ws_dst.cell(report_r, report_c).value = val


# ═══════════════════════════════════════════════════════════════════════════════
#                    ZIP-LEVEL POST-PROCESSING  (chartEx injection)
# ═══════════════════════════════════════════════════════════════════════════════

def inject_dados_manuais_ext(wb, dados_manuais_path):
    """
    Injeta dados manuais no workbook de saida.

    Suporta:
    - Legado: Roadmap, Planejamento, Historico
    - gp_*: gp_Roadmap, gp_Plann_Sprint, gp_Plann_Projeto,
      gp_Disponibilidade, gp_Transversalidade, gp_Cabecalhos

    Retorno:
      {"roadmap": bool, "report_plan_hist": bool, "gp_tabs": [..]}
    """
    dm = load_workbook(dados_manuais_path)
    applied_gp_tabs = set()
    applied_roadmap = False
    applied_report_plan_hist = False

    def _clear_rect(ws, r1, r2, c1, c2):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).value = None

    def _copy_sheet_all(ws_src, ws_dst):
        for r in range(1, ws_dst.max_row + 1):
            for c in range(1, ws_dst.max_column + 1):
                try:
                    ws_dst.cell(r, c).value = None
                except Exception:
                    pass
        for r_idx, row in enumerate(ws_src.iter_rows(values_only=True), 1):
            for c_idx, val in enumerate(row, 1):
                try:
                    ws_dst.cell(r_idx, c_idx).value = val
                except Exception:
                    pass

    for sheet_name in (
        "gp_Plann_Projeto",
        "gp_Plann_Sprint",
        "gp_Cabecalhos",
        "gp_Disponibilidade",
        "gp_Transversalidade",
        "gp_Roadmap",
    ):
        if sheet_name in dm.sheetnames and sheet_name in wb.sheetnames:
            _copy_sheet_all(dm[sheet_name], wb[sheet_name])
            if sheet_name == "gp_Plann_Projeto":
                _normalize_gp_plann_projeto_sheet(wb[sheet_name])
            if sheet_name == "gp_Plann_Sprint":
                _normalize_gp_plann_sprint_sheet(wb[sheet_name])
            applied_gp_tabs.add(sheet_name)
            if sheet_name == "gp_Roadmap":
                applied_roadmap = True

    if "Roadmap" in dm.sheetnames and "Roadmap" in wb.sheetnames:
        _copy_sheet_all(dm["Roadmap"], wb["Roadmap"])
        applied_roadmap = True

    if "Roadmap" in dm.sheetnames and "gp_Roadmap" in wb.sheetnames and "gp_Roadmap" not in applied_gp_tabs:
        _copy_sheet_all(dm["Roadmap"], wb["gp_Roadmap"])
        applied_gp_tabs.add("gp_Roadmap")
        applied_roadmap = True

    if "Planejamento" in dm.sheetnames and "Report" in wb.sheetnames:
        ws_src = dm["Planejamento"]
        ws_dst = wb["Report"]
        _clear_rect(ws_dst, 6, 18, REPORT_PLAN_COL_START, REPORT_PLAN_COL_END)
        for local_r, report_r in enumerate(range(6, 19), 1):
            if local_r > ws_src.max_row:
                break
            for local_c, report_c in enumerate(range(REPORT_PLAN_COL_START, REPORT_PLAN_COL_END + 1), 1):
                ws_dst.cell(report_r, report_c).value = ws_src.cell(local_r, local_c).value
        applied_report_plan_hist = True

    if "Historico" in dm.sheetnames and "Report" in wb.sheetnames:
        ws_src = dm["Historico"]
        ws_dst = wb["Report"]
        _clear_rect(ws_dst, 6, 21, REPORT_HIST_PROFILE_COL, REPORT_HIST_SPRINT_END)
        for local_r, report_r in enumerate(range(6, 22), 1):
            if local_r > ws_src.max_row:
                break
            for local_c, report_c in enumerate(range(REPORT_HIST_PROFILE_COL, REPORT_HIST_SPRINT_END + 1), 1):
                ws_dst.cell(report_r, report_c).value = ws_src.cell(local_r, local_c).value
        applied_report_plan_hist = True

    # IMPORTANTE:
    # Nao copiar a aba Report inteira do arquivo manual.
    # Isso reintroduz dados/graficos antigos e gera sobreposicoes no dashboard final.
    # O fluxo correto e copiar apenas:
    #   - abas gp_* (entrada manual)
    #   - legado Planejamento/Historico (quando usados explicitamente)
    # Portanto, mesmo que o arquivo manual tenha "Report", ela e ignorada aqui.

    return {
        "roadmap": applied_roadmap,
        "report_plan_hist": applied_report_plan_hist,
        "gp_tabs": sorted(applied_gp_tabs),
    }



def inject_dados_formulario(wb, form_data):
    """
    Injeta dados do formulario web nas abas gp_* do workbook.

    form_data: dict com chaves opcionais:
      cabecalhos      -> gp_Cabecalhos (PROJETO, GERENTE, LINKEDIN, PRODUCT_GOAL, SPRINT_GOAL)
      plann_projeto   -> gp_Plann_Projeto (matrix: list[list[str]])
      plann_sprint    -> gp_Plann_Sprint (hus: list[str], values: list[list[str]])
      disponibilidade -> gp_Disponibilidade (list de {situacao, empregador, nome, pct})
      transversalidade-> gp_Transversalidade (list de {nome, funcao, squad, intervalo})
      roadmap         -> gp_Roadmap (list de {marco, posicao})

    Retorna lista de nomes de abas que foram preenchidas.
    """
    filled = []

    def _clear_non_formula(ws):
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    continue
                try:
                    cell.value = None
                except Exception:
                    pass

    def _coerce_cell(v):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return v
        s = str(v).strip()
        if s == "":
            return None
        num = s.replace(".", "").replace(",", ".") if "," in s else s
        if re.fullmatch(r"-?\d+(\.\d+)?", num):
            try:
                f = float(num)
                return int(f) if abs(f - int(f)) < 1e-9 else f
            except Exception:
                return s
        return s

    def _normalize_plann_projeto_matrix(matrix):
        """
        Corrige matriz colada com coluna extra de "Média..." no início.
        Esperado em gp_Plann_Projeto:
          col A = perfil/label
          col B.. = sprints/valores
        """
        if not isinstance(matrix, list) or not matrix:
            return matrix

        safe_rows = [r for r in matrix if isinstance(r, list) and len(r) > 0]
        if not safe_rows:
            return matrix

        def _norm_txt(s):
            s = str(s or "").strip().lower()
            s = unicodedata.normalize("NFKD", s)
            return "".join(ch for ch in s if not unicodedata.combining(ch))

        first_header = _norm_txt(safe_rows[0][0] if len(safe_rows[0]) > 0 else "")
        second_header = _norm_txt(safe_rows[0][1] if len(safe_rows[0]) > 1 else "")
        looks_media_header = ("media" in first_header)
        looks_scope_header = bool(re.search(r"escopo|sprint|continuo|jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez", second_header))

        non_empty = 0
        numeric_first = 0
        label_second = 0
        max_check = min(len(safe_rows), 16)
        for i in range(1, max_check):
            row = safe_rows[i]
            a = str(row[0]).strip() if len(row) > 0 and row[0] is not None else ""
            b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            if a or b:
                non_empty += 1
            if re.fullmatch(r"-?\d+([.,]\d+)?", a):
                numeric_first += 1
            if re.search(r"po|requis|ux|dev|teste|arquitet|mensur|throughput|historia", _norm_txt(b)):
                label_second += 1

        looks_shifted = looks_media_header or (
            looks_scope_header and non_empty >= 4 and
            numeric_first >= int((non_empty * 0.6) + 0.9999) and
            label_second >= 2
        )
        if not looks_shifted:
            return matrix

        return [row[1:] if isinstance(row, list) else row for row in matrix]

    # ── gp_Cabecalhos ─────────────────────────────────────────────────────────
    if "cabecalhos" in form_data and "gp_Cabecalhos" in wb.sheetnames:
        c = form_data["cabecalhos"]
        ws = wb["gp_Cabecalhos"]
        cab_rows = [
            ("PROJETO",                                        c.get("projeto", "")),
            ("NOME DO GERENTE DO PROJETO",                     c.get("gerente", "")),
            ("link para LINKEDIN DO GERENTE DO PROJETO",       c.get("linkedin", "")),
            ("PRODUCT GOAL (OBJETIVO DO PROJETO)",             c.get("product_goal", "")),
            ("SPRINT GOAL (OBJETIVO DA SPRINT)",               c.get("sprint_goal", "")),
        ]
        for i, (lbl, val) in enumerate(cab_rows, start=1):
            _w(ws, i, 1, lbl)
            if val:  # Nao sobrescreve com vazio
                _w(ws, i, 2, val)
        filled.append("gp_Cabecalhos")

    # ── gp_Plann_Projeto ──────────────────────────────────────────────────────
    if "plann_projeto" in form_data and "gp_Plann_Projeto" in wb.sheetnames:
        pp = form_data["plann_projeto"] or {}
        matrix = pp.get("matrix", [])
        if isinstance(matrix, list) and matrix:
            matrix = _normalize_gp_plann_projeto_matrix(matrix)
            ws = wb["gp_Plann_Projeto"]
            _clear_non_formula(ws)
            for r_idx, row in enumerate(matrix, start=1):
                if not isinstance(row, list):
                    continue
                for c_idx, raw in enumerate(row, start=1):
                    _w(ws, r_idx, c_idx, _coerce_cell(raw))
            filled.append("gp_Plann_Projeto")

    # ── gp_Plann_Sprint ───────────────────────────────────────────────────────
    if "plann_sprint" in form_data and "gp_Plann_Sprint" in wb.sheetnames:
        ps = form_data["plann_sprint"]
        ws = wb["gp_Plann_Sprint"]
        _ROLES = [
            "PO (DOCS, LINKS, LEIS, ETC.)",
            "REQUISITOS",
            "UX/UI  (DESKTOP)",
            "UX/UI  (MOBILE)",
            "DEV FRONT (DESKTOP)",
            "DEV FRONT (MOBILE)",
            "DEV BACK",
            "DEV BANCO DE DADOS",
            "DEV BI",
            "TESTE (possíveis no prazo da sprint)",
            "ARQUITETURA",
        ]
        hus = ps.get("hus", [])
        values = ps.get("values")
        if values is None:
            values = []
            rows_payload = ps.get("rows", [])
            if isinstance(rows_payload, list):
                for row in rows_payload:
                    if isinstance(row, dict):
                        vals = row.get("valores")
                        if vals is None:
                            vals = row.get("values", [])
                        values.append(vals if isinstance(vals, list) else [])
                    else:
                        values.append([])
        _w(ws, 1, 1, "ID ou DESCRIÇÃO DO ITEM")
        for j, hu in enumerate(hus, start=2):
            _w(ws, 1, j, hu)
        for i, role in enumerate(_ROLES):
            row_vals = values[i] if i < len(values) else []
            _w(ws, i + 2, 1, role)
            for j, val in enumerate(row_vals):
                _w(ws, i + 2, j + 2, val if str(val).strip() else "N/A")
        filled.append("gp_Plann_Sprint")

    # ── gp_Disponibilidade ────────────────────────────────────────────────────
    if "disponibilidade" in form_data and "gp_Disponibilidade" in wb.sheetnames:
        rows = form_data["disponibilidade"]
        ws = wb["gp_Disponibilidade"]
        _w(ws, 6, 1, "SITUACAO")
        _w(ws, 6, 2, "EMPREGADOR")
        _w(ws, 6, 3, "NOME")
        _w(ws, 6, 4, "Percentual Disponivel")
        # Limpa area de dados mantendo formulas do template fora de A:D.
        for r in range(8, 501):
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)

        # Dados iniciam na linha 8 (linha 7 e reservada no template).
        for i, row in enumerate(rows, start=8):
            _w(ws, i, 1, row.get("situacao", "titular"))
            _w(ws, i, 2, row.get("empregador", ""))
            _w(ws, i, 3, row.get("nome", ""))
            try:
                pct_raw = row.get("pct", row.get("percentual", row.get("percent", None)))
                _w(ws, i, 4, _parse_percent_fraction(pct_raw, default_value=1.0))
            except (ValueError, TypeError):
                _w(ws, i, 4, 1.0)
        filled.append("gp_Disponibilidade")

    # ── gp_Transversalidade ───────────────────────────────────────────────────
    if "transversalidade" in form_data and "gp_Transversalidade" in wb.sheetnames:
        rows = form_data["transversalidade"]
        ws = wb["gp_Transversalidade"]
        _w(ws, 1, 1, "NOME")
        _w(ws, 1, 2, "FUNÇÃO")
        _w(ws, 1, 3, "SQUAD")
        _w(ws, 1, 4, "INTERVALO")
        _w(ws, 1, 5, "coluna")
        for i, row in enumerate(rows, start=2):
            _w(ws, i, 1, row.get("nome", ""))
            _w(ws, i, 2, row.get("funcao", ""))
            _w(ws, i, 3, row.get("squad", ""))
            _w(ws, i, 4, row.get("intervalo", ""))
            _w(ws, i, 5, 1)
        filled.append("gp_Transversalidade")

    # ── gp_Roadmap ────────────────────────────────────────────────────────────
    if "roadmap" in form_data and "gp_Roadmap" in wb.sheetnames:
        items = form_data["roadmap"]
        ws = wb["gp_Roadmap"]
        _w(ws, 1, 1, "Marco")
        _w(ws, 1, 2, "Posição")
        for i, item in enumerate(items, start=2):
            _w(ws, i, 1, item.get("marco", ""))
            try:
                _w(ws, i, 2, float(str(item.get("posicao", "0")).replace(",", ".")))
            except (ValueError, TypeError):
                _w(ws, i, 2, 0)
        filled.append("gp_Roadmap")

    return filled


def _parse_percent_fraction(raw, default_value=1.0):
    """
    Converte diferentes formatos de percentual para fração:
      100   -> 1.0
      100%  -> 1.0
      0.8   -> 0.8
      80,5% -> 0.805
    """
    if raw is None:
        return default_value
    txt = str(raw).strip()
    if txt == "":
        return default_value
    has_pct = "%" in txt
    txt = txt.replace("%", "").replace(" ", "").replace(",", ".")
    val = float(txt)
    if has_pct or abs(val) > 1.0:
        val = val / 100.0
    return val


def _normalize_gp_disponibilidade_sheet(ws):
    """
    Normaliza percentuais em gp_Disponibilidade!D8:D* para fração (0..1),
    preservando linhas totalmente vazias.
    """
    for r in range(8, 501):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        d = ws.cell(r, 4).value
        has_row = any(
            v is not None and str(v).strip() != ""
            for v in (a, b, c, d)
        )
        if not has_row:
            continue
        try:
            ws.cell(r, 4).value = _parse_percent_fraction(d, default_value=1.0)
        except Exception:
            ws.cell(r, 4).value = 1.0


def _postprocess_xlsx(output_path: str, template_path: str = None) -> None:
    """
    Pós-processamento MINIMAL do xlsx via manipulação direta do ZIP.

    Estratégia (cirúrgica — só toca o que é necessário):
      - NÃO substitui drawing1.xml (openpyxl gera um válido para charts normais).
      - NÃO injeta chart .rels ou style/colors de charts NORMAIS (openpyxl funciona sem eles).
      - SOMENTE injeta arquivos necessários para chartEx (sunburst / tree-map):
        * chartEx1.xml, chartEx2.xml (os gráficos em si)
        * chartEx1.xml.rels, chartEx2.xml.rels (referências de style/colors dos chartEx)
        * style/colors referenciados PELOS chartEx (style6, colors6, style11, colors11)
      - Adiciona anchors chartEx ao drawing1.xml do openpyxl (append, não replace).
      - Adiciona refs no drawing1.xml.rels e Content_Types.
      - Limpa cache do gráfico Roadmap (chart10).
    """
    import re as _re

    # Fast-path: preserve chart visuals exactly as in template (including transparency).
    if template_path and os.path.exists(template_path):
        tmp_visual = output_path + ".tmp_postproc_visual"
        try:
            replace_map = {}
            with zipfile.ZipFile(template_path, "r") as ztpl:
                tpl_names = set(ztpl.namelist())

                # Drawing + normal chart parts
                for name in tpl_names:
                    is_drawing = name in {
                        "xl/drawings/drawing1.xml",
                        "xl/drawings/_rels/drawing1.xml.rels",
                    }
                    is_chart = (
                        name.startswith("xl/charts/chart")
                        and name.endswith(".xml")
                        and "chartEx" not in name
                    )
                    is_chart_rel = (
                        name.startswith("xl/charts/_rels/chart")
                        and name.endswith(".rels")
                        and "chartEx" not in name
                    )
                    if is_drawing or is_chart or is_chart_rel:
                        replace_map[name] = ztpl.read(name)

                # chartEx + dependencies
                chart_ex_rels = []
                for name in tpl_names:
                    if name.startswith("xl/charts/chartEx") and name.endswith(".xml"):
                        replace_map[name] = ztpl.read(name)
                    elif name.startswith("xl/charts/_rels/chartEx") and name.endswith(".rels"):
                        rel_bytes = ztpl.read(name)
                        replace_map[name] = rel_bytes
                        chart_ex_rels.append(rel_bytes.decode("utf-8", errors="ignore"))

                for rel_xml in chart_ex_rels:
                    for dep in _re.findall(r'Target="([^"]+)"', rel_xml):
                        dep_path = f"xl/charts/{dep}"
                        if dep_path in tpl_names:
                            replace_map[dep_path] = ztpl.read(dep_path)

                # Normal chart style/color dependencies used by chart rels.
                for name in tpl_names:
                    if name.startswith("xl/charts/style") and name.endswith(".xml"):
                        replace_map[name] = ztpl.read(name)
                    elif name.startswith("xl/charts/colors") and name.endswith(".xml"):
                        replace_map[name] = ztpl.read(name)

            if replace_map:
                with zipfile.ZipFile(output_path, "r") as zin, \
                     zipfile.ZipFile(tmp_visual, "w", zipfile.ZIP_DEFLATED) as zout:

                    existing_names = set(zin.namelist())
                    for item in zin.infolist():
                        name = item.filename
                        data = replace_map.get(name, zin.read(name))

                        if name == "[Content_Types].xml":
                            xml = data.decode("utf-8", errors="ignore")
                            additions = ""
                            for part_name in replace_map.keys():
                                if not part_name.startswith("xl/charts/") or "/_rels/" in part_name:
                                    continue
                                base = part_name.split("/")[-1]
                                if base.startswith("chartEx") and base.endswith(".xml"):
                                    ctype = "application/vnd.ms-office.chartex+xml"
                                elif base.startswith("style") and base.endswith(".xml"):
                                    ctype = "application/vnd.ms-office.chartstyle+xml"
                                elif base.startswith("colors") and base.endswith(".xml"):
                                    ctype = "application/vnd.ms-office.chartcolorstyle+xml"
                                elif base.startswith("chart") and base.endswith(".xml"):
                                    ctype = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
                                else:
                                    continue
                                override = f'/xl/charts/{base}'
                                if override not in xml:
                                    additions += (
                                        f'<Override PartName="{override}" '
                                        f'ContentType="{ctype}"/>'
                                    )
                            if additions:
                                xml = xml.replace("</Types>", additions + "</Types>")
                            data = xml.encode("utf-8")

                        if (
                            name.startswith("xl/charts/chart")
                            and name.endswith(".xml")
                            and "chartEx" not in name
                        ):
                            xml = data.decode("utf-8", errors="ignore")
                            xml = _re.sub(
                                r'<c:strCache>.*?</c:strCache>',
                                '<c:strCache><c:ptCount val="0"/></c:strCache>',
                                xml,
                                flags=_re.DOTALL,
                            )
                            xml = _re.sub(
                                r'<c:numCache>.*?</c:numCache>',
                                '<c:numCache><c:formatCode>General</c:formatCode>'
                                '<c:ptCount val="0"/></c:numCache>',
                                xml,
                                flags=_re.DOTALL,
                            )
                            data = xml.encode("utf-8")

                        zout.writestr(item, data)

                    for part_name, part_data in replace_map.items():
                        if part_name not in existing_names:
                            zout.writestr(part_name, part_data)

                os.replace(tmp_visual, output_path)
                print("      [OK] visual dos gráficos preservado do template + caches de séries limpos.")
                return
        except Exception as _preserve_err:
            if os.path.exists(tmp_visual):
                os.remove(tmp_visual)
            print(f"      [aviso] preservação visual falhou, usando fallback: {_preserve_err}")

    # ── 1. Extrai SOMENTE assets de chartEx do template ──────────────────────
    chartex_assets = {}    # zip_path -> bytes  (SOMENTE chartEx e seus deps)
    chartex_rids = []      # lista de (rId_original, target_relativo) do template
    chartex_deps = set()   # style/colors que os chartEx referenciam

    if template_path and os.path.exists(template_path):
        try:
            with zipfile.ZipFile(template_path, "r") as ztpl:
                tpl_names = set(ztpl.namelist())

                # Descobre rIds dos chartEx no drawing1.xml.rels do template
                if "xl/drawings/_rels/drawing1.xml.rels" in tpl_names:
                    _rels_xml = ztpl.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
                    for m in _re.finditer(
                        r'Id="(rId\d+)"\s+Type="[^"]*chartEx[^"]*"\s+Target="([^"]*)"', _rels_xml
                    ):
                        chartex_rids.append((m.group(1), m.group(2)))

                # Extrai os chartEx XML
                for _zn in tpl_names:
                    if _zn.startswith("xl/charts/chartEx") and _zn.endswith(".xml"):
                        chartex_assets[_zn] = ztpl.read(_zn)

                # Extrai os .rels dos chartEx e descobre quais style/colors eles usam
                for _zn in tpl_names:
                    if _zn.startswith("xl/charts/_rels/chartEx") and _zn.endswith(".rels"):
                        rels_data = ztpl.read(_zn)
                        chartex_assets[_zn] = rels_data
                        # Parse para descobrir dependências
                        for dep_m in _re.finditer(r'Target="([^"]+)"', rels_data.decode("utf-8")):
                            dep_target = dep_m.group(1)  # ex: "style6.xml", "colors6.xml"
                            dep_path = f"xl/charts/{dep_target}"
                            chartex_deps.add(dep_path)

                # Extrai SOMENTE os style/colors referenciados pelos chartEx
                for dep_path in chartex_deps:
                    if dep_path in tpl_names:
                        chartex_assets[dep_path] = ztpl.read(dep_path)

                # Extrai as anchors de chartEx do template drawing1.xml (posição original)
                if "xl/drawings/drawing1.xml" in tpl_names:
                    _tpl_drawing = ztpl.read("xl/drawings/drawing1.xml").decode("utf-8")
                    chartex_assets["__tpl_drawing__"] = _tpl_drawing.encode("utf-8")

            if chartex_assets:
                print(f"      [OK] {len(chartex_assets)} asset(s) chartEx extraído(s) do template.")
                if chartex_rids:
                    print(f"      [OK] chartEx encontrados: {[r[0] for r in chartex_rids]}")
                if chartex_deps:
                    print(f"      [OK] Dependências chartEx: {sorted(chartex_deps)}")
        except Exception as _e:
            print(f"      [aviso] Não foi possível ler assets do template: {_e}")
            chartex_assets = {}

    has_any_chartex = bool(chartex_rids)

    if not has_any_chartex:
        print("      [aviso] chartEx não encontrado no template – gráficos chartEx não injetados.")

    # ── 2. Processa o ZIP ─────────────────────────────────────────────────────
    tmp_path = output_path + ".tmp_postproc"
    chartex_rid_map = {}  # target -> novo rId (definido no escopo externo para drawing1.xml.rels)

    try:
        with zipfile.ZipFile(output_path, "r") as zin, \
             zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:

            existing_names = set(zin.namelist())

            # Monta mapa de injeção: SOMENTE chartEx e suas deps
            inject_map = {}
            for _tn, _td in chartex_assets.items():
                if _tn.startswith("__"):  # skip metadados internos
                    continue
                if _tn not in existing_names:
                    inject_map[_tn] = _td

            for item in zin.infolist():
                name = item.filename
                data = zin.read(name)

                # ── drawing1.xml: APPEND anchors dos chartEx ──────────────────
                if name == "xl/drawings/drawing1.xml" and has_any_chartex:
                    xml = data.decode("utf-8")
                    uses_prefix = "xdr:wsDr" in xml
                    xdr = "xdr:" if uses_prefix else ""
                    close_tag = f"</{xdr}wsDr>"

                    # Descobre o maior rId no .rels do openpyxl para não colidir
                    _rels_data = zin.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8") \
                        if "xl/drawings/_rels/drawing1.xml.rels" in existing_names else ""
                    _existing_rids = [int(x) for x in _re.findall(r'rId(\d+)', _rels_data)]
                    _next_rid = max(_existing_rids) + 1 if _existing_rids else 30

                    # Descobre o maior shape id no drawing
                    _existing_ids = [int(x) for x in _re.findall(r'<(?:\w+:)?cNvPr id="(\d+)"', xml)]
                    _next_id = max(_existing_ids) + 1 if _existing_ids else 100

                    # Mapeia chartEx targets para novos rIds
                    for _orig_rid, tgt in chartex_rids:
                        chartex_rid_map[tgt] = f"rId{_next_rid}"
                        _next_rid += 1

                    # Tenta extrair as anchors originais do template para preservar posição
                    tpl_drawing = chartex_assets.get("__tpl_drawing__", b"").decode("utf-8")
                    tpl_anchors = {}  # orig_rId -> anchor_xml
                    if tpl_drawing:
                        for am in _re.finditer(
                            r'<xdr:twoCellAnchor>(.*?)</xdr:twoCellAnchor>',
                            tpl_drawing, _re.DOTALL
                        ):
                            block = am.group(0)
                            rid_match = _re.search(r'r:id="(rId\d+)"', block)
                            if rid_match and rid_match.group(1) in {r[0] for r in chartex_rids}:
                                tpl_anchors[rid_match.group(1)] = block

                    for _orig_rid, tgt in chartex_rids:
                        new_rid = chartex_rid_map[tgt]
                        label = tgt.split("/")[-1].replace(".xml", "")

                        if _orig_rid in tpl_anchors:
                            # Usa anchor do template com posição original, troca rId e id
                            anchor_xml = tpl_anchors[_orig_rid]
                            # Troca o rId interno para o novo
                            anchor_xml = anchor_xml.replace(
                                f'r:id="{_orig_rid}"', f'r:id="{new_rid}"')
                            # Troca xdr: prefix se necessário (template usa xdr:, openpyxl não)
                            if not uses_prefix:
                                anchor_xml = _re.sub(r'xdr:', '', anchor_xml)
                                anchor_xml = anchor_xml.replace('<twoCellAnchor>', '<twoCellAnchor>')
                            # Atualiza shape ids para não colidir
                            for old_id in _re.findall(r'<(?:\w+:)?cNvPr id="(\d+)"', anchor_xml):
                                old_id_int = int(old_id)
                                if old_id_int > 0:  # Preserva id="0" no Fallback
                                    anchor_xml = anchor_xml.replace(
                                        f'id="{old_id}"', f'id="{_next_id}"', 1)
                                    _next_id += 1
                            anchor = anchor_xml
                        else:
                            # Fallback: gera anchor minimal
                            anchor = (
                                f'<{xdr}twoCellAnchor>'
                                f'<{xdr}from><{xdr}col>0</{xdr}col><{xdr}colOff>0</{xdr}colOff>'
                                f'<{xdr}row>0</{xdr}row><{xdr}rowOff>0</{xdr}rowOff></{xdr}from>'
                                f'<{xdr}to><{xdr}col>5</{xdr}col><{xdr}colOff>0</{xdr}colOff>'
                                f'<{xdr}row>5</{xdr}row><{xdr}rowOff>0</{xdr}rowOff></{xdr}to>'
                                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                                '<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">'
                                f'<{xdr}graphicFrame macro="">'
                                f'<{xdr}nvGraphicFramePr>'
                                f'<{xdr}cNvPr id="{_next_id}" name="{label}"/>'
                                f'<{xdr}cNvGraphicFramePr/>'
                                f'</{xdr}nvGraphicFramePr>'
                                f'<{xdr}xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></{xdr}xfrm>'
                                '<a:graphic>'
                                '<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">'
                                '<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
                                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                                f'r:id="{new_rid}"/>'
                                '</a:graphicData>'
                                '</a:graphic>'
                                f'</{xdr}graphicFrame>'
                                '</mc:Choice>'
                                '</mc:AlternateContent>'
                                f'<{xdr}clientData/>'
                                f'</{xdr}twoCellAnchor>'
                            )
                            _next_id += 1

                        xml = xml.replace(close_tag, anchor + close_tag)
                    data = xml.encode("utf-8")

                # ── drawing1.xml.rels: adiciona refs para chartEx ─────────────
                elif name == "xl/drawings/_rels/drawing1.xml.rels" and has_any_chartex:
                    xml = data.decode("utf-8")
                    for tgt, new_rid in chartex_rid_map.items():
                        if new_rid not in xml:
                            xml = xml.replace("</Relationships>",
                                f'<Relationship Id="{new_rid}" '
                                'Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" '
                                f'Target="{tgt}"/>'
                                '</Relationships>')
                    data = xml.encode("utf-8")

                # ── [Content_Types].xml: registra SOMENTE chartEx + deps ──────
                elif name == "[Content_Types].xml" and has_any_chartex:
                    xml = data.decode("utf-8")
                    additions = ""
                    # chartEx content types
                    for _tn in inject_map:
                        bn = _tn.split("/")[-1]
                        if bn.startswith("chartEx") and bn.endswith(".xml") and bn not in xml:
                            additions += (
                                f'<Override PartName="/xl/charts/{bn}" '
                                'ContentType="application/vnd.ms-office.chartex+xml"/>')
                        elif bn.startswith("style") and bn.endswith(".xml") and bn not in xml:
                            additions += (
                                f'<Override PartName="/xl/charts/{bn}" '
                                'ContentType="application/vnd.ms-office.chartstyle+xml"/>')
                        elif bn.startswith("colors") and bn.endswith(".xml") and bn not in xml:
                            additions += (
                                f'<Override PartName="/xl/charts/{bn}" '
                                'ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>')
                    if additions:
                        xml = xml.replace("</Types>", additions + "</Types>")
                    data = xml.encode("utf-8")

                # ── Limpa caches das séries dos charts normais ───────────────
                elif (
                    name.startswith("xl/charts/chart")
                    and name.endswith(".xml")
                    and "chartEx" not in name
                ):
                    xml = data.decode("utf-8")
                    xml = _re.sub(
                        r'<c:strCache>.*?</c:strCache>',
                        '<c:strCache><c:ptCount val="0"/></c:strCache>',
                        xml, flags=_re.DOTALL)
                    xml = _re.sub(
                        r'<c:numCache>.*?</c:numCache>',
                        '<c:numCache><c:formatCode>General</c:formatCode>'
                        '<c:ptCount val="0"/></c:numCache>',
                        xml, flags=_re.DOTALL)
                    data = xml.encode("utf-8")

                zout.writestr(item, data)

            # ── Escreve arquivos novos (SOMENTE chartEx + deps) ───────────────
            for zip_path, file_data in inject_map.items():
                if zip_path not in existing_names:
                    zout.writestr(zip_path, file_data)

        os.replace(tmp_path, output_path)
        if has_any_chartex:
            print("      [OK] chartEx injetado (minimal) + caches de séries limpos.")
        else:
            print("      [OK] Caches de séries limpos.")

    except Exception as exc:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        print(f"      [aviso] pós-processamento zip falhou: {exc}")


# ═══════════════════════════════════════════════════════════════════════════════
#                           MAIN ENTRY POINTS
# ═══════════════════════════════════════════════════════════════════════════════

def _fill_wip_sheet(wb, dates, bucket_names, bucket_matrix, x_total=None):
    sheet_name = "xWIP" if "xWIP" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    _w(ws, 1, 1, "Data")
    for j, bkt in enumerate(bucket_names):
        _w(ws, 1, j + 2, bkt)

    for i, dt in enumerate(dates):
        r = i + 2
        _w(ws, r, 1, dt)
        for j, bkt in enumerate(bucket_names):
            _w(ws, r, j + 2, bucket_matrix[bkt][i])


def _fill_cts_sheet(wb, dates, hu_labels, cts_matrix, fora_hu_daily):
    sheet_name = "xCTS" if "xCTS" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    def _short_hu_label(lbl):
        s = str(lbl)
        s = s.replace("Construção UI/UX Desktop", "Desktop")
        s = s.replace("Construção UI/UX Mobile", "Mobile")
        s = s.replace("Construção DEV", "Construção")
        s = s.replace("Testes DEV", "Testes")
        s = s.replace("Registro e Publicação", "Registro")
        return s

    _w(ws, 1, 1, "Data")
    all_labels = list(hu_labels) + ["Fora de HU"]
    for j, lbl in enumerate(all_labels):
        _w(ws, 1, j * 2 + 2, _short_hu_label(lbl))
        _w(ws, 1, j * 2 + 3, "qtd")

    for i, dt in enumerate(dates):
        r = i + 2
        _w(ws, r, 1, dt)
        for j, vals in enumerate(cts_matrix):
            _w(ws, r, j * 2 + 2, j + 1)
            _w(ws, r, j * 2 + 3, vals[i])
        j = len(hu_labels)
        _w(ws, r, j * 2 + 2, j + 1)
        _w(ws, r, j * 2 + 3, fora_hu_daily[i])


def _fill_dispersao_sheet(wb, hu_list, days, hu_matrix, nao_hu_daily, bd_start, hu_full_names):
    sheet_name = "xDispersaoTarefas" if "xDispersaoTarefas" in wb.sheetnames else "Dispersao"
    ws = wb[sheet_name]

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    def _short_disp_name(name):
        s = str(name)
        s = s.replace("UI/UX Desktop", "Desktop")
        s = s.replace("UI/UX Mobile", "Mobile")
        return s

    category_names = [_short_disp_name(hu_full_names.get(h, h)) for h in hu_list] + ["Não Presente em HU"]
    category_points = [dict() for _ in category_names]
    slots = [(-0.2, -0.2), (-0.2, 0.2), (0.2, -0.2), (0.2, 0.2)]

    for j in range(days):
        day_num = (bd_start + timedelta(days=j)).day
        hits = []
        for idx in range(len(hu_list)):
            cnt = hu_matrix[idx][j]
            if cnt > 0:
                hits.append((idx, float(cnt)))
        if nao_hu_daily[j] > 0:
            hits.append((len(hu_list), float(nao_hu_daily[j])))

        by_count = defaultdict(list)
        for idx, cnt in hits:
            by_count[cnt].append(idx)

        for cnt, idxs in by_count.items():
            idxs = sorted(idxs)
            if len(idxs) == 1:
                idx = idxs[0]
                category_points[idx][float(day_num)] = float(cnt)
            else:
                for k, idx in enumerate(idxs):
                    ox, oy = slots[k % len(slots)]
                    x = round(float(day_num) + ox, 1)
                    y = round(float(cnt) + oy, 1)
                    category_points[idx][x] = y

    x_values = sorted({x for pts in category_points for x in pts.keys()})

    _w(ws, 1, 1, "Nome da HU / Categoria")
    for c, x in enumerate(x_values, start=2):
        if abs(x - int(x)) < 1e-9:
            _w(ws, 1, c, int(x))
        else:
            _w(ws, 1, c, x)

    for i, name in enumerate(category_names):
        r = i + 2
        _w(ws, r, 1, name)
        pts = category_points[i]
        for c, x in enumerate(x_values, start=2):
            _w(ws, r, c, pts.get(x))


def _fill_burndown_sheet(wb, df, sprint_start, sprint_end, export_date):
    sheet_name = "xBurndownTarefas" if "xBurndownTarefas" in wb.sheetnames else "BurndownTarefas"
    ws = wb[sheet_name]
    ROWS = 31
    dfw = df[~df["is_backlog"]].copy()
    total = len(dfw)

    _w(ws, 1, 1, "Data")
    _w(ws, 1, 2, "Meta (Linear)")
    _w(ws, 1, 3, "Planejado")
    _w(ws, 1, 4, "A Realizar")

    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1
    plan_by_day = _build_step_plan(total, days, blocks=5)

    done_dates = sorted(
        d for d in dfw.loc[dfw["done_kpi"] & dfw["date_done"].notna(), "date_done"].tolist()
    )

    for i in range(ROWS):
        r = i + 2
        if i < days:
            d = bd_start + timedelta(days=i)
            dt = datetime(d.year, d.month, d.day)
            meta = round(total * (1 - i / max(days - 1, 1)))
            plan = plan_by_day[i]
            rlz = total - bisect_right(done_dates, d)
            _w(ws, r, 1, dt)
            _w(ws, r, 2, meta)
            _w(ws, r, 3, plan)
            _w(ws, r, 4, rlz)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)


def _fill_burndown_hu_sheet(wb, df, sprint_start, sprint_end, export_date):
    sheet_name = "xBurndownHU" if "xBurndownHU" in wb.sheetnames else "BurndownHU"
    ws = wb[sheet_name]
    ROWS = 31
    dfw = df[df["hu"] != ""].copy()

    _w(ws, 1, 1, "Data")
    _w(ws, 1, 2, "Meta (Linear)")
    _w(ws, 1, 3, "Planejado")
    _w(ws, 1, 4, "A Realizar")

    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1

    hu_info = []
    for _, grp in dfw.groupby("hu"):
        hu_total = len(grp)
        if bool(grp["done_kpi"].all()) and grp["date_done"].notna().all():
            completion_date = max(grp["date_done"].tolist())
        else:
            completion_date = None
        hu_info.append((hu_total, completion_date))

    total_hu_tasks = sum(t for t, _ in hu_info)
    plan_by_day = _build_step_plan(total_hu_tasks, days, blocks=5)

    for i in range(ROWS):
        r = i + 2
        if i < days:
            d = bd_start + timedelta(days=i)
            dt = datetime(d.year, d.month, d.day)
            meta = round(total_hu_tasks * (1 - i / max(days - 1, 1)))
            plan = plan_by_day[i]
            a_realiz = 0
            for hu_total, comp_date in hu_info:
                if comp_date is None or comp_date > d:
                    a_realiz += hu_total
            if i == 0:
                a_realiz = total_hu_tasks

            _w(ws, r, 1, dt)
            _w(ws, r, 2, meta)
            _w(ws, r, 3, plan)
            _w(ws, r, 4, a_realiz)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)


def _fill_burndown_storypoints_sheet(wb, df, sprint_end, hu_storypoints, storypoints_total=None):
    sheet_name = "xBurndownStorypoints" if "xBurndownStorypoints" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]
    ROWS = 31

    dfw_hu = df[(~df["is_backlog"]) & (df["hu"] != "")].copy()
    dfw_all = df[(~df["is_backlog"])].copy()
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1

    total_sp = float(sum(v for v in hu_storypoints.values() if v))
    if total_sp <= 0 and storypoints_total is not None:
        try:
            total_sp = float(storypoints_total)
        except Exception:
            total_sp = 0.0

    if total_sp <= 0:
        _w(ws, 1, 1, "Data")
        _w(ws, 1, 2, "Meta (Linear)")
        _w(ws, 1, 3, "Planejado")
        _w(ws, 1, 4, "A Realizar")
        for i in range(ROWS):
            r = i + 2
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)
        return

    plan_by_day = _build_step_plan(total_sp, days, blocks=5)
    task_weight = {}
    use_hu_weights = bool(hu_storypoints) and sum(v for v in hu_storypoints.values() if v) > 0
    if use_hu_weights:
        task_counts_by_hu = dfw_hu["hu"].value_counts().to_dict()
        for hu, count in task_counts_by_hu.items():
            sp = float(hu_storypoints.get(str(hu).upper(), 0) or 0)
            if count > 0:
                task_weight[hu] = sp / count
    else:
        total_tasks = len(dfw_all)
        default_weight = (total_sp / total_tasks) if total_tasks > 0 else 0.0
        done_dates_all = sorted(
            d for d in dfw_all.loc[dfw_all["done_kpi"] & dfw_all["date_done"].notna(), "date_done"].tolist()
        )

    _w(ws, 1, 1, "Data")
    _w(ws, 1, 2, "Meta (Linear)")
    _w(ws, 1, 3, "Planejado")
    _w(ws, 1, 4, "A Realizar")

    for i in range(ROWS):
        r = i + 2
        if i < days:
            d = bd_start + timedelta(days=i)
            dt = datetime(d.year, d.month, d.day)
            meta = round(total_sp * (1 - i / max(days - 1, 1)))
            plan = plan_by_day[i]

            if i == 0:
                a_realizar = total_sp
            else:
                if use_hu_weights:
                    a_realizar = 0.0
                    for _, row in dfw_hu.iterrows():
                        hu = row.get("hu", "")
                        w = task_weight.get(hu, 0.0)
                        d_done = row.get("date_done")
                        if not _is_valid_date(d_done) or d_done > d:
                            a_realizar += w
                else:
                    done_count = bisect_right(done_dates_all, d)
                    a_realizar = total_sp - (done_count * default_weight)
                    if a_realizar < 0:
                        a_realizar = 0.0

            _w(ws, r, 1, dt)
            _w(ws, r, 2, meta)
            _w(ws, r, 3, plan)
            _w(ws, r, 4, a_realizar)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)


def _precompute_burndown(df, sprint_start, sprint_end, export_date):
    ROWS = 31
    dfw = df[~df["is_backlog"]].copy()
    total = len(dfw)
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1
    plan_by_day = _build_step_plan(total, days, blocks=5)
    done_dates = sorted(
        d for d in dfw.loc[dfw["done_kpi"] & dfw["date_done"].notna(), "date_done"].tolist()
    )

    result = []
    for i in range(ROWS):
        if i < days:
            d = bd_start + timedelta(days=i)
            meta = round(total * (1 - i / max(days - 1, 1)))
            plan = plan_by_day[i]
            rlz = total - bisect_right(done_dates, d)
            result.append([d.strftime("%Y-%m-%d"), meta, plan, rlz])
        else:
            result.append([None, None, None, None])
    return result


def _precompute_burndown_hu(df, sprint_start, sprint_end, export_date):
    ROWS = 31
    dfw = df[df["hu"] != ""].copy()
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1

    hu_info = []
    for _, grp in dfw.groupby("hu"):
        hu_total = len(grp)
        if bool(grp["done_kpi"].all()) and grp["date_done"].notna().all():
            completion_date = max(grp["date_done"].tolist())
        else:
            completion_date = None
        hu_info.append((hu_total, completion_date))

    total_hu_tasks = sum(t for t, _ in hu_info)
    plan_by_day = _build_step_plan(total_hu_tasks, days, blocks=5)

    result = []
    for i in range(ROWS):
        if i < days:
            d = bd_start + timedelta(days=i)
            meta = round(total_hu_tasks * (1 - i / max(days - 1, 1)))
            plan = plan_by_day[i]
            a_realiz = 0
            for hu_total, comp_date in hu_info:
                if comp_date is None or comp_date > d:
                    a_realiz += hu_total
            if i == 0:
                a_realiz = total_hu_tasks
            result.append([d.strftime("%Y-%m-%d"), meta, plan, a_realiz])
        else:
            result.append([None, None, None, None])
    return result


def _fill_hu_sheet(wb, hu_list, hu_full_names, extras_count=0):
    sheet_name = "xHUs" if "xHUs" in wb.sheetnames else "HUs"
    ws = wb[sheet_name]

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    _w(ws, 1, 1, "Rótulo (História de Usuário)")
    _w(ws, 1, 2, "Quantidade de Registros")

    rows = sorted(hu_list, key=lambda x: x[0])
    r = 2
    for hu_id, total, _done in rows:
        _w(ws, r, 1, hu_full_names.get(hu_id, hu_id))
        _w(ws, r, 2, total)
        r += 1

    _w(ws, r, 1, "Ausente em HU")
    _w(ws, r, 2, extras_count)


def _fill_areas_sheet(wb, area_rows):
    sheet_name = "xAreas" if "xAreas" in wb.sheetnames else "Areas"
    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    _w(ws, 1, 1, "“Área”")
    _w(ws, 1, 2, "“Qtd”")
    for i, (area, done, pend) in enumerate(area_rows, start=2):
        _w(ws, i, 1, area)
        _w(ws, i, 2, done + pend)


def _fill_hu_in_out_sheet(wb, in_out):
    sheet_name = "xHU_inOut" if "xHU_inOut" in wb.sheetnames else "HUInOut"
    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    _w(ws, 1, 1, "Descrição")
    _w(ws, 1, 2, "Quantidade")
    for i, (label, cnt) in enumerate(in_out, start=2):
        if label.lower().startswith("em hu"):
            label_txt = '"Em HU"'
        elif label.lower().startswith("fora"):
            label_txt = '"Fora de HU"'
        else:
            label_txt = label
        _w(ws, i, 1, label_txt)
        _w(ws, i, 2, cnt)


def _extract_gp_effort_by_hu(wb):
    sheet_name = None
    for cand in ("gp_Plann_Sprint", "Gp_Plann_Sprint"):
        if cand in wb.sheetnames:
            sheet_name = cand
            break
    if sheet_name is None:
        return {}

    ws = wb[sheet_name]
    hu_cols = []
    for c in range(2, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is None or str(v).strip() == "":
            continue
        m = re.search(r"\b(HU\d+)\b", str(v), flags=re.IGNORECASE)
        if m:
            hu_cols.append((c, m.group(1).upper()))

    out = {hu: {"max": 0.0, "total": 0.0} for _, hu in hu_cols}

    max_row_idx = None
    total_row_idx = None
    for r in range(1, ws.max_row + 1):
        txt = _norm_label_key(ws.cell(r, 1).value)
        if "mensuracao (maior esforco individual)" in txt:
            max_row_idx = r
        if "mensuracao soma total esforco por hu" in txt:
            total_row_idx = r

    for c, hu in hu_cols:
        if max_row_idx and total_row_idx:
            v_max = ws.cell(max_row_idx, c).value
            v_tot = ws.cell(total_row_idx, c).value
            try:
                out[hu]["max"] = float(str(v_max).replace(",", ".")) if v_max not in (None, "") else 0.0
            except Exception:
                out[hu]["max"] = 0.0
            try:
                out[hu]["total"] = float(str(v_tot).replace(",", ".")) if v_tot not in (None, "") else 0.0
            except Exception:
                out[hu]["total"] = 0.0
            continue

        # Fallback: deriva da tabela de perfis (linhas numéricas da planilha)
        vals = []
        for r in range(2, ws.max_row + 1):
            first = ws.cell(r, 1).value
            if first is None and r > 20:
                break
            raw = ws.cell(r, c).value
            try:
                if isinstance(raw, str) and raw.strip().upper() == "N/A":
                    continue
                if raw in (None, ""):
                    continue
                vals.append(float(str(raw).replace(",", ".")))
            except Exception:
                continue
        if vals:
            out[hu]["max"] = max(vals)
            out[hu]["total"] = sum(vals)
    return out


def _fill_esforco_hu_sheet(wb, df_all, hu_full_names):
    sheet_name = "xEsforcoHU" if "xEsforcoHU" in wb.sheetnames else None
    if not sheet_name:
        return
    ws = wb[sheet_name]

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    effort = _extract_gp_effort_by_hu(wb)
    hu_ids = sorted(set(list(effort.keys()) + df_all[df_all["hu"] != ""]["hu"].unique().tolist()))

    _w(ws, 1, 1, "História de Usuário (HU)")
    _w(ws, 1, 2, "Esforço Max. por Perfil")
    _w(ws, 1, 3, "Esforço Total HU")
    _w(ws, 1, 4, "Total Tarefas")
    _w(ws, 1, 5, "Concluídas")

    r = 2
    for hu in hu_ids:
        grp = df_all[df_all["hu"] == hu]
        _w(ws, r, 1, hu_full_names.get(hu, hu))
        _w(ws, r, 2, effort.get(hu, {}).get("max", 0))
        _w(ws, r, 3, effort.get(hu, {}).get("total", 0))
        _w(ws, r, 4, int(len(grp)))
        _w(ws, r, 5, int(grp["done_kpi"].sum()))
        r += 1

    sem_hu_grp = df_all[df_all["hu"] == ""]
    _w(ws, r, 1, "Geral / Não Previstos / Bugs*")
    _w(ws, r, 2, 0)
    _w(ws, r, 3, 0)
    _w(ws, r, 4, int(len(sem_hu_grp)))
    _w(ws, r, 5, int(sem_hu_grp["done_kpi"].sum()))


def fill_template(template_path, input_path, output_path,
                  author="Gerado automaticamente",
                  dados_manuais_path=None,
                  report_type="equipe",
                  ignore_labels=None,
                  form_data=None,
                  storypoints_total=None):
    """
    Main pipeline:
    1. Load + compute from Planner export
    2. Copy template to output path
    3. Fill all data sheets in the template copy
    4. Save

    report_type:
      "equipe" (padrão) → One Page Report completo com KPIs, burndown, dispersão, etc.
      "gp"              → [Em desenvolvimento] Visão executiva para gerente de projeto.

    Dados automáticos: Roadmap, KPIs, HUs, Burndown, Dispersão, Áreas,
    Colaboradores, Categorias — extraídos do export do Planner.
    Dados manuais preservados no template: Planejamento, Histórico, EsforcoHU.
    """
    if report_type == "gp":
        raise NotImplementedError(
            "O Relatório para GP ainda está em desenvolvimento. "
            "Por enquanto, use o Relatório para Equipe."
        )
    print(f"[1/6] Lendo arquivo base: {input_path}")
    df_raw, plan_name, export_date_str = load_base(input_path)

    print(f"[2/6] Processando dados ({len(df_raw)} tarefas)...")
    df_all, sprint_name, sprint_start, sprint_end, export_date, sprint_goal = \
        compute_all(df_raw, plan_name, export_date_str)

    removed_ignored = 0
    df_scope = df_all
    if ignore_labels:
        df_scope, removed_ignored = _apply_ignore_labels(df_all, ignore_labels=ignore_labels)
        if removed_ignored:
            print(f"      [filtro] {removed_ignored} tarefa(s) removida(s) por rótulos ignorados: {ignore_labels}")

    print(f"[3/6] Calculando métricas...")
    # Abas gerais (sem filtro)
    kpis = build_kpis(df_all)
    hu_list = build_hu_list(df_all)
    hu_full_names = build_hu_full_names(df_all)
    collab_rows = build_por_colaborador(df_all)
    area_rows = build_areas(df_all)
    cat_rows, bub_rows = build_por_categoria(df_all)
    in_out = build_hu_in_out(df_all)

    # Abas de escopo (com filtro de rótulos ignorados)
    hu_disp, days, hu_matrix, nao_hu_daily = \
        build_dispersao_daily(df_scope, sprint_start, sprint_end)
    rotulos_rows = build_rotulos(df_all)
    resp_rows = build_responsaveis(df_scope)
    hist_31, indicativos = build_histograma(df_scope)
    stats_rows = build_histograma2(df_scope)

    ref_month = sprint_end if sprint_end is not None else sprint_start
    cfd_dates, cfd_todo, cfd_doing, cfd_done = build_cfd(df_scope, ref_month, export_date)
    wip_dates, wip_buckets, wip_matrix, wip_total = build_wip(df_scope, ref_month, export_date)
    cts_dates, cts_hu_labels, cts_matrix, fora_hu_daily_cts = build_cts(df_scope, ref_month)

    # Extrai metadados da tarefa Gestão e Roadmap diretamente do Planner
    gestao_meta = build_gestao_meta(df_all)
    roadmap_items = build_roadmap_from_df(df_all)

    print(f"[4/6] Carregando template: {template_path}")
    wb = load_workbook(template_path)

    # Lê esforço planejado do template (Report + EsforcoHU – mantidos manualmente)
    template_effort = read_template_effort(wb)

    manual_injection = {"roadmap": False, "report_plan_hist": False, "gp_tabs": []}

    # Injeta dados manuais do arquivo adicional (legado + gp_*)
    if dados_manuais_path:
        print(f"      [manual] Injetando dados manuais: {os.path.basename(dados_manuais_path)}")
        manual_injection = inject_dados_manuais_ext(wb, dados_manuais_path)
        if manual_injection.get("gp_tabs"):
            print(f"      [manual] Abas gp_* aplicadas: {manual_injection.get('gp_tabs')}")

    # Injeta dados do formulário web (app.py) nas abas gp_* do template
    form_filled = []
    if form_data:
        form_filled = inject_dados_formulario(wb, form_data)
        if form_filled:
            print(f"      [form] Dados do formulário injetados: {form_filled}")

    # Normaliza percentuais da disponibilidade (form/manual) para fração 0..1.
    if "gp_Disponibilidade" in wb.sheetnames:
        _normalize_gp_disponibilidade_sheet(wb["gp_Disponibilidade"])

    # Storypoints e métricas ponderadas por HU devem refletir a aba gp_Plann_Sprint
    hu_storypoints = _extract_gp_storypoints_by_hu(wb)
    if hu_storypoints:
        _enrich_kpis_with_hu_storypoints(kpis, df_scope, hu_storypoints)
    elif storypoints_total is not None and float(storypoints_total) > 0:
        kpis["storypoints"] = round(float(storypoints_total), 2)

    print(f"[5/6] Preenchendo abas de dados...")
    _form_cab = form_data.get("cabecalhos", {}) if form_data else {}
    _has_form = bool(form_data)
    _fill_kpi_sheet(
        wb, kpis, sprint_name,
        _form_cab.get("sprint_goal") or sprint_goal,
        projeto=_form_cab.get("projeto") or gestao_meta.get("projeto", sprint_name),
        gerente=_form_cab.get("gerente") or gestao_meta.get("gerente", ""),
        linkedin=_form_cab.get("linkedin") or gestao_meta.get("linkedin", ""),
        product_goal=_form_cab.get("product_goal", ""),
        export_date=export_date,
        nome_arquivo=os.path.basename(input_path),
        write_cabecalhos=_has_form,
    )

    # Roadmap: preenchido automaticamente a partir da tarefa 'Roadmap' no Planner
    roadmap_already_manual = bool(
        manual_injection.get("roadmap")
        or ("gp_Roadmap" in manual_injection.get("gp_tabs", []))
        or ("gp_Roadmap" in form_filled)
        or ("roadmap" in (form_data or {}))
    )
    if roadmap_items and not roadmap_already_manual:
        _fill_roadmap_sheet(wb, roadmap_items)
    _fill_hu_sheet(wb, hu_list, hu_full_names, extras_count=kpis["sem_hu"])
    _fill_esforco_hu_sheet(wb, df_all, hu_full_names)
    _fill_burndown_sheet(wb, df_scope, sprint_start, sprint_end, export_date)
    _fill_burndown_hu_sheet(wb, df_scope, sprint_start, sprint_end, export_date)
    _fill_burndown_storypoints_sheet(
        wb, df_scope, sprint_end, hu_storypoints, storypoints_total=storypoints_total
    )
    bd_start, _ = _month_bounds(sprint_end)
    _fill_dispersao_sheet(wb, hu_disp, days, hu_matrix, nao_hu_daily,
                          bd_start, hu_full_names)
    _fill_colaborador_sheet(wb, collab_rows)
    _fill_areas_sheet(wb, area_rows)
    _fill_hu_in_out_sheet(wb, in_out)
    _fill_categoria_sheet(wb, cat_rows)
    # Bubbles: usa esforço do template (planejamento) em vez de esforço de checklist
    _fill_categoria_bubbles_sheet(wb, bub_rows, template_effort=template_effort)
    # Fill new sheets
    _fill_rotulos_sheet(wb, rotulos_rows)
    _fill_responsaveis_sheet(wb, resp_rows)
    _fill_histograma_sheet(wb, hist_31, indicativos)
    _fill_histograma2_sheet(wb, stats_rows)
    _fill_cfd_sheet(wb, cfd_dates, cfd_todo, cfd_doing, cfd_done)
    _fill_wip_sheet(wb, wip_dates, wip_buckets, wip_matrix, wip_total)
    _fill_cts_sheet(wb, cts_dates, cts_hu_labels, cts_matrix, fora_hu_daily_cts)
    # Mantém tabelas manuais quando vierem de planilha/formulário.
    if not dados_manuais_path and not form_data:
        _clear_report_planning_tables(wb)
    _clear_report_roadmap_overlay_cells(wb)
    if not dados_manuais_path:
        _update_report_metadata(wb, export_date, author)

    print(f"[6/6] Salvando: {output_path}")
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    wb.save(output_path)

    # Pós-processamento ZIP: injeta chartEx e limpa caches de séries dos charts
    print("      Pós-processamento...")
    _postprocess_xlsx(output_path, template_path=template_path)

    print(f"[OK] Dashboard gerado com sucesso!")
    print(f"     Sprint : {sprint_name}")
    print(f"     Total  : {kpis['total']}  |  Concluídas: {kpis['done']}  "
          f"|  Pendentes: {kpis['pending']}")
    print(f"     HUs    : {len(hu_list)}")


def main():
    parser = argparse.ArgumentParser(description="OnePageReport Dashboard Generator v3")
    parser.add_argument("input", nargs="?", help="Arquivo base exportado do Planner (.xlsx)")
    parser.add_argument("output", nargs="?", help="Arquivo de saída (.xlsx)")
    parser.add_argument("--template", default=None,
                        help="Arquivo espelho/template (.xlsx). "
                             "Se não informado, usa template.xlsx na mesma pasta do script.")
    parser.add_argument("--author", default="Gerado automaticamente",
                        help="Nome do autor para o relatório")
    parser.add_argument(
        "--make-clean-template",
        nargs=2,
        metavar=("SOURCE_XLSX", "OUTPUT_XLSX"),
        help="Gera um template limpo a partir de um workbook de referência.",
    )
    args = parser.parse_args()

    if args.make_clean_template:
        src, dst = args.make_clean_template
        if not os.path.exists(src):
            print(f"[ERRO] Arquivo de origem não encontrado: {src}", file=sys.stderr)
            sys.exit(1)
        print(f"[clean-template] origem: {src}")
        print(f"[clean-template] saída : {dst}")
        create_clean_template(src, dst)
        print("[OK] Template limpo gerado com sucesso.")
        return

    if not args.input:
        parser.error("Informe <input.xlsx> ou use --make-clean-template SOURCE OUTPUT")

    input_path  = args.input
    output_path = args.output or os.path.splitext(input_path)[0] + "_Dashboard.xlsx"

    if not os.path.exists(input_path):
        print(f"[ERRO] Arquivo não encontrado: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Resolve template path
    if args.template:
        template_path = args.template
    else:
        # Default: prefer template.next.xlsx (quando existir), fallback template.xlsx
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidate_next = os.path.join(script_dir, "template.next.xlsx")
        candidate_default = os.path.join(script_dir, "template.xlsx")
        template_path = candidate_next if os.path.exists(candidate_next) else candidate_default

    if not os.path.exists(template_path):
        print(f"[ERRO] Template não encontrado: {template_path}", file=sys.stderr)
        print("       Informe o caminho com --template ou coloque 'template.xlsx' "
              "na mesma pasta do script.", file=sys.stderr)
        sys.exit(1)

    fill_template(template_path, input_path, output_path, author=args.author)



# ═══════════════════════════════════════════════════════════════════════════════
#                    JSON INTERMEDIATE LAYER  (xlsx → JSON → Excel)
# ═══════════════════════════════════════════════════════════════════════════════

def _precompute_burndown(df, sprint_start, sprint_end, export_date):
    ROWS = 31
    dfw = df[~df["is_backlog"]].copy()
    total = len(dfw)
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1
    plan_by_day = _build_step_plan(total, days, blocks=5)

    done_dates = sorted(
        d for d in dfw.loc[dfw["done_kpi"] & dfw["date_done"].notna(), "date_done"].tolist()
    )

    result = []
    for i in range(ROWS):
        if i < days:
            d = bd_start + timedelta(days=i)
            dt_str = d.strftime("%Y-%m-%d")
            meta = round(total * (1 - i / max(days - 1, 1)))
            plan = plan_by_day[i]
            concluded = bisect_right(done_dates, d)
            rlz = total - concluded
            result.append([dt_str, meta, plan, rlz])
        else:
            result.append([None, None, None, None])
    return result


def _precompute_burndown_hu(df, sprint_start, sprint_end, export_date):
    ROWS = 31
    dfw = df[df["hu"] != ""].copy()
    bd_start, bd_end = _month_bounds(sprint_end)
    days = (bd_end - bd_start).days + 1

    hu_info = []
    for _, grp in dfw.groupby("hu"):
        hu_total = len(grp)
        if bool(grp["done_kpi"].all()) and grp["date_done"].notna().all():
            completion_date = max(grp["date_done"].tolist())
        else:
            completion_date = None
        hu_info.append((hu_total, completion_date))

    total_hu_tasks = sum(t for t, _ in hu_info)
    plan_by_day = _build_step_plan(total_hu_tasks, days, blocks=5)

    result = []
    for i in range(ROWS):
        if i < days:
            d = bd_start + timedelta(days=i)
            dt_str = d.strftime("%Y-%m-%d")
            meta = round(total_hu_tasks * (1 - i / max(days - 1, 1)))
            plan = plan_by_day[i]
            a_realiz = 0
            for hu_total, comp_date in hu_info:
                if comp_date is None or comp_date > d:
                    a_realiz += hu_total
            if i == 0:
                a_realiz = total_hu_tasks
            result.append([dt_str, meta, plan, a_realiz])
        else:
            result.append([None, None, None, None])
    return result


def _fill_burndown_from_rows(wb, rows, sheet_key="tarefas"):
    """
    Preenche uma aba de burndown a partir de linhas pré-computadas.
    sheet_key: "tarefas" ou "hu"
    rows: lista de [date_str, meta, plan, a_realizar] com 31 itens.
    """
    if sheet_key == "tarefas":
        sheet_name = "xBurndownTarefas" if "xBurndownTarefas" in wb.sheetnames else "BurndownTarefas"
    else:
        sheet_name = "xBurndownHU" if "xBurndownHU" in wb.sheetnames else "BurndownHU"

    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]

    _w(ws, 1, 1, "Data")
    _w(ws, 1, 2, "Meta (Linear)")
    _w(ws, 1, 3, "Planejado")
    _w(ws, 1, 4, "A Realizar")

    for i, row in enumerate(rows):
        r = i + 2
        dt_str, meta, plan, rlz = row
        if dt_str:
            dt = datetime.strptime(dt_str, "%Y-%m-%d")
            _w(ws, r, 1, dt)
            _w(ws, r, 2, meta)
            _w(ws, r, 3, plan)
            _w(ws, r, 4, rlz)
        else:
            _w(ws, r, 1, None)
            _w(ws, r, 2, None)
            _w(ws, r, 3, None)
            _w(ws, r, 4, None)


def _collect_warnings(df, kpis, hu_list, sprint_start, sprint_end, export_date,
                       roadmap_items):
    """
    Analisa a qualidade dos dados e retorna lista de avisos para o usuario.
    Cada item: {"nivel": "erro"|"aviso"|"info", "msg": "..."}.
    Niveis:
      "erro"  - dados criticos ausentes, o relatorio pode estar incompleto.
      "aviso" - situacao preocupante, verificar antes de usar o relatorio.
      "info"  - observacao informativa, nao critica.
    """
    from datetime import date as _date_cls
    warns = []

    total = kpis.get("total", 0)

    # --- Sem tarefas ---
    if total == 0:
        warns.append({"nivel": "erro",
                      "msg": "Nenhuma tarefa encontrada no arquivo. "
                             "Verifique se e um export valido do Planner."})
        return warns  # sem dados uteis para analise adicional

    # --- Sprint sem datas ---
    if sprint_start is None:
        warns.append({"nivel": "aviso",
                      "msg": "Data de inicio do sprint nao encontrada. "
                             "O burndown de tarefas pode ficar vazio."})
    if sprint_end is None:
        warns.append({"nivel": "aviso",
                      "msg": "Data de fim do sprint nao encontrada. "
                             "Burndown e KPIs de prazo podem ficar vazios."})

    # --- Export muito antigo ---
    if export_date:
        today = _date_cls.today()
        days_old = (today - export_date).days
        if days_old > 14:
            warns.append({"nivel": "info",
                          "msg": f"Export com {days_old} dias de antecedencia. "
                                  "Considere usar um export mais recente para "
                                  "resultados precisos."})

    # --- Sprint encerrada ha muito tempo ---
    if sprint_end and export_date:
        delta = (export_date - sprint_end).days
        if delta > 30:
            warns.append({"nivel": "info",
                          "msg": f"Sprint encerrada ha {delta} dias. "
                                  "Os dados podem estar desatualizados."})

    # --- Poucas tarefas ---
    if 0 < total < 5:
        warns.append({"nivel": "aviso",
                      "msg": f"Apenas {total} tarefa(s) encontrada(s). "
                             "O relatorio pode estar incompleto."})

    # --- Nenhuma HU cadastrada ---
    if not hu_list:
        warns.append({"nivel": "aviso",
                      "msg": "Nenhuma Historia de Usuario (HU) identificada. "
                             "Verifique se as tarefas estao vinculadas a HUs."})
    else:
        # --- Alta proporcao de tarefas sem HU ---
        sem_hu = kpis.get("sem_hu", 0)
        if total > 0 and sem_hu / total > 0.5:
            pct = int(sem_hu / total * 100)
            warns.append({"nivel": "aviso",
                          "msg": f"{pct}% das tarefas ({sem_hu}/{total}) "
                                  "nao estao vinculadas a nenhuma HU."})

    # --- Muitas tarefas sem responsavel ---
    try:
        import pandas as _pd
        sem_resp = int((df["assignee"].fillna("").astype(str).str.strip() == "").sum())
        if total > 0 and sem_resp / total > 0.3:
            pct = int(sem_resp / total * 100)
            warns.append({"nivel": "info",
                          "msg": f"{pct}% das tarefas ({sem_resp}/{total}) "
                                  "nao possuem responsavel atribuido."})
    except Exception:
        pass

    # --- Roadmap vazio ---
    if not roadmap_items:
        warns.append({"nivel": "info",
                      "msg": "Nenhum item de Roadmap encontrado no arquivo. "
                             "O grafico de Roadmap ficara vazio."})

    # --- Sprint encerrada com baixa conclusao ---
    done = kpis.get("done", 0)
    if sprint_end and export_date and sprint_end <= export_date:
        if total > 0 and done / total < 0.5:
            pct = int(done / total * 100)
            warns.append({"nivel": "aviso",
                          "msg": f"Sprint encerrada com apenas {pct}% de conclusao "
                                  f"({done}/{total} tarefas)."})

    return warns


def compute_json(input_path, ignore_labels=None):
    """
    Pipeline parcial: le o export do Planner e retorna um dict JSON-serializavel
    com todos os dados computados (KPIs, burndown, CFD, WIP, CTS, etc.).

    Uso tipico:
        data = compute_json("planner_export.xlsx")
        import json
        with open("dados.json", "w") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    """
    def _d(v):
        if v is None:
            return None
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        return str(v)

    df_raw, plan_name, export_date_str = load_base(input_path)
    df_all, sprint_name, sprint_start, sprint_end, export_date, sprint_goal = \
        compute_all(df_raw, plan_name, export_date_str)

    df_scope = df_all
    if ignore_labels:
        df_scope, _ = _apply_ignore_labels(df_all, ignore_labels=ignore_labels)

    kpis = build_kpis(df_all)
    hu_list = build_hu_list(df_all)
    hu_full_names = build_hu_full_names(df_all)
    collab_rows = build_por_colaborador(df_all)
    area_rows = build_areas(df_all)
    cat_rows, bub_rows = build_por_categoria(df_all)
    in_out = build_hu_in_out(df_all)
    hu_disp, days, hu_matrix, nao_hu_daily = build_dispersao_daily(df_scope, sprint_start, sprint_end)
    rotulos_rows = build_rotulos(df_all)
    resp_rows = build_responsaveis(df_scope)
    hist_31, indicativos = build_histograma(df_scope)
    stats_rows = build_histograma2(df_scope)
    gestao_meta = build_gestao_meta(df_all)
    roadmap_items = build_roadmap_from_df(df_all)

    # Coleta warnings de qualidade dos dados
    warnings = _collect_warnings(
        df_all, kpis, hu_list, sprint_start, sprint_end, export_date, roadmap_items
    )

    ref_month = sprint_end if sprint_end is not None else sprint_start
    cfd_dates, cfd_todo, cfd_doing, cfd_done = build_cfd(df_scope, ref_month, export_date)
    wip_dates, wip_buckets, wip_matrix, _wip_total = build_wip(df_scope, ref_month, export_date)
    cts_dates, cts_hu_labels, cts_matrix, fora_hu_daily_cts = build_cts(df_scope, ref_month)

    burndown_rows = _precompute_burndown(df_scope, sprint_start, sprint_end, export_date)
    burndown_hu_rows = _precompute_burndown_hu(df_scope, sprint_start, sprint_end, export_date)
    bd_start, _      = _month_bounds(sprint_end)

    # Garante tipos JSON-serializaveis em kpis
    kpis_json = {}
    for k, v in kpis.items():
        if v is None:
            kpis_json[k] = None
        elif isinstance(v, float):
            kpis_json[k] = round(v, 6)
        elif hasattr(v, "__int__"):
            kpis_json[k] = int(v)
        else:
            kpis_json[k] = v

    return {
        "meta": {
            "sprint_name":  sprint_name,
            "sprint_goal":  sprint_goal,
            "sprint_start": _d(sprint_start),
            "sprint_end":   _d(sprint_end),
            "export_date":  _d(export_date),
            "nome_arquivo": os.path.basename(input_path),
            "projeto":      gestao_meta.get("projeto", sprint_name),
            "gerente":      gestao_meta.get("gerente", ""),
            "linkedin":     gestao_meta.get("linkedin", ""),
        },
        "kpis":        kpis_json,
        "hu_list":     [[h, t, d] for h, t, d in hu_list],
        "hu_full_names": hu_full_names,
        "collab_rows": [[n, d, p] for n, d, p in collab_rows],
        "area_rows":   [[a, d, p] for a, d, p in area_rows],
        "cat_rows":    [[c, d, p] for c, d, p in cat_rows],
        "bub_rows":    [[c, d, t, e] for c, d, t, e in bub_rows],
        "in_out":      [[label, cnt] for label, cnt in in_out],
        "rotulos_rows": [[disp, dn, pn, lt, ct] for disp, dn, pn, lt, ct in rotulos_rows],
        "resp_rows":   [[disp, dn, pn, lt, ct] for disp, dn, pn, lt, ct in resp_rows],
        "hist_31":     hist_31,
        "indicativos": indicativos,
        "stats_rows":  [[desc, val] for desc, val in stats_rows],
        "cfd": {
            "dates": [_d(dt) for dt in cfd_dates],
            "todo":  cfd_todo,
            "doing": cfd_doing,
            "done":  cfd_done,
        },
        "wip": {
            "dates":   [_d(dt) for dt in wip_dates],
            "buckets": wip_buckets,
            "matrix":  {bkt: wip_matrix[bkt] for bkt in wip_buckets},
            "x_total": _wip_total,
        },
        "cts": {
            "dates":     [_d(dt) for dt in cts_dates],
            "hu_labels": list(cts_hu_labels),
            "matrix":    cts_matrix,
            "fora_hu":   fora_hu_daily_cts,
        },
        "burndown": {
            "rows": burndown_rows,
        },
        "burndown_hu": {
            "rows": burndown_hu_rows,
        },
        "dispersao": {
            "hu_list":      hu_disp,
            "days":         days,
            "hu_matrix":    hu_matrix,
            "nao_hu_daily": nao_hu_daily,
            "bd_start":     _d(bd_start),
        },
        "roadmap_items": [[t, p] for t, p in roadmap_items],
        "warnings":      warnings,
    }


def fill_template_from_json(template_path, data, output_path,
                             author="Gerado automaticamente"):
    """
    Preenche o template Excel a partir de um dict JSON (produzido por compute_json).
    Esta e a etapa 2 do pipeline xlsx->JSON->Excel.

    Util quando o JSON ja foi pre-computado (ex: enviado via browser) e
    so e necessario gerar o arquivo final.
    """
    meta = data.get("meta", {})

    def _parse_date(s):
        if not s:
            return date.today()
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except Exception:
            return date.today()

    export_date = _parse_date(meta.get("export_date"))
    sprint_name = meta.get("sprint_name", "Sprint")
    sprint_goal = meta.get("sprint_goal", "")

    wb = load_workbook(template_path)
    template_effort = read_template_effort(wb)
    hu_full_names = data.get("hu_full_names", {})

    # KPIs
    _fill_kpi_sheet(
        wb, data["kpis"], sprint_name, sprint_goal,
        projeto=meta.get("projeto", sprint_name),
        gerente=meta.get("gerente", ""),
        linkedin=meta.get("linkedin", ""),
        export_date=export_date,
        nome_arquivo=meta.get("nome_arquivo", ""),
    )

    # Roadmap
    roadmap_items = [(t, p) for t, p in data.get("roadmap_items", [])]
    if roadmap_items:
        _fill_roadmap_sheet(wb, roadmap_items)

    # HUs
    hu_list = [(h, t, d) for h, t, d in data.get("hu_list", [])]
    _fill_hu_sheet(wb, hu_list, hu_full_names,
                   extras_count=data["kpis"].get("sem_hu", 0))

    # Burndown
    _fill_burndown_from_rows(wb, data["burndown"]["rows"], sheet_key="tarefas")
    _fill_burndown_from_rows(wb, data["burndown_hu"]["rows"], sheet_key="hu")

    # Dispersao
    disp = data.get("dispersao", {})
    bd_start_str = disp.get("bd_start")
    if bd_start_str:
        bd_start = _parse_date(bd_start_str)
        _fill_dispersao_sheet(
            wb,
            disp.get("hu_list", []),
            disp.get("days", 31),
            disp.get("hu_matrix", []),
            disp.get("nao_hu_daily", []),
            bd_start,
            hu_full_names,
        )

    # Colaboradores, Areas, Categorias
    _fill_colaborador_sheet(wb, [(n, d, p) for n, d, p in data.get("collab_rows", [])])
    _fill_areas_sheet(wb,   [(a, d, p) for a, d, p in data.get("area_rows", [])])
    _fill_hu_in_out_sheet(wb, [(lbl, cnt) for lbl, cnt in data.get("in_out", [])])
    _fill_categoria_sheet(wb, [(c, d, p) for c, d, p in data.get("cat_rows", [])])
    _fill_categoria_bubbles_sheet(
        wb,
        [(c, d, t, e) for c, d, t, e in data.get("bub_rows", [])],
        template_effort=template_effort,
    )

    # Rotulos e Responsaveis
    _fill_rotulos_sheet(
        wb, [(d, dn, pn, lt, ct) for d, dn, pn, lt, ct in data.get("rotulos_rows", [])]
    )
    _fill_responsaveis_sheet(
        wb, [(d, dn, pn, lt, ct) for d, dn, pn, lt, ct in data.get("resp_rows", [])]
    )

    # Histograma
    _fill_histograma_sheet(wb, data.get("hist_31", [0]*31), data.get("indicativos", [None]*31))
    _fill_histograma2_sheet(wb, [(d, v) for d, v in data.get("stats_rows", [])])

    # CFD / WIP / CTS
    cfd = data.get("cfd", {})
    if cfd.get("dates"):
        cfd_dates = [datetime.strptime(s, "%Y-%m-%d") for s in cfd["dates"]]
        _fill_cfd_sheet(wb, cfd_dates, cfd["todo"], cfd["doing"], cfd["done"])

    wip = data.get("wip", {})
    if wip.get("dates"):
        wip_dates = [datetime.strptime(s, "%Y-%m-%d") for s in wip["dates"]]
        _fill_wip_sheet(
            wb, wip_dates, wip["buckets"], wip["matrix"], wip.get("x_total")
        )

    cts = data.get("cts", {})
    if cts.get("dates"):
        cts_dates = [datetime.strptime(s, "%Y-%m-%d") for s in cts["dates"]]
        _fill_cts_sheet(wb, cts_dates, cts["hu_labels"], cts["matrix"], cts["fora_hu"])

    # Report manual structure
    _clear_report_planning_tables(wb)
    _clear_report_roadmap_overlay_cells(wb)
    _update_report_metadata(wb, export_date, author)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    wb.save(output_path)

    print("      Pos-processamento...")
    _postprocess_xlsx(output_path, template_path=template_path)
    print("[OK] Dashboard gerado a partir de JSON!")



if __name__ == "__main__":
    main()
