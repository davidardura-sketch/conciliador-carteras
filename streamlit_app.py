# app_conciliacion_carteras.py
# v2.4 - Parser Bloomberg flexible + diagnóstico de conciliación
# Autor: ChatGPT para David Ardura / Finaccess Value

from __future__ import annotations

import io
import re
import math
import tempfile
from copy import copy
from datetime import datetime, date
from typing import Dict, List, Tuple, Optional, Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import PieChart, Reference
from openpyxl.styles import Font, PatternFill, Border, Alignment, Protection
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

ISIN_RE = re.compile(r"^[A-Z]{2}[A-Z0-9]{9}[0-9]$")

SPANISH_CATEGORY = {
    "cash": "Efectivo",
    "efectivo": "Efectivo",
    "euro": "Efectivo",
    "etf": "ETF",
    "etcs": "ETF",
    "financial institutions": "Financieros",
    "instituciones financieras": "Financieros",
    "financial": "Financieros",
    "industrial": "Industrial",
    "industrials": "Industrial",
    "sovereign": "Soberanos",
    "treasury": "Soberanos",
    "tesorería": "Soberanos",
    "tesoreria": "Soberanos",
    "government": "Soberanos",
    "deuda publica": "Soberanos",
    "deuda pública": "Soberanos",
    "pagares": "Pagarés",
    "pagarés": "Pagarés",
    "pagare": "Pagarés",
    "pagaré": "Pagarés",
}

CATEGORY_ORDER = ["Efectivo", "ETF", "Financieros", "Industrial", "Soberanos", "Pagarés", "Solo depositario", "Solo Bloomberg", "Otros"]


def norm_text(x: Any) -> str:
    if x is None:
        return ""
    return str(x).replace("\xa0", " ").strip()


def clean_isin(x: Any) -> str:
    s = norm_text(x).upper().replace(" ", "")
    return s if ISIN_RE.match(s) else ""


def as_float(x: Any, default: Optional[float] = None) -> Optional[float]:
    if x is None or x == "":
        return default
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        if math.isnan(x):
            return default
        return float(x)
    s = str(x).strip().replace("%", "")
    # formatos europeos básicos
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return default


def spanish_category(raw: Any, fallback: str = "Otros") -> str:
    s = norm_text(raw).lower().replace("›", "").strip()
    for key, val in SPANISH_CATEGORY.items():
        if key in s:
            return val
    return fallback


def infer_depositary_category(product: str, market: str = "") -> str:
    p = product.lower()
    m = market.lower()
    if any(k in p for k in ["pagare", "pagaré", "pg ", "pagar "]):
        return "Pagarés"
    if any(k in p for k in ["etf", "etc", "crude", "brent"]) or "xtr" in m:
        return "ETF"
    if any(k in p for k in ["dbr", "ragb", " t ", "treasury", "bund", "bono", "obligacion"]):
        return "Soberanos"
    return "Solo depositario"


def find_header_row(ws, required_terms: List[str], max_scan: int = 25) -> Optional[int]:
    terms = [t.lower() for t in required_terms]
    for r in range(1, min(ws.max_row, max_scan) + 1):
        vals = [norm_text(ws.cell(r, c).value).lower() for c in range(1, ws.max_column + 1)]
        if all(any(t in v for v in vals) for t in terms):
            return r
    return None


def column_map(ws, header_row: int) -> Dict[str, int]:
    m: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        key = norm_text(ws.cell(header_row, c).value).lower()
        key = re.sub(r"\s+", " ", key)
        if key:
            m[key] = c
    return m


def find_col(cols: Dict[str, int], candidates: List[str]) -> Optional[int]:
    for cand in candidates:
        cand_l = cand.lower()
        for k, v in cols.items():
            if cand_l == k or cand_l in k:
                return v
    return None


def parse_depositary(xlsx_bytes: bytes) -> Tuple[List[Dict[str, Any]], List[List[Any]], List[str]]:
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb["Data"] if "Data" in wb.sheetnames else wb[wb.sheetnames[0]]
    header_row = find_header_row(ws, ["isin", "títulos"], max_scan=30) or find_header_row(ws, ["isin", "nominal"], max_scan=30)
    if not header_row:
        raise ValueError("No encuentro la fila de encabezados del depositario. Debe incluir ISIN y Títulos/Nominal.")
    cols = column_map(ws, header_row)
    c_prod = find_col(cols, ["producto"])
    c_isin = find_col(cols, ["isin"])
    c_div = find_col(cols, ["divisa"])
    c_nom = find_col(cols, ["títulos / nominal", "titulos / nominal", "nominal"])
    c_coste = find_col(cols, ["coste medio"])
    c_precio = find_col(cols, ["precio de mercado"])
    c_vm = find_col(cols, ["valor de mercado"])
    c_eur = find_col(cols, ["contravalor en eur", "contravalor"])
    c_costeadq = find_col(cols, ["coste adquisición", "coste adquisicion"])
    c_pm = find_col(cols, ["plus minus", "plusvalía", "plusvalia"])
    c_fx = find_col(cols, ["tasa de cambio"])
    c_mkt = find_col(cols, ["mercado"])

    positions: List[Dict[str, Any]] = []
    output_rows: List[List[Any]] = []
    headers = [norm_text(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]

    for r in range(header_row + 1, ws.max_row + 1):
        isin = clean_isin(ws.cell(r, c_isin).value if c_isin else None)
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in vals):
            output_rows.append(vals)
        if not isin:
            continue
        prod = norm_text(ws.cell(r, c_prod).value if c_prod else "")
        market = norm_text(ws.cell(r, c_mkt).value if c_mkt else "")
        positions.append({
            "isin": isin,
            "name": prod,
            "currency": norm_text(ws.cell(r, c_div).value if c_div else ""),
            "nominal": as_float(ws.cell(r, c_nom).value if c_nom else None, 0.0),
            "price": as_float(ws.cell(r, c_precio).value if c_precio else None, None),
            "market_value_local": as_float(ws.cell(r, c_vm).value if c_vm else None, None),
            "market_value_eur": as_float(ws.cell(r, c_eur).value if c_eur else None, None),
            "cost_avg": as_float(ws.cell(r, c_coste).value if c_coste else None, None),
            "cost_acq": as_float(ws.cell(r, c_costeadq).value if c_costeadq else None, None),
            "pnl": as_float(ws.cell(r, c_pm).value if c_pm else None, None),
            "fx": as_float(ws.cell(r, c_fx).value if c_fx else None, None),
            "market": market,
            "category_dep": infer_depositary_category(prod, market),
        })
    return positions, output_rows, headers


def parse_bloomberg(xlsx_bytes: bytes) -> List[Dict[str, Any]]:
    """Lee posiciones Bloomberg de forma más flexible.

    La v2.3 dependía demasiado de una exportación Bloomberg concreta
    —ISIN en AP, valor de mercado en F, nominal en AN—. Si Bloomberg cambia
    ligeramente el layout, la app no encuentra los ISIN y casi todo aparece
    como "Solo depositario". Esta versión escanea toda la fila en busca de
    ISIN válidos y usa cabeceras/fallbacks para localizar métricas.
    """
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb["Worksheet"] if "Worksheet" in wb.sheetnames else wb[wb.sheetnames[0]]

    max_col = ws.max_column
    max_row = ws.max_row

    def row_text(r: int) -> str:
        vals = []
        for c in range(1, min(max_col, 80) + 1):
            v = norm_text(ws.cell(r, c).value)
            if v:
                vals.append(v)
        return " ".join(vals)

    # Mapa aproximado de cabeceras combinando las primeras filas útiles.
    header_blob_by_col: Dict[int, str] = {}
    for c in range(1, max_col + 1):
        parts = []
        for r in range(1, min(max_row, 12) + 1):
            v = norm_text(ws.cell(r, c).value)
            if v:
                parts.append(v)
        header_blob_by_col[c] = " ".join(parts).lower()

    def find_col(*patterns: str, prefer_cart: bool = True, avoid_benchmark: bool = True) -> Optional[int]:
        candidates = []
        pats = [p.lower() for p in patterns]
        for c, blob in header_blob_by_col.items():
            if all(p in blob for p in pats):
                score = 0
                if prefer_cart and any(x in blob for x in ["cart", "portfolio", "cartera"]):
                    score += 10
                if avoid_benchmark and any(x in blob for x in ["bmrk", "benchmark", "índice", "indice"]):
                    score -= 10
                candidates.append((score, c))
        if not candidates:
            return None
        return sorted(candidates, reverse=True)[0][1]

    col_market = (
        find_col("valmrc")
        or find_col("market", "value")
        or find_col("valor", "mercado")
        or find_col("valor", "merc")
    )
    col_nominal = find_col("nominal") or find_col("face") or find_col("par")
    col_ytm = find_col("rend") or find_col("yield") or find_col("ytm")
    col_years = find_col("años") or find_col("years") or find_col("maturity")
    col_rating = find_col("rating")
    col_duration = find_col("dur") or find_col("duration")
    col_spread = find_col("spread")
    col_coupon = find_col("cup") or find_col("cupón") or find_col("cupon")

    def val_at(r: int, col: Optional[int], default: Optional[float] = None) -> Optional[float]:
        if not col:
            return default
        return as_float(ws.cell(r, col).value, default)

    def text_at(r: int, col: Optional[int]) -> str:
        if not col:
            return ""
        return norm_text(ws.cell(r, col).value)

    def first_row_isin(r: int) -> str:
        for c in range(1, max_col + 1):
            isin = clean_isin(ws.cell(r, c).value)
            if isin:
                return isin
        return ""

    def likely_market_value(r: int, nominal_guess: Optional[float]) -> Optional[float]:
        # 1) Cabecera detectada.
        v = val_at(r, col_market, None)
        if v is not None:
            return v
        # 2) Fallback del layout original.
        if max_col >= 6:
            v = as_float(ws.cell(r, 6).value, None)
            if v is not None:
                return v
        # 3) Último recurso: primer importe razonable no idéntico al nominal.
        nums = []
        for c in range(1, max_col + 1):
            x = as_float(ws.cell(r, c).value, None)
            if x is not None and abs(x) > 1000:
                nums.append(x)
        if not nums:
            return None
        if nominal_guess:
            diffs = [(abs(x - nominal_guess), x) for x in nums if abs(x - nominal_guess) > 1]
            if diffs:
                # En bonos, el valor mercado suele estar cerca del nominal, no necesariamente igual.
                return sorted(diffs, key=lambda t: t[0])[0][1]
        return nums[0]

    positions: List[Dict[str, Any]] = []
    seen: set[str] = set()
    current_category = "Otros"

    for r in range(1, max_row + 1):
        isin = first_row_isin(r)
        text = row_text(r)

        # Actualización de categoría en filas sin ISIN.
        if not isin and text:
            cat = spanish_category(text, "")
            # Evitamos que cabeceras genéricas cambien la categoría.
            if cat and not any(h in text.lower() for h in ["isin", "nominal", "rating", "duration"]):
                current_category = cat
            continue

        if not isin or isin in seen:
            continue
        seen.add(isin)

        # Nombre: priorizamos columna A/B si tienen texto no-ISIN; si no, primer texto útil de la fila.
        name = ""
        for c in [1, 2, 3]:
            candidate = norm_text(ws.cell(r, c).value)
            if candidate and not clean_isin(candidate):
                name = candidate
                break
        if not name:
            for c in range(1, min(max_col, 15) + 1):
                candidate = norm_text(ws.cell(r, c).value)
                if candidate and not clean_isin(candidate) and not as_float(candidate, None):
                    name = candidate
                    break
        if not name:
            name = isin

        nominal = val_at(r, col_nominal, None)
        if nominal is None and max_col >= 40:
            nominal = as_float(ws.cell(r, 40).value, None)

        mv = likely_market_value(r, nominal)

        positions.append({
            "isin": isin,
            "name": name,
            "issuer": norm_text(ws.cell(r, 2).value),
            "currency": "EUR",
            "weight_pct_bbg": as_float(ws.cell(r, 4).value, None) if max_col >= 4 else None,
            "market_value_bbg": mv,
            "ytm": val_at(r, col_ytm, None) if col_ytm else (as_float(ws.cell(r, 8).value, None) if max_col >= 8 else None),
            "years": val_at(r, col_years, None) if col_years else (as_float(ws.cell(r, 14).value, None) if max_col >= 14 else None),
            "rating": text_at(r, col_rating) if col_rating else (norm_text(ws.cell(r, 16).value) if max_col >= 16 else ""),
            "duration": val_at(r, col_duration, None) if col_duration else (as_float(ws.cell(r, 18).value, None) if max_col >= 18 else None),
            "spread": val_at(r, col_spread, None) if col_spread else (as_float(ws.cell(r, 36).value, None) if max_col >= 36 else None),
            "nominal_bbg": nominal,
            "coupon": val_at(r, col_coupon, None) if col_coupon else (as_float(ws.cell(r, 50).value, None) if max_col >= 50 else None),
            "category": current_category,
            "source": "Bloomberg",
        })

    return positions

def maturity_from_text(name: str) -> Optional[date]:
    # Formatos habituales: 310726, 31/07/26, 05/19/27, 08/15/34
    s = name.upper()
    m = re.search(r"\b(\d{2})(\d{2})(\d{2})\b", s)
    if m:
        dd, mm, yy = map(int, m.groups())
        year = 2000 + yy
        try:
            return date(year, mm, dd)
        except Exception:
            pass
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", s)
    if m:
        a, b, y = m.groups()
        a, b = int(a), int(b)
        year = int(y) if len(y) == 4 else 2000 + int(y)
        # Bloomberg suele MM/DD/YY en bonos; si el primer número >12, asumimos DD/MM/YY
        mm, dd = (a, b) if a <= 12 else (b, a)
        try:
            return date(year, mm, dd)
        except Exception:
            pass
    return None


def estimate_bill_yield(dep: Dict[str, Any], asof: Optional[date] = None) -> Tuple[Optional[float], Optional[float]]:
    asof = asof or date.today()
    mat = maturity_from_text(dep.get("name", ""))
    nominal = as_float(dep.get("nominal"), None)
    mv = as_float(dep.get("market_value_eur"), None) or as_float(dep.get("market_value_local"), None)
    if not mat or not nominal or not mv or mv <= 0:
        return None, None
    days = max((mat - asof).days, 1)
    years = days / 365.0
    ytm = ((nominal / mv) ** (1 / years) - 1) * 100
    return ytm, years


def build_conciliation(dep_positions: List[Dict[str, Any]], bbg_positions: List[Dict[str, Any]], cash_manual: float) -> Tuple[List[Dict[str, Any]], pd.DataFrame, pd.DataFrame]:
    dep_by_isin = {p["isin"]: p for p in dep_positions}
    bbg_by_isin = {p["isin"]: p for p in bbg_positions}
    all_isins = sorted(set(dep_by_isin) | set(bbg_by_isin))

    rows: List[Dict[str, Any]] = []
    for isin in all_isins:
        dep = dep_by_isin.get(isin)
        bbg = bbg_by_isin.get(isin)
        if bbg:
            category = bbg.get("category") or "Otros"
            name = bbg.get("name") or (dep or {}).get("name")
            mv_bbg = as_float(bbg.get("market_value_bbg"), None)
            nominal = as_float(bbg.get("nominal_bbg"), None) or as_float((dep or {}).get("nominal"), 0.0)
            market_value = mv_bbg if mv_bbg is not None else as_float((dep or {}).get("market_value_eur"), 0.0)
            row = {
                "ISIN": isin,
                "Bono": name,
                "Divisa": (dep or {}).get("currency") or bbg.get("currency") or "",
                "Emisor": bbg.get("issuer") or "",
                "Categoría": category,
                "Valor Mercado": market_value,
                "Rendimiento a Vencimiento": bbg.get("ytm"),
                "Años a vencimiento": bbg.get("years"),
                "Rating": bbg.get("rating"),
                "Duración": bbg.get("duration"),
                "Spread sobre Swap": bbg.get("spread"),
                "Cupón": bbg.get("coupon"),
                "Nominal": nominal,
                "Valor Depositario": as_float((dep or {}).get("market_value_eur"), None),
                "Valor Bloomberg": mv_bbg,
                "Nominal Depositario": as_float((dep or {}).get("nominal"), None),
                "Nominal Bloomberg": as_float(bbg.get("nominal_bbg"), None),
                "Fuente": "Ambos" if dep else "Solo Bloomberg",
            }
        else:
            category = (dep or {}).get("category_dep", "Solo depositario")
            ytm, years = estimate_bill_yield(dep or {}) if category == "Pagarés" else (None, None)
            row = {
                "ISIN": isin,
                "Bono": (dep or {}).get("name", ""),
                "Divisa": (dep or {}).get("currency", ""),
                "Emisor": "",
                "Categoría": category,
                "Valor Mercado": as_float((dep or {}).get("market_value_eur"), 0.0),
                "Rendimiento a Vencimiento": ytm,
                "Años a vencimiento": years,
                "Rating": "",
                "Duración": years if category == "Pagarés" else None,
                "Spread sobre Swap": None,
                "Cupón": 0 if category == "Pagarés" else None,
                "Nominal": as_float((dep or {}).get("nominal"), 0.0),
                "Valor Depositario": as_float((dep or {}).get("market_value_eur"), None),
                "Valor Bloomberg": None,
                "Nominal Depositario": as_float((dep or {}).get("nominal"), None),
                "Nominal Bloomberg": None,
                "Fuente": "Solo depositario",
            }
        row["Diferencia"] = (row.get("Valor Depositario") or 0) - (row.get("Valor Bloomberg") or row.get("Valor Mercado") or 0)
        rows.append(row)

    if cash_manual and abs(cash_manual) > 0:
        rows.insert(0, {
            "ISIN": "",
            "Bono": "EURO",
            "Divisa": "EUR",
            "Emisor": "",
            "Categoría": "Efectivo",
            "Valor Mercado": cash_manual,
            "Rendimiento a Vencimiento": None,
            "Años a vencimiento": 0,
            "Rating": "",
            "Duración": 0,
            "Spread sobre Swap": 0,
            "Cupón": 0,
            "Nominal": cash_manual,
            "Valor Depositario": cash_manual,
            "Valor Bloomberg": None,
            "Nominal Depositario": cash_manual,
            "Nominal Bloomberg": None,
            "Fuente": "Efectivo",
            "Diferencia": 0,
        })

    total = sum(as_float(r.get("Valor Mercado"), 0.0) or 0 for r in rows)
    for r in rows:
        r["Ponderación"] = (as_float(r.get("Valor Mercado"), 0.0) or 0) / total if total else 0
        abs_diff = abs(as_float(r.get("Diferencia"), 0.0) or 0)
        if r["Fuente"] == "Solo depositario":
            status = "SOLO DEPOSITARIO"
        elif r["Fuente"] == "Solo Bloomberg":
            status = "SOLO BLOOMBERG"
        elif abs_diff > 1000:
            status = "REVISAR"
        else:
            status = "OK"
        r["Estado"] = status

    df = pd.DataFrame(rows)
    summary = df.groupby("Categoría", dropna=False)["Valor Mercado"].sum().reset_index()
    summary["Peso"] = summary["Valor Mercado"] / total if total else 0
    summary["Orden"] = summary["Categoría"].apply(lambda x: CATEGORY_ORDER.index(x) if x in CATEGORY_ORDER else 999)
    summary = summary.sort_values("Orden").drop(columns="Orden")
    return rows, df, summary


def copy_row_style(ws, source_row: int, target_row: int, max_col: int = 28):
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for c in range(1, max_col + 1):
        src = ws.cell(source_row, c)
        dst = ws.cell(target_row, c)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)


def clear_values(ws, min_row: int, max_row: int, min_col: int, max_col: int):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            cell.value = None


def safe_sheet(wb, name: str):
    return wb[name] if name in wb.sheetnames else wb.create_sheet(name)


def write_template_output(template_bytes: bytes, rows: List[Dict[str, Any]], dep_positions: List[Dict[str, Any]], summary: pd.DataFrame, control_df: pd.DataFrame, include_control: bool = False) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = safe_sheet(wb, "Worksheet")
    inv = safe_sheet(wb, "Cartera Inversis")
    aux = safe_sheet(wb, "Hoja1")

    # 1) Cabeceras en español, sin reconstruir el estilo de la plantilla.
    headers = ["ISIN", "Bono", "Divisa", "Emisor", "Ponderación", "Valor Mercado", None,
               "Rendimiento a Vencimiento", "Años a vencimiento", "Rating", "Duración", "Spread sobre Swap",
               "ISIN", "Rating Emisor", "Cupón", "Nominal", None, "Valor Depositario", "Diferencia"]
    for c, v in enumerate(headers, 1):
        if v is not None:
            ws.cell(2, c).value = v
    ws.cell(1, 4).value = "NOMBRE"
    ws.cell(1, 8).value = "TIR A VENCIMIENTO"

    # 2) Limpiar solo valores de la zona dinámica; se mantienen columnas, anchos, estilos y estructura visual.
    clear_values(ws, 3, max(ws.max_row, 220), 1, 28)

    # Fila estilo de plantilla: 3 total, 4 categoría, 5 efectivo, 7 detalle.
    total_style_row, cat_style_row, cash_style_row, detail_style_row = 3, 4, 5, 7

    # Agrupación ordenada.
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for r in rows:
        grouped.setdefault(r.get("Categoría") or "Otros", []).append(r)
    ordered_cats = [c for c in CATEGORY_ORDER if c in grouped] + sorted([c for c in grouped if c not in CATEGORY_ORDER])

    total_value = sum(as_float(r.get("Valor Mercado"), 0.0) or 0 for r in rows)
    current_row = 3

    # Fila total.
    copy_row_style(ws, total_style_row, current_row)
    ws.cell(current_row, 2).value = "JAVEA AUGUSTA S.L"
    ws.cell(current_row, 5).value = 1
    ws.cell(current_row, 6).value = total_value
    ws.cell(current_row, 8).value = weighted_avg(rows, "Rendimiento a Vencimiento", "Valor Mercado")
    ws.cell(current_row, 9).value = weighted_avg(rows, "Años a vencimiento", "Valor Mercado")
    ws.cell(current_row, 10).value = ""
    ws.cell(current_row, 11).value = weighted_avg(rows, "Duración", "Valor Mercado")
    ws.cell(current_row, 12).value = weighted_avg(rows, "Spread sobre Swap", "Valor Mercado")
    ws.cell(current_row, 15).value = weighted_avg(rows, "Cupón", "Valor Mercado")
    ws.cell(current_row, 16).value = sum(as_float(r.get("Nominal"), 0.0) or 0 for r in rows)
    current_row += 1

    for cat in ordered_cats:
        cat_rows = grouped[cat]
        start_detail = current_row + 1
        end_detail = start_detail + len(cat_rows) - 1

        copy_row_style(ws, cat_style_row, current_row)
        ws.cell(current_row, 2).value = f"› {cat}"
        ws.cell(current_row, 5).value = sum((as_float(r.get("Valor Mercado"), 0.0) or 0) for r in cat_rows) / total_value if total_value else 0
        ws.cell(current_row, 6).value = sum((as_float(r.get("Valor Mercado"), 0.0) or 0) for r in cat_rows)
        ws.cell(current_row, 8).value = weighted_avg(cat_rows, "Rendimiento a Vencimiento", "Valor Mercado")
        ws.cell(current_row, 9).value = weighted_avg(cat_rows, "Años a vencimiento", "Valor Mercado")
        ws.cell(current_row, 10).value = category_rating(cat_rows)
        ws.cell(current_row, 11).value = weighted_avg(cat_rows, "Duración", "Valor Mercado")
        ws.cell(current_row, 12).value = weighted_avg(cat_rows, "Spread sobre Swap", "Valor Mercado")
        ws.cell(current_row, 15).value = weighted_avg(cat_rows, "Cupón", "Valor Mercado")
        ws.cell(current_row, 16).value = sum(as_float(r.get("Nominal"), 0.0) or 0 for r in cat_rows)
        current_row += 1

        for item in cat_rows:
            style_row = cash_style_row if cat == "Efectivo" else detail_style_row
            copy_row_style(ws, style_row, current_row)
            ws.cell(current_row, 1).value = item.get("ISIN") or None
            ws.cell(current_row, 2).value = "      " + norm_text(item.get("Bono")) if cat != "Efectivo" else norm_text(item.get("Bono"))
            ws.cell(current_row, 3).value = item.get("Divisa")
            ws.cell(current_row, 4).value = item.get("Emisor")
            ws.cell(current_row, 5).value = item.get("Ponderación")
            ws.cell(current_row, 6).value = item.get("Valor Mercado")
            ws.cell(current_row, 8).value = item.get("Rendimiento a Vencimiento")
            ws.cell(current_row, 9).value = item.get("Años a vencimiento")
            ws.cell(current_row, 10).value = item.get("Rating")
            ws.cell(current_row, 11).value = item.get("Duración")
            ws.cell(current_row, 12).value = item.get("Spread sobre Swap")
            ws.cell(current_row, 13).value = item.get("ISIN") or None
            ws.cell(current_row, 15).value = item.get("Cupón")
            ws.cell(current_row, 16).value = item.get("Nominal")
            ws.cell(current_row, 18).value = item.get("Valor Depositario")
            ws.cell(current_row, 19).value = item.get("Diferencia")
            current_row += 1

    last_row = current_row - 1
    # Formatos numéricos principales; no cambiamos colores/anchos.
    for r in range(3, last_row + 1):
        ws.cell(r, 5).number_format = "0.00%"
        for c in [6, 16, 18, 19]:
            ws.cell(r, c).number_format = '#,##0.00'
        for c in [8, 9, 11, 12, 15]:
            ws.cell(r, c).number_format = '0.00'

    # 3) Cartera Inversis: mantener pestaña y estilo, volcar posición del depositario en español.
    # El modelo contiene algunos rangos combinados en filas de totales; para escribir de forma segura,
    # descombinamos únicamente esta pestaña operativa. Worksheet queda intacta.
    for merged_range in list(inv.merged_cells.ranges):
        inv.unmerge_cells(str(merged_range))
    clear_values(inv, 1, max(inv.max_row, len(dep_positions) + 12), 1, 15)
    inv.cell(1, 1).value = "TOTAL POSICIÓN DEPOSITARIO"
    inv.cell(1, 2).value = "PERSONA"
    inv.cell(2, 1).value = sum(as_float(p.get("market_value_eur"), 0.0) or 0 for p in dep_positions)
    inv.cell(4, 1).value = "Mi inversión"
    dep_headers = ["Producto", "ISIN", "Divisa", "Títulos / Nominal", "Tit > 10 años", "Coste medio", "Precio de Mercado", "Valor de mercado", "Plusvalía / Minusvalía", "% Ganancia / Pérdida", "Contravalor en EUR", "% en Cartera", "Coste de adquisición", "Coste medio en EUR", "Descripción Bloomberg"]
    for c, h in enumerate(dep_headers, 1):
        inv.cell(6, c).value = h
    dep_total = sum(as_float(p.get("market_value_eur"), 0.0) or 0 for p in dep_positions)
    for i, p in enumerate(dep_positions, start=8):
        copy_row_style(inv, 8 if inv.max_row >= 8 else 6, i, max_col=15)
        inv.cell(i, 1).value = p.get("name")
        inv.cell(i, 2).value = p.get("isin")
        inv.cell(i, 3).value = p.get("currency")
        inv.cell(i, 4).value = p.get("nominal")
        inv.cell(i, 5).value = 0
        inv.cell(i, 6).value = p.get("cost_avg")
        inv.cell(i, 7).value = p.get("price")
        inv.cell(i, 8).value = p.get("market_value_local")
        inv.cell(i, 9).value = p.get("pnl")
        inv.cell(i, 10).value = None
        inv.cell(i, 11).value = p.get("market_value_eur")
        inv.cell(i, 12).value = (as_float(p.get("market_value_eur"), 0.0) or 0) / dep_total if dep_total else 0
        inv.cell(i, 13).value = p.get("cost_acq")
        inv.cell(i, 14).value = p.get("cost_avg")
        inv.cell(i, 15).value = p.get("name")
    for r in range(8, 8 + len(dep_positions)):
        inv.cell(r, 12).number_format = "0.00%"
        for c in [4, 6, 7, 8, 9, 11, 13, 14]:
            inv.cell(r, c).number_format = '#,##0.00'

    # 4) Hoja1: resumen y gráfico de tarta. No tocamos Worksheet para preservar su layout.
    for row in aux.iter_rows(min_row=1, max_row=max(aux.max_row, 80), min_col=1, max_col=8):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            cell.value = None
    aux.cell(1, 1).value = "Distribución por tipo de activo"
    aux.cell(3, 1).value = "Tipo de activo"
    aux.cell(3, 2).value = "Valor Mercado"
    aux.cell(3, 3).value = "Peso"
    for cell in aux[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    start = 4
    for idx, rec in enumerate(summary.to_dict("records"), start=start):
        aux.cell(idx, 1).value = rec["Categoría"]
        aux.cell(idx, 2).value = float(rec["Valor Mercado"])
        aux.cell(idx, 3).value = float(rec["Peso"])
        aux.cell(idx, 2).number_format = '#,##0.00'
        aux.cell(idx, 3).number_format = '0.00%'
    end = start + len(summary) - 1
    aux.column_dimensions["A"].width = 24
    aux.column_dimensions["B"].width = 18
    aux.column_dimensions["C"].width = 12

    # borrar charts anteriores en Hoja1 y añadir uno nuevo
    aux._charts = []
    if end >= start:
        pie = PieChart()
        labels = Reference(aux, min_col=1, min_row=start, max_row=end)
        data = Reference(aux, min_col=2, min_row=3, max_row=end)
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(labels)
        pie.title = "Distribución por tipo de activo"
        pie.height = 9
        pie.width = 13
        aux.add_chart(pie, "E3")

    # 5) Control de conciliación opcional.
    if include_control:
        if "Control Conciliación" in wb.sheetnames:
            del wb["Control Conciliación"]
        ctrl = wb.create_sheet("Control Conciliación")
        cols = ["Estado", "Fuente", "Categoría", "ISIN", "Bono", "Valor Depositario", "Valor Bloomberg", "Valor Mercado", "Diferencia", "Nominal Depositario", "Nominal Bloomberg"]
        for c, h in enumerate(cols, 1):
            ctrl.cell(1, c).value = h
            ctrl.cell(1, c).font = Font(bold=True, color="FFFFFF")
            ctrl.cell(1, c).fill = PatternFill("solid", fgColor="1F4E78")
        for r_idx, rec in enumerate(control_df.to_dict("records"), 2):
            for c_idx, col in enumerate(cols, 1):
                ctrl.cell(r_idx, c_idx).value = rec.get(col)
        for col in range(1, len(cols) + 1):
            ctrl.column_dimensions[get_column_letter(col)].width = 18

    # Print settings
    ws.sheet_view.showGridLines = False
    inv.sheet_view.showGridLines = False
    aux.sheet_view.showGridLines = False

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def weighted_avg(rows: List[Dict[str, Any]], val_col: str, weight_col: str) -> Optional[float]:
    num = 0.0
    den = 0.0
    for r in rows:
        v = as_float(r.get(val_col), None)
        w = as_float(r.get(weight_col), None)
        if v is not None and w is not None and w != 0:
            num += v * w
            den += w
    return num / den if den else None


def category_rating(rows: List[Dict[str, Any]]) -> str:
    vals = [norm_text(r.get("Rating")) for r in rows if norm_text(r.get("Rating"))]
    if not vals:
        return ""
    # mantener simple: rating más frecuente
    return max(set(vals), key=vals.count)


def make_pdf_report(control_df: pd.DataFrame, summary: pd.DataFrame, cash_manual: float) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=1.2*cm, leftMargin=1.2*cm, topMargin=1.0*cm, bottomMargin=1.0*cm)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("TitleFA", parent=styles["Title"], textColor=colors.HexColor("#1F4E78"), fontSize=18, leading=22)
    normal = styles["BodyText"]
    elems = [Paragraph("Reporte de conciliación de cartera", title), Spacer(1, 0.25*cm)]

    total = float(summary["Valor Mercado"].sum()) if len(summary) else 0.0
    n_ok = int((control_df["Estado"] == "OK").sum()) if "Estado" in control_df else 0
    n_revisar = int((control_df["Estado"] == "REVISAR").sum()) if "Estado" in control_df else 0
    n_solo_dep = int((control_df["Estado"] == "SOLO DEPOSITARIO").sum()) if "Estado" in control_df else 0
    n_solo_bbg = int((control_df["Estado"] == "SOLO BLOOMBERG").sum()) if "Estado" in control_df else 0

    kpi_data = [
        ["Valor mercado total", f"{total:,.2f} €"],
        ["Efectivo", f"{cash_manual:,.2f} €"],
        ["Posiciones OK", n_ok],
        ["A revisar", n_revisar],
        ["Solo depositario", n_solo_dep],
        ["Solo Bloomberg", n_solo_bbg],
    ]
    t = Table(kpi_data, colWidths=[7*cm, 5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9EAF7")),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#C7D3DD")),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
    ]))
    elems += [t, Spacer(1, 0.4*cm)]

    elems += [Paragraph("Peso por magnitud", styles["Heading2"])]
    sum_data = [["Magnitud", "Valor mercado", "Peso"]] + [[r["Categoría"], f"{r['Valor Mercado']:,.2f}", f"{r['Peso']:.2%}"] for _, r in summary.iterrows()]
    t2 = Table(sum_data, colWidths=[6*cm, 5*cm, 3*cm])
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#C7D3DD")),
        ("ALIGN", (1,1), (-1,-1), "RIGHT"),
    ]))
    elems += [t2, Spacer(1, 0.4*cm)]

    if n_revisar or n_solo_dep or n_solo_bbg:
        elems += [Paragraph("Excepciones principales", styles["Heading2"])]
        exc = control_df[control_df["Estado"].isin(["REVISAR", "SOLO DEPOSITARIO", "SOLO BLOOMBERG"])].copy()
        exc = exc.head(12)
        exc_data = [["Estado", "ISIN", "Bono", "Diferencia"]]
        for _, r in exc.iterrows():
            exc_data.append([r.get("Estado", ""), r.get("ISIN", ""), str(r.get("Bono", ""))[:45], f"{as_float(r.get('Diferencia'), 0.0):,.2f}"])
        t3 = Table(exc_data, colWidths=[4*cm, 4*cm, 12*cm, 4*cm])
        t3.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#C7D3DD")),
            ("ALIGN", (3,1), (3,-1), "RIGHT"),
        ]))
        elems.append(t3)
    else:
        elems.append(Paragraph("No se han detectado excepciones relevantes con la tolerancia aplicada.", normal))

    doc.build(elems)
    buf.seek(0)
    return buf.getvalue()


def run_engine(depositary_bytes: bytes, bloomberg_bytes: bytes, template_bytes: bytes, cash_manual: float, include_control: bool) -> Tuple[bytes, bytes, pd.DataFrame, pd.DataFrame]:
    dep_positions, _, _ = parse_depositary(depositary_bytes)
    bbg_positions = parse_bloomberg(bloomberg_bytes)
    rows, control_df, summary = build_conciliation(dep_positions, bbg_positions, cash_manual)
    xlsx_out = write_template_output(template_bytes, rows, dep_positions, summary, control_df, include_control=include_control)
    pdf_out = make_pdf_report(control_df, summary, cash_manual)
    return xlsx_out, pdf_out, control_df, summary


def validate_xlsx_bytes(xlsx_bytes: bytes) -> Tuple[bool, str]:
    """Validación ligera para evitar ofrecer un Excel corrupto en la descarga."""
    try:
        if not xlsx_bytes or len(xlsx_bytes) < 1000:
            return False, "El archivo generado está vacío o incompleto."
        wb_test = load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=False)
        required = {"Worksheet", "Cartera Inversis", "Hoja1"}
        missing = required.difference(set(wb_test.sheetnames))
        wb_test.close()
        if missing:
            return False, f"Faltan hojas esperadas: {', '.join(sorted(missing))}."
        return True, "OK"
    except Exception as exc:
        return False, f"El Excel generado no supera la validación interna: {exc}"


def main():
    import streamlit as st
    st.set_page_config(page_title="Conciliador de carteras", layout="wide")
    st.title("Conciliador de carteras")
    st.caption("v2.4 · Parser Bloomberg flexible · Formato exacto del modelo · Efectivo · Gráfico por tipo de activo")

    st.markdown("""
    Sube los tres ficheros: depositario, Bloomberg y el Excel modelo de salida.  
    La app usa el modelo como **base real**, por lo que conserva la estructura visual, estilos, anchos de columna y pestañas originales.
    """)

    # Importante: guardamos los bytes generados en session_state. Así el botón de descarga no depende
    # de variables locales que puedan perderse con los reruns internos de Streamlit.
    for key in ["xlsx_out", "pdf_out", "control_df", "summary_df", "last_error"]:
        if key not in st.session_state:
            st.session_state[key] = None

    col1, col2, col3 = st.columns(3)
    with col1:
        dep_file = st.file_uploader("1. Excel depositario", type=["xlsx", "xlsm"], key="dep")
    with col2:
        bbg_file = st.file_uploader("2. Excel Bloomberg", type=["xlsx", "xlsm"], key="bbg")
    with col3:
        template_file = st.file_uploader("3. Excel plantilla/modelo de salida", type=["xlsx", "xlsm"], key="tpl")

    col_a, col_b = st.columns([1, 2])
    with col_a:
        cash_manual = st.number_input("Efectivo en EUR", min_value=0.0, value=0.0, step=1000.0, format="%.2f")
    with col_b:
        include_control = st.checkbox(
            "Añadir pestaña 'Control Conciliación'",
            value=False,
            help="Desactívala si quieres mantener solo las pestañas del modelo: Worksheet, Cartera Inversis y Hoja1."
        )

    generate = st.button("Generar conciliación", type="primary")

    if generate:
        # Limpiamos resultados anteriores para no descargar por error una conciliación vieja.
        st.session_state.xlsx_out = None
        st.session_state.pdf_out = None
        st.session_state.control_df = None
        st.session_state.summary_df = None
        st.session_state.last_error = None

        if not dep_file or not bbg_file or not template_file:
            st.error("Necesito los tres ficheros para generar la conciliación.")
            return
        try:
            with st.spinner("Generando Excel y PDF..."):
                xlsx_out, pdf_out, control_df, summary = run_engine(
                    dep_file.getvalue(), bbg_file.getvalue(), template_file.getvalue(), cash_manual, include_control
                )
                ok, msg = validate_xlsx_bytes(xlsx_out)
                if not ok:
                    raise ValueError(msg)

            st.session_state.xlsx_out = bytes(xlsx_out)
            st.session_state.pdf_out = bytes(pdf_out)
            st.session_state.control_df = control_df
            st.session_state.summary_df = summary
            st.success("Conciliación generada y validada correctamente.")
        except Exception as e:
            st.session_state.last_error = str(e)
            st.exception(e)

    if st.session_state.last_error:
        st.error(st.session_state.last_error)

    if st.session_state.xlsx_out is not None:
        control_df = st.session_state.control_df
        summary = st.session_state.summary_df

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Posiciones", len(control_df))
        c2.metric("OK", int((control_df["Estado"] == "OK").sum()))
        c3.metric("A revisar", int((control_df["Estado"] == "REVISAR").sum()))
        c4.metric("Solo depositario", int((control_df["Estado"] == "SOLO DEPOSITARIO").sum()))

        st.subheader("Resumen por magnitud")
        st.dataframe(summary, use_container_width=True)
        st.subheader("Control de conciliación")
        st.dataframe(control_df, use_container_width=True)

        st.download_button(
            "Descargar Excel conciliado",
            data=st.session_state.xlsx_out,
            file_name="cartera_conciliada_formato_modelo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_conciliado",
        )
        st.download_button(
            "Descargar reporte PDF",
            data=st.session_state.pdf_out,
            file_name="reporte_conciliacion.pdf",
            mime="application/pdf",
            key="download_pdf_conciliacion",
        )


if __name__ == "__main__":
    main()
