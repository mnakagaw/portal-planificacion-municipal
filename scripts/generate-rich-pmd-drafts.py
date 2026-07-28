#!/usr/bin/env python3
"""Generate substantive, source-traceable PMD draft DOCX files.

This is the second-generation builder for the Paquete Mínimo. It turns the
territorial dashboard into a five-page visual diagnostic and adds municipal
narrative based only on:

* the dashboard's official statistical datasets;
* locally archived historical PMDs;
* Spanish Wikipedia as an explicitly secondary source.

Información General and Diagnóstico are written as near-final technical text.
FODA, vision and projects remain proposals for OMPP/CDM validation. The script
does not invent meetings, participation, approvals, budgets, beneficiaries,
authorities or execution status.
"""

from __future__ import annotations

import argparse
import csv
import importlib.util
import json
import math
import re
import shutil
import sys
import textwrap
import time
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import fitz
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Polygon

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


SCRIPT_DIR = Path(__file__).resolve().parent
BASE_GENERATOR_PATH = SCRIPT_DIR / "generate-simple-pmd-drafts.py"
spec = importlib.util.spec_from_file_location("simple_pmd_builder", BASE_GENERATOR_PATH)
if spec is None or spec.loader is None:
    raise RuntimeError(f"Could not load {BASE_GENERATOR_PATH}")
base = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = base
spec.loader.exec_module(base)


PERIOD = "2025-2028"
TODAY = date(2026, 7, 28)
USER_AGENT = "DDPT-PMD-Draft-Builder/2.0 (municipal planning research)"

WIKIPEDIA_TITLE_OVERRIDES = {
    base.territory_key("Salcedo", "Hermanas Mirabal"): "Salcedo (República Dominicana)",
}

COLORS = {
    "ink": "203740",
    "muted": "5E7076",
    "teal": "14866D",
    "blue": "2F6CC6",
    "purple": "7651B8",
    "orange": "C96B21",
    "rose": "C9526A",
    "gold": "A27719",
    "line": "D7E0E1",
    "soft": "F3F7F6",
    "white": "FFFFFF",
    "warn": "9A5A18",
    "warn_soft": "FFF6E7",
}

SERVICE_LABELS = {
    "servicios_sanitarios": {
        "inodoro": "Inodoro",
        "letrina": "Letrina",
        "no_tiene": "No tiene",
        "sin_informacion": "Sin información",
    },
    "agua_uso_domestico": {
        "del_acueducto_dentro_de_la_vivienda": "Acueducto dentro",
        "del_acueducto_en_el_patio_de_la_vivienda": "Acueducto en patio",
        "de_una_llave_publica": "Llave pública",
        "de_una_llave_de_otra_vivienda": "Llave de otra vivienda",
        "de_un_tubo_de_la_calle": "Tubo de la calle",
        "manantial_rio_arroyo": "Manantial / río",
        "pozo_tubular": "Pozo tubular",
        "pozo_cavado": "Pozo cavado",
        "lluvia": "Agua de lluvia",
        "camion_tanque": "Camión tanque",
        "otro": "Otro",
    },
    "agua_para_beber": {
        "del_acueducto_dentro_de_la_vivienda": "Acueducto dentro",
        "del_acueducto_en_el_patio_de_la_vivienda": "Acueducto en patio",
        "de_una_llave_publica": "Llave pública",
        "de_una_llave_de_otra_vivienda": "Llave de otra vivienda",
        "manantial_rio_arroyo": "Manantial / río",
        "pozo_tubular": "Pozo tubular",
        "pozo_cavado": "Pozo cavado",
        "lluvia": "Agua de lluvia",
        "camion_tanque": "Camión tanque",
        "botellones": "Botellones",
        "camioncito_procesada": "Camioncito procesada",
        "otro": "Otro",
    },
    "alumbrado": {
        "energia_electrica_del_tendido_publico": "Red pública",
        "lampara_de_gas_propano": "Lámpara de gas",
        "lampara_de_gas_kerosene": "Lámpara de kerosene",
        "energia_electrica_de_planta_propia": "Planta propia",
        "paneles_solares": "Paneles solares",
        "otros": "Otros",
        "sin_informacion": "Sin información",
    },
    "combustible_cocinar": {
        "gas_propano": "Gas propano",
        "carbon": "Carbón",
        "lena": "Leña",
        "electricidad": "Electricidad",
        "otro": "Otro",
        "no_cocina": "No cocina",
        "sin_informacion": "Sin información",
    },
    "eliminacion_basura": {
        "la_recoge_el_ayuntamiento": "Recoge ayuntamiento",
        "la_recoge_una_empresa_privada": "Empresa privada",
        "la_queman": "La queman",
        "la_tiran_en_el_patio_o_sola": "Patio / solar",
        "la_tiran_en_un_vertedero": "Vertedero",
        "la_tiran_en_un_rio_o_canada": "Río / cañada",
        "otros": "Otros",
    },
}


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def save_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(value, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def fmt_number(value: float | int | None, decimals: int = 0) -> str:
    if value is None:
        return "No disponible"
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return "No disponible"
    return f"{float(value):,.{decimals}f}"


def fmt_pct(value: float | int | None, decimals: int = 1) -> str:
    if value is None:
        return "No disponible"
    return f"{float(value):.{decimals}f}%"


def pct(part: float | int | None, total: float | int | None) -> float | None:
    if part is None or total in (None, 0):
        return None
    return float(part) * 100.0 / float(total)


def sentences(text: str) -> list[str]:
    cleaned = re.sub(r"\s+", " ", text or "").strip()
    cleaned = re.sub(r"\[\s*cita requerida\s*\]", "", cleaned, flags=re.I)
    return [
        item.strip()
        for item in re.split(r"(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÑ0-9«])", cleaned)
        if 25 <= len(item.strip()) <= 520
    ]


def split_wikipedia_sections(extract: str) -> dict[str, str]:
    parts = re.split(r"\n+==+\s*([^=\n]+?)\s*==+\n+", extract or "")
    result = {"Introducción": parts[0].strip() if parts else ""}
    for index in range(1, len(parts) - 1, 2):
        result[parts[index].strip()] = parts[index + 1].strip()
    return result


def is_municipal_wikipedia(row: dict[str, Any]) -> bool:
    title = str(row.get("title") or "").strip()
    return row.get("status") == "verified" and not base.clean(title).startswith("provinciade")


def sanitize_wikipedia_cache(cache: dict[str, Any], targets: list[dict[str, Any]]) -> None:
    for item in targets:
        key = base.territory_key(item["municipio"], item["provincia"])
        override = WIKIPEDIA_TITLE_OVERRIDES.get(key)
        row = cache.get(key, {})
        if override:
            if row.get("title") != override:
                cache[key] = {
                    "status": "verified",
                    "title": override,
                    "url": "https://es.wikipedia.org/wiki/" + urllib.parse.quote(override.replace(" ", "_")),
                    "extract": "",
                    "retrieved_at": TODAY.isoformat(),
                }
            continue
        if row.get("status") == "verified" and not is_municipal_wikipedia(row):
            cache[key] = {
                "status": "rejected_nonmunicipal",
                "title": row.get("title", ""),
                "url": row.get("url", ""),
                "reason": "La página corresponde a una provincia, no al municipio.",
                "retrieved_at": TODAY.isoformat(),
            }


def fetch_full_wikipedia(cache: dict[str, Any], targets: list[dict[str, Any]]) -> None:
    pending = []
    for item in targets:
        key = base.territory_key(item["municipio"], item["provincia"])
        row = cache.get(key, {})
        if is_municipal_wikipedia(row) and not row.get("full_extract") and row.get("title"):
            pending.append((key, row["title"]))

    # MediaWiki's TextExtracts extension can return an empty extract for all
    # but one page in a multi-title request. Fetch one verified municipal page
    # at a time so a successful lookup is never mistaken for missing content.
    for key, title in pending:
        params = {
            "action": "query",
            "prop": "extracts|info|revisions",
            "explaintext": "1",
            "titles": title,
            "inprop": "url",
            "rvprop": "ids",
            "format": "json",
            "formatversion": "2",
        }
        url = "https://es.wikipedia.org/w/api.php?" + urllib.parse.urlencode(params)
        payload = None
        for attempt in range(3):
            try:
                request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
                with urllib.request.urlopen(request, timeout=45) as response:
                    payload = json.load(response)
                break
            except Exception as exc:
                cache.setdefault(key, {})["full_extract_error"] = str(exc)[:180]
                time.sleep(3.0 if "429" in str(exc) else 1.0)
        if payload is None:
            continue

        page = next(iter(payload.get("query", {}).get("pages", [])), None)
        if not page or page.get("missing"):
            continue
        full_extract = page.get("extract", "") or ""
        cache[key]["full_extract"] = full_extract
        cache[key]["sections"] = split_wikipedia_sections(full_extract)
        cache[key]["revid"] = (page.get("revisions") or [{}])[0].get("revid")
        cache[key]["url"] = page.get("fullurl") or cache[key].get("url", "")
        cache[key]["retrieved_at"] = TODAY.isoformat()
        cache[key].pop("full_extract_error", None)
        time.sleep(0.75)


def extract_historical_context(
    rows: list[dict[str, Any]], cache: dict[str, Any], cache_key: str
) -> dict[str, Any]:
    if cache_key in cache:
        cached = cache[cache_key]
        cached["history"] = re.sub(r"(?<=\w)-\s+(?=\w)", "", cached.get("history", ""))
        cached["history"] = re.sub(
            r"(?i)\boriginalmente llamado originalmente\b",
            "Originalmente llamado",
            cached["history"],
        )
        if (
            cached["history"]
            and not re.search(r"(?i)¡?error!?|\.{12,}", cached["history"])
            and len(cached["history"]) >= 120
        ):
            return cached
        cache.pop(cache_key, None)
    result: dict[str, Any] = {"history": "", "history_pages": "", "strategic": "", "strategic_pages": ""}
    candidates = sorted(rows, key=lambda row: row.get("period_end") or 0, reverse=True)
    for row in candidates:
        path = Path(row.get("local_path", ""))
        if not path.exists():
            continue
        try:
            pdf = fitz.open(path)
        except Exception:
            continue
        page_texts: list[str] = []
        history_hits: list[int] = []
        strategy_hits: list[int] = []
        for index in range(pdf.page_count):
            text = re.sub(r"\s+", " ", pdf[index].get_text("text") or "").strip()
            page_texts.append(text)
            normalized = base.clean(text)
            if any(token in normalized for token in ("antecedenteshistoricos", "resenahistorica", "historiadelmunicipio")):
                history_hits.append(index)
            if "lineasestrategicas" in normalized and len(strategy_hits) < 2:
                strategy_hits.append(index)
        if history_hits and not result["history"]:
            for index in history_hits:
                combined = " ".join(page_texts[index : min(index + 2, len(page_texts))])
                if re.search(r"(?i)¡?error!?|\.{12,}", combined):
                    continue
                combined = re.sub(r"(?<=\w)-\s+(?=\w)", "", combined)
                combined = re.sub(r"(?i)\boriginalmente llamado originalmente\b", "Originalmente llamado", combined)
                start = re.search(
                    r"(?i)(antecedentes\s+hist[oó]ricos|reseña\s+hist[oó]rica|historia\s+del\s+municipio)",
                    combined,
                )
                if start:
                    combined = combined[start.end() :]
                useful = [
                    sentence
                    for sentence in sentences(combined)
                    if not re.match(r"(?i)plan de desarrollo municipal|p[aá]gina \d+", sentence)
                ][:6]
                if len(" ".join(useful)) < 120:
                    continue
                result["history"] = " ".join(useful)[:2400]
                result["history_pages"] = f"{index + 1}-{min(index + 2, len(page_texts))}"
                result["history_period"] = f"{row.get('period_start')}-{row.get('period_end')}"
                result["history_source"] = row.get("source_url") or row.get("source_page") or str(path)
                break
        pdf.close()
        if result["history"]:
            break
    cache[cache_key] = result
    return result


def service_share(living: dict[str, Any] | None, service: str, category: str) -> float | None:
    if not living:
        return None
    item = living.get("servicios", {}).get(service, {})
    return pct(item.get("categorias", {}).get(category), item.get("total"))


def national_service_share(national: dict[str, Any], service: str, category: str) -> float | None:
    item = national.get(service, {})
    return pct(item.get("categorias", {}).get(category), item.get("total"))


def group_index(rows: list[dict[str, Any]], field: str = "adm2_code") -> dict[str, list[dict[str, Any]]]:
    result: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        result[base.code_key(row.get(field))].append(row)
    return result


def polygon_rings(geometry: dict[str, Any]) -> Iterable[list[list[float]]]:
    if geometry.get("type") == "Polygon":
        for ring in geometry.get("coordinates", [])[:1]:
            yield ring
    elif geometry.get("type") == "MultiPolygon":
        for polygon in geometry.get("coordinates", []):
            for ring in polygon[:1]:
                yield ring


def dashboard_source_note() -> str:
    return "Fuente: Censo 2022, DEE 2024, Anuario Estadístico Educativo 2024 y registros SNS disponibles en el Dashboard."


def add_card(
    fig,
    rect: tuple[float, float, float, float],
    title: str,
    *,
    accent: str = "#2F6CC6",
    face: str = "#FFFFFF",
    title_size: float = 9.2,
):
    x, y, width, height = rect
    patch = FancyBboxPatch(
        (x, y),
        width,
        height,
        transform=fig.transFigure,
        boxstyle="round,pad=0.006,rounding_size=0.009",
        linewidth=0.7,
        edgecolor="#D7E0E1",
        facecolor=face,
        zorder=0,
    )
    fig.patches.append(patch)
    fig.text(x + 0.015, y + height - 0.029, title, color=accent, fontsize=title_size, weight="bold")
    return (x + 0.018, y + 0.018, width - 0.036, height - 0.065)


def new_dashboard_page(title: str, municipality: str, subtitle: str = ""):
    fig = plt.figure(figsize=(8.5, 10.2), dpi=145, facecolor="#F7F9FA")
    fig.text(0.055, 0.965, title, fontsize=15, weight="bold", color="#203740")
    fig.text(0.055, 0.944, municipality, fontsize=9.5, color="#5E7076")
    if subtitle:
        fig.text(0.945, 0.95, subtitle, fontsize=7.5, color="#7A898D", ha="right")
    return fig


def finish_dashboard_page(fig, path: Path, source: str = "") -> None:
    fig.text(0.055, 0.026, source or dashboard_source_note(), fontsize=6.7, color="#75868B")
    fig.savefig(path, dpi=145, facecolor=fig.get_facecolor(), bbox_inches=None)
    plt.close(fig)


def render_map(
    fig,
    rect: tuple[float, float, float, float],
    geojson: dict[str, Any],
    adm2_code: str,
) -> None:
    ax = fig.add_axes(rect, zorder=2)
    selected = None
    bounds = [999.0, 999.0, -999.0, -999.0]
    for feature in geojson.get("features", []):
        if base.code_key(feature.get("properties", {}).get("adm2_code")) == adm2_code:
            selected = feature
        for ring in polygon_rings(feature.get("geometry", {})):
            xs = [point[0] for point in ring]
            ys = [point[1] for point in ring]
            if not xs:
                continue
            bounds[0] = min(bounds[0], min(xs))
            bounds[1] = min(bounds[1], min(ys))
            bounds[2] = max(bounds[2], max(xs))
            bounds[3] = max(bounds[3], max(ys))
            ax.add_patch(
                Polygon(
                    ring,
                    closed=True,
                    facecolor="#E7ECEC",
                    edgecolor="#FFFFFF",
                    linewidth=0.35,
                )
            )
    if selected:
        for ring in polygon_rings(selected.get("geometry", {})):
            ax.add_patch(
                Polygon(
                    ring,
                    closed=True,
                    facecolor="#E5574F",
                    edgecolor="#A62F2B",
                    linewidth=1.0,
                )
            )
    if bounds[0] < 900:
        ax.set_xlim(bounds[0] - 0.15, bounds[2] + 0.15)
        ax.set_ylim(bounds[1] - 0.08, bounds[3] + 0.08)
    ax.set_aspect("equal")
    ax.axis("off")


def draw_pyramid(ax, records: list[dict[str, Any]], title: str) -> None:
    if not records:
        ax.text(0.5, 0.5, "Datos no disponibles", ha="center", va="center", color="#75868B")
        ax.axis("off")
        return
    rows = [row for row in records if "declarado" not in str(row.get("age_group", "")).lower()]
    labels = [str(row.get("age_group", "")).replace(" - ", "–") for row in rows]
    male = [-float(row.get("male", 0) or 0) for row in rows]
    female = [float(row.get("female", 0) or 0) for row in rows]
    positions = list(range(len(rows)))
    ax.barh(positions, male, color="#2F6CC6", height=0.82)
    ax.barh(positions, female, color="#E94F4F", height=0.82)
    tick_step = 2 if len(labels) > 15 else 1
    ax.set_yticks(positions[::tick_step])
    ax.set_yticklabels(labels[::tick_step], fontsize=5.3)
    max_value = max([abs(value) for value in male + female] or [1])
    ax.set_xlim(-max_value * 1.12, max_value * 1.12)
    ax.axvline(0, color="#839196", linewidth=0.5)
    ax.grid(axis="x", color="#DCE3E5", linewidth=0.45)
    ax.tick_params(axis="x", labelsize=5.2)
    ax.set_title(title, fontsize=7.2, weight="bold", color="#203740", pad=4)
    for side in ax.spines.values():
        side.set_visible(False)


def text_rows(
    fig,
    rect: tuple[float, float, float, float],
    rows: list[tuple[str, str]],
    *,
    color: str = "#203740",
    max_rows: int = 8,
) -> None:
    x, y, width, height = rect
    visible = rows[:max_rows]
    line_height = height / max(1, len(visible))
    for index, (label, value) in enumerate(visible):
        yy = y + height - (index + 0.62) * line_height
        fig.text(x, yy, label, fontsize=6.9, color=color)
        fig.text(x + width, yy, value, fontsize=6.9, color=color, ha="right", weight="bold")


def top_categories(item: dict[str, Any], labels: dict[str, str], limit: int = 6) -> list[tuple[str, str]]:
    total = item.get("total") or 0
    categories = item.get("categorias", {})
    ordered = sorted(categories.items(), key=lambda pair: pair[1] or 0, reverse=True)
    result = []
    for key, value in ordered[:limit]:
        result.append((labels.get(key, key.replace("_", " ").title()), f"{fmt_number(value)} · {fmt_pct(pct(value, total))}"))
    return result


def generate_dashboard_pages(
    municipality: dict[str, Any],
    adm2_code: str,
    data: dict[str, Any],
    national: dict[str, Any],
    geojson: dict[str, Any],
    asset_dir: Path,
) -> list[Path]:
    asset_dir.mkdir(parents=True, exist_ok=True)
    name = municipality["municipio"]
    pages = [asset_dir / f"diagnostico-{index}.png" for index in range(1, 6)]
    basic = data.get("basic") or {}
    households = data.get("households") or {}
    urban = data.get("urban_rural") or {}
    living = data.get("living") or {}
    pyramid = data.get("pyramid") or []
    pyramid2010 = data.get("pyramid2010") or []
    household_size = data.get("household_size") or []
    education = data.get("education") or {}
    education_level = data.get("education_level") or {}
    education_offer = data.get("education_offer") or {}
    economy = data.get("economy") or {}
    tic = data.get("tic") or {}
    health = data.get("health") or {}

    # Page 1: demography and households
    fig = new_dashboard_page("Diagnóstico municipal · Población y hogares", name, "Censo 2022")
    cards = [
        ("Población total", fmt_number(basic.get("poblacion_total")), COLORS["blue"]),
        ("Variación 2010–2022", fmt_pct(basic.get("variacion_pct")), COLORS["teal"]),
        (
            "Hombres / mujeres",
            f"{fmt_pct(pct(basic.get('poblacion_hombres'), basic.get('poblacion_total')))} / "
            f"{fmt_pct(pct(basic.get('poblacion_mujeres'), basic.get('poblacion_total')))}",
            COLORS["purple"],
        ),
        (
            "Urbana / rural",
            f"{fmt_pct(pct(urban.get('urbana'), urban.get('poblacion_total')))} / "
            f"{fmt_pct(pct(urban.get('rural'), urban.get('poblacion_total')))}",
            COLORS["orange"],
        ),
    ]
    for index, (label, value, accent) in enumerate(cards):
        x = 0.055 + index * 0.225
        content = add_card(fig, (x, 0.835, 0.205, 0.082), label, accent=f"#{accent}", face="#FFFFFF")
        fig.text(content[0], content[1] + 0.018, value, fontsize=12, weight="bold", color="#203740")
    map_rect = add_card(fig, (0.055, 0.50, 0.43, 0.31), "Ubicación geográfica", accent="#2F6CC6")
    render_map(fig, map_rect, geojson, adm2_code)
    pyramid_rect = add_card(fig, (0.505, 0.50, 0.44, 0.31), "Pirámide de población 2022", accent="#2F6CC6")
    draw_pyramid(fig.add_axes(pyramid_rect, zorder=2), pyramid, "")
    household_rect = add_card(fig, (0.055, 0.12, 0.43, 0.35), "Hogares y viviendas", accent="#2F6CC6")
    text_rows(
        fig,
        household_rect,
        [
            ("Hogares totales", fmt_number(households.get("hogares_total"))),
            ("Población en hogares", fmt_number(households.get("poblacion_en_hogares"))),
            ("Personas por hogar", fmt_number(households.get("personas_por_hogar"), 2)),
            ("Viviendas totales", fmt_number(basic.get("viviendas_total"))),
            ("Viviendas ocupadas", fmt_number(basic.get("viviendas_ocupadas"))),
            ("Viviendas desocupadas", fmt_number(basic.get("viviendas_desocupadas"))),
        ],
    )
    size_rect = add_card(fig, (0.505, 0.12, 0.44, 0.35), "Tamaño de los hogares", accent="#2F6CC6")
    ax = fig.add_axes(size_rect, zorder=2)
    if household_size:
        labels = [str(row.get("miembros", "")) for row in household_size]
        values = [float(row.get("hogares", 0) or 0) for row in household_size]
        ax.bar(labels, values, color="#3B82F6")
        ax.grid(axis="y", color="#DCE3E5", linewidth=0.5)
        ax.tick_params(axis="both", labelsize=6)
        ax.set_ylabel("Hogares", fontsize=6)
        for side in ax.spines.values():
            side.set_visible(False)
    else:
        ax.text(0.5, 0.5, "Datos no disponibles", ha="center", va="center", color="#75868B")
        ax.axis("off")
    finish_dashboard_page(fig, pages[0], "Fuente: X Censo Nacional de Población y Vivienda 2022, ONE.")

    # Page 2: living conditions
    fig = new_dashboard_page("Diagnóstico municipal · Condición de vida y servicios básicos", name, "Hogares · Censo 2022")
    service_specs = [
        ("Servicios sanitarios", "servicios_sanitarios", COLORS["teal"]),
        ("Agua para uso doméstico", "agua_uso_domestico", COLORS["teal"]),
        ("Agua para beber", "agua_para_beber", COLORS["teal"]),
        ("Alumbrado principal", "alumbrado", COLORS["orange"]),
        ("Combustible para cocinar", "combustible_cocinar", COLORS["orange"]),
        ("Eliminación de basura", "eliminacion_basura", COLORS["orange"]),
    ]
    for index, (title, key, accent) in enumerate(service_specs):
        col = index % 3
        row = index // 3
        rect = (0.055 + col * 0.300, 0.535 - row * 0.405, 0.28, 0.375)
        content = add_card(fig, rect, title, accent=f"#{accent}", face="#FFFFFF")
        item = (living.get("servicios") or {}).get(key, {})
        text_rows(fig, content, top_categories(item, SERVICE_LABELS[key], 7), max_rows=7)
    finish_dashboard_page(fig, pages[1], "Fuente: X Censo Nacional de Población y Vivienda 2022, ONE.")

    # Page 3: education
    fig = new_dashboard_page("Diagnóstico municipal · Educación", name, "Censo 2022 y Anuario 2024")
    offer_rect = add_card(fig, (0.055, 0.62, 0.40, 0.29), "Oferta educativa asociada", accent="#C96B21", face="#FFF9F1")
    offer_rows = [("Centros totales", fmt_number(education_offer.get("centros_total")))]
    for level, label in (("inicial_primario", "Inicial / Primario"), ("secundario", "Secundario"), ("adultos", "Adultos")):
        row = (education_offer.get("niveles") or {}).get(level, {})
        offer_rows.append((label, f"{fmt_number(row.get('centros'))} centros · {fmt_number(row.get('matricula'))} estudiantes"))
    text_rows(fig, offer_rect, offer_rows, max_rows=5)
    level_rect = add_card(fig, (0.475, 0.62, 0.47, 0.29), "Nivel de instrucción", accent="#C96B21", face="#FFF9F1")
    ax = fig.add_axes((level_rect[0], level_rect[1], level_rect[2] * 0.45, level_rect[3]), zorder=2)
    levels = education_level.get("nivel") or education_level.get("niveles") or {}
    level_order = ["ninguno", "preprimaria", "primaria", "primaria_basica", "secundaria", "secundaria_media", "superior", "universitaria_superior", "postgrado"]
    values = []
    labels = []
    seen = set()
    label_map = {
        "ninguno": "Ninguno",
        "preprimaria": "Preprimaria",
        "primaria": "Primaria",
        "primaria_basica": "Primaria",
        "secundaria": "Secundaria",
        "secundaria_media": "Secundaria",
        "superior": "Superior",
        "universitaria_superior": "Superior",
        "postgrado": "Postgrado",
    }
    for key in level_order:
        label = label_map[key]
        if key in levels and label not in seen:
            seen.add(label)
            labels.append(label)
            values.append(float((levels[key] or {}).get("total", 0) or 0))
    if sum(values):
        ax.pie(values, colors=["#EF4444", "#F59E0B", "#22C55E", "#38BDF8", "#8B5CF6"][: len(values)], startangle=90)
        ax.axis("equal")
    else:
        ax.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        ax.axis("off")
    total_levels = sum(values)
    text_rows(
        fig,
        (level_rect[0] + level_rect[2] * 0.50, level_rect[1], level_rect[2] * 0.50, level_rect[3]),
        [(label, fmt_pct(pct(value, total_levels))) for label, value in zip(labels, values)],
        max_rows=6,
    )
    infra_rect = add_card(fig, (0.055, 0.33, 0.89, 0.25), "Infraestructura educativa del distrito asociado", accent="#C96B21", face="#FFF9F1")
    infra = ((education.get("anuario") or {}).get("infraestructura") or {})
    infra_rows = [
        ("Aulas por plantel", fmt_number(infra.get("aulas_por_plantel"), 1)),
        ("Secciones por centro", fmt_number(infra.get("secciones_por_centro"), 1)),
        ("Alumnos por aula", fmt_number(infra.get("alumnos_por_aula"), 1)),
        ("Alumnos por sección", fmt_number(infra.get("alumnos_por_seccion"), 1)),
        ("Alumnos por docente", fmt_number(infra.get("alumnos_por_docente"), 1)),
        ("Docentes por centro", fmt_number(infra.get("docentes_por_centro"), 1)),
    ]
    for index, (label, value) in enumerate(infra_rows):
        col = index % 3
        row = index // 3
        x = infra_rect[0] + col * infra_rect[2] / 3
        y = infra_rect[1] + infra_rect[3] - (row + 1) * infra_rect[3] / 2
        fig.text(x, y + 0.042, label, fontsize=6.7, color="#6C5A4C")
        fig.text(x, y + 0.015, value, fontsize=11, weight="bold", color="#203740")
    eff_rect = add_card(fig, (0.055, 0.08, 0.89, 0.21), "Eficiencia del sistema educativo", accent="#C96B21", face="#FFF9F1")
    efficiency = ((education.get("anuario") or {}).get("eficiencia") or {})
    for index, level in enumerate(("inicial", "primario", "secundario")):
        values_eff = efficiency.get(level) or {}
        x = eff_rect[0] + index * eff_rect[2] / 3
        fig.text(x, eff_rect[1] + eff_rect[3] - 0.030, level.capitalize(), fontsize=7.3, weight="bold", color="#C96B21")
        text_rows(
            fig,
            (x, eff_rect[1], eff_rect[2] / 3 - 0.025, eff_rect[3] - 0.045),
            [
                ("Promoción", fmt_pct(values_eff.get("promocion"))),
                ("Abandono", fmt_pct(values_eff.get("abandono"))),
                ("Reprobación", fmt_pct(values_eff.get("reprobacion"))),
            ],
            max_rows=3,
        )
    finish_dashboard_page(fig, pages[2], "Fuente: MINERD, Anuario Estadístico Educativo 2024 y Censo 2022.")

    # Page 4: economy and employment
    fig = new_dashboard_page("Diagnóstico municipal · Economía y empleo", name, "DEE 2024")
    dee = economy.get("dee_2024") or {}
    econ_cards = [
        ("Establecimientos", fmt_number(dee.get("total_establishments"))),
        ("Empleo estimado", fmt_number(dee.get("total_employees"), 1)),
        ("Promedio por establecimiento", fmt_number(dee.get("avg_employees_per_establishment"), 1)),
        (
            "Establecimientos / 1,000 hab.",
            fmt_number(
                (float(dee.get("total_establishments")) * 1000 / float(basic.get("poblacion_total")))
                if dee.get("total_establishments") is not None and basic.get("poblacion_total")
                else None,
                1,
            ),
        ),
    ]
    for index, (label, value) in enumerate(econ_cards):
        x = 0.055 + index * 0.225
        content = add_card(fig, (x, 0.82, 0.205, 0.09), label, accent="#C9526A", face="#FFF5F7")
        fig.text(content[0], content[1] + 0.018, value, fontsize=11.5, weight="bold", color="#7D3241")
    size_rect = add_card(fig, (0.055, 0.48, 0.41, 0.30), "Empleo por tamaño de establecimiento", accent="#C9526A", face="#FFF5F7")
    bands = dee.get("employment_size_bands") or []
    ax = fig.add_axes((size_rect[0], size_rect[1], size_rect[2] * 0.48, size_rect[3]), zorder=2)
    band_values = [float(row.get("employees", 0) or 0) for row in bands]
    if sum(band_values):
        ax.pie(band_values, colors=["#0EA5E9", "#22C55E", "#F97316", "#8B5CF6"], startangle=90)
        ax.axis("equal")
    else:
        ax.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        ax.axis("off")
    text_rows(
        fig,
        (size_rect[0] + size_rect[2] * 0.51, size_rect[1], size_rect[2] * 0.49, size_rect[3]),
        [(row.get("label", ""), fmt_pct(float(row.get("employees_share", 0) or 0) * 100)) for row in bands],
        max_rows=5,
    )
    sector_rect = add_card(fig, (0.49, 0.48, 0.455, 0.30), "Principales secciones CIIU por empleo", accent="#C9526A", face="#FFF5F7")
    sectors = dee.get("sectors") or []
    sector_rows = [
        (
            textwrap.shorten(str(row.get("label", "")), width=43, placeholder="…"),
            f"{fmt_number(row.get('employees'), 1)} · LQ {fmt_number(row.get('lq'), 2)}",
        )
        for row in sectors[:6]
    ]
    text_rows(fig, sector_rect, sector_rows, max_rows=6)
    bar_rect = add_card(fig, (0.055, 0.10, 0.89, 0.34), "Número de establecimientos por tamaño", accent="#C9526A", face="#FFF5F7")
    ax = fig.add_axes(bar_rect, zorder=2)
    if bands:
        labels = [row.get("label", "") for row in bands]
        establishments = [float(row.get("establishments", 0) or 0) for row in bands]
        ax.bar(range(len(labels)), establishments, color=["#0EA5E9", "#22C55E", "#F97316", "#8B5CF6"])
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, fontsize=6.5)
        ax.tick_params(axis="y", labelsize=6)
        ax.grid(axis="y", color="#E5DDE0", linewidth=0.5)
        for side in ax.spines.values():
            side.set_visible(False)
    else:
        ax.text(0.5, 0.5, "Datos no disponibles", ha="center", va="center", color="#75868B")
        ax.axis("off")
    finish_dashboard_page(fig, pages[3], "Fuente: Directorio de Empresas y Establecimientos (DEE) 2024.")

    # Page 5: health and comparison
    fig = new_dashboard_page("Diagnóstico municipal · Salud y comparación", name, "Registros disponibles")
    centers = health.get("centros") or []
    by_type = Counter((row.get("tipo_centro") or "Sin clasificación").title() for row in centers)
    health_rect = add_card(fig, (0.055, 0.57, 0.89, 0.34), f"Establecimientos de salud registrados · Total {len(centers)}", accent="#A27719", face="#FFFBEF")
    ax = fig.add_axes(health_rect, zorder=2)
    if by_type:
        items = by_type.most_common(9)
        labels = [textwrap.shorten(label, width=35, placeholder="…") for label, _ in items]
        values = [value for _, value in items]
        ax.barh(range(len(labels)), values, color="#D29A31")
        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels, fontsize=6.2)
        ax.invert_yaxis()
        ax.tick_params(axis="x", labelsize=6)
        ax.grid(axis="x", color="#E8E1CF", linewidth=0.5)
        for side in ax.spines.values():
            side.set_visible(False)
    else:
        ax.text(0.5, 0.5, "No se dispone de registros municipales en el Dashboard", ha="center", va="center", color="#75868B")
        ax.axis("off")
    compare_rect = add_card(fig, (0.055, 0.10, 0.89, 0.42), "Resumen comparativo: municipio y República Dominicana", accent="#2F6CC6")
    internet_local = ((tic.get("internet") or {}).get("rate_used"))
    computer_local = ((tic.get("computer") or {}).get("rate_used"))
    national_tic = national.get("tic") or {}
    compare = [
        ("Población 2022", fmt_number(basic.get("poblacion_total")), fmt_number((national.get("basic") or {}).get("poblacion_total"))),
        ("Personas por hogar", fmt_number(households.get("personas_por_hogar"), 2), fmt_number((national.get("hogares") or {}).get("personas_por_hogar"), 2)),
        (
            "Agua de acueducto dentro",
            fmt_pct(service_share(living, "agua_uso_domestico", "del_acueducto_dentro_de_la_vivienda")),
            fmt_pct(national_service_share(national.get("living") or {}, "agua_uso_domestico", "del_acueducto_dentro_de_la_vivienda")),
        ),
        (
            "Hogares con inodoro",
            fmt_pct(service_share(living, "servicios_sanitarios", "inodoro")),
            fmt_pct(national_service_share(national.get("living") or {}, "servicios_sanitarios", "inodoro")),
        ),
        (
            "Recogida municipal de basura",
            fmt_pct(service_share(living, "eliminacion_basura", "la_recoge_el_ayuntamiento")),
            fmt_pct(national_service_share(national.get("living") or {}, "eliminacion_basura", "la_recoge_el_ayuntamiento")),
        ),
        (
            "Uso de internet",
            fmt_pct(float(internet_local) * 100 if internet_local is not None else None),
            fmt_pct(float((national_tic.get("internet") or {}).get("rate_used")) * 100 if (national_tic.get("internet") or {}).get("rate_used") is not None else None),
        ),
        (
            "Uso de computadora",
            fmt_pct(float(computer_local) * 100 if computer_local is not None else None),
            fmt_pct(float((national_tic.get("computer") or {}).get("rate_used")) * 100 if (national_tic.get("computer") or {}).get("rate_used") is not None else None),
        ),
        (
            "Centros de salud / 10,000 hab.",
            fmt_number(len(centers) * 10000 / basic.get("poblacion_total") if basic.get("poblacion_total") else None, 2),
            fmt_number(
                (national.get("health") or {}).get("total_centros", 0) * 10000
                / (national.get("basic") or {}).get("poblacion_total", 1),
                2,
            ),
        ),
    ]
    x0, y0, width, height = compare_rect
    fig.text(x0, y0 + height - 0.028, "Indicador", fontsize=7.2, weight="bold", color="#203740")
    fig.text(x0 + width * 0.72, y0 + height - 0.028, name, fontsize=7.2, weight="bold", color="#2F6CC6", ha="right")
    fig.text(x0 + width, y0 + height - 0.028, "País", fontsize=7.2, weight="bold", color="#5E7076", ha="right")
    line_height = (height - 0.050) / len(compare)
    for index, (label, local_value, national_value) in enumerate(compare):
        yy = y0 + height - 0.055 - index * line_height
        fig.text(x0, yy, label, fontsize=6.7, color="#203740")
        fig.text(x0 + width * 0.72, yy, local_value, fontsize=6.7, color="#2F6CC6", ha="right", weight="bold")
        fig.text(x0 + width, yy, national_value, fontsize=6.7, color="#5E7076", ha="right")
    finish_dashboard_page(fig, pages[4], dashboard_source_note())
    return pages


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_geometry(table, widths_dxa: list[int]) -> None:
    total = sum(widths_dxa)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "120")
    tbl_ind.set(qn("w:type"), "dxa")
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths_dxa[index] / 1440)
            tc_w = cell._tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                cell._tc.get_or_add_tcPr().append(tc_w)
            tc_w.set(qn("w:w"), str(widths_dxa[index]))
            tc_w.set(qn("w:type"), "dxa")
            set_cell_margins(cell)


def set_table_borders(table, color: str = "D7E0E1", size: int = 5) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), str(size))
        node.set(qn("w:color"), color)


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_run(paragraph, text: str, *, bold=False, italic=False, color: str | None = None, size: float | None = None):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if size:
        run.font.size = Pt(size)
    run.font.name = "Calibri"
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Calibri")
    return run


def add_source_note(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="Source Note")
    add_run(paragraph, f"Fuente: {text}", color=COLORS["muted"], size=8.5)


def add_callout(doc: Document, title: str, body: str, *, warning: bool = False) -> None:
    table = doc.add_table(rows=1, cols=1)
    set_table_geometry(table, [9360])
    set_table_borders(table, "E3D2B7" if warning else "C9DCD7", 5)
    cell = table.cell(0, 0)
    set_cell_shading(cell, COLORS["warn_soft"] if warning else COLORS["soft"])
    set_cell_margins(cell, top=150, start=180, bottom=150, end=180)
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(3)
    add_run(
        paragraph,
        title,
        bold=True,
        color=COLORS["warn"] if warning else COLORS["teal"],
        size=10.5,
    )
    paragraph = cell.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(0)
    add_run(paragraph, body, color=COLORS["ink"], size=9.5)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)


def add_data_table(
    doc: Document,
    headers: list[str],
    rows: Iterable[Iterable[Any]],
    widths_dxa: list[int],
) -> None:
    values = [list(row) for row in rows]
    if sum(widths_dxa) != 9360:
        raise ValueError(f"Table widths must total 9360 DXA: {widths_dxa}")
    table = doc.add_table(rows=1, cols=len(headers))
    set_table_geometry(table, widths_dxa)
    set_table_borders(table)
    header = table.rows[0]
    set_repeat_table_header(header)
    for index, value in enumerate(headers):
        cell = header.cells[index]
        set_cell_shading(cell, COLORS["ink"])
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.space_after = Pt(0)
        add_run(paragraph, str(value), bold=True, color=COLORS["white"], size=8.7)
    for row_index, row_values in enumerate(values):
        row = table.add_row()
        for index, value in enumerate(row_values):
            cell = row.cells[index]
            if row_index % 2:
                set_cell_shading(cell, "F7F9F9")
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_after = Pt(0)
            add_run(paragraph, str(value), color=COLORS["ink"], size=8.7)
    set_table_geometry(table, widths_dxa)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)


def add_page_break(doc: Document) -> None:
    paragraph = doc.add_paragraph()
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def configure_document(doc: Document, municipality: str, province: str, code: str) -> None:
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor.from_string(COLORS["ink"])
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.333

    style_tokens = {
        "Heading 1": (16, "2E74B5", 18, 10),
        "Heading 2": (13, "2E74B5", 12, 6),
        "Heading 3": (12, "1F4D78", 8, 4),
    }
    for name, (size, color, before, after) in style_tokens.items():
        style = styles[name]
        style.font.name = "Calibri"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    for name in ("List Bullet", "List Number"):
        style = styles[name]
        style.font.name = "Calibri"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        style.font.size = Pt(11)
        style.paragraph_format.left_indent = Inches(0.375)
        style.paragraph_format.first_line_indent = Inches(-0.194)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.line_spacing = 1.208

    if "Source Note" not in styles:
        style = styles.add_style("Source Note", 1)
    else:
        style = styles["Source Note"]
    style.font.name = "Calibri"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    style.font.size = Pt(8.5)
    style.font.color.rgb = RGBColor.from_string(COLORS["muted"])
    style.paragraph_format.space_before = Pt(4)
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.line_spacing = 1.0

    header = section.header
    header_table = header.add_table(rows=1, cols=2, width=Inches(6.5))
    set_table_geometry(header_table, [6000, 3360])
    for cell in header_table.row_cells(0):
        set_cell_margins(cell, top=0, start=0, bottom=0, end=0)
    left = header_table.cell(0, 0).paragraphs[0]
    add_run(left, "PLAN MUNICIPAL DE DESARROLLO · BORRADOR TÉCNICO", bold=True, color=COLORS["teal"], size=7.5)
    right = header_table.cell(0, 1).paragraphs[0]
    right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(right, f"{municipality} · {code or 'código por confirmar'}", color=COLORS["muted"], size=7.5)

    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(paragraph, f"{municipality}, {province} · PMD {PERIOD} · ", color=COLORS["muted"], size=7.5)
    add_run(paragraph, "Borrador no aprobado", bold=True, color=COLORS["warn"], size=7.5)
    add_run(paragraph, " · ", color=COLORS["muted"], size=7.5)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    paragraph._p.append(field)


def age_structure(pyramid: list[dict[str, Any]]) -> dict[str, float | None]:
    if not pyramid:
        return {"youth": None, "working": None, "older": None, "dependency": None}
    youth = working = older = total = 0
    for row in pyramid:
        label = str(row.get("age_group", "")).strip().lower()
        value = float(row.get("male", 0) or 0) + float(row.get("female", 0) or 0)
        if "declarado" in label:
            continue
        total += value
        first = 0 if "menos" in label else int(re.search(r"\d+", label).group()) if re.search(r"\d+", label) else 0
        if first <= 14:
            youth += value
        elif first >= 65:
            older += value
        else:
            working += value
    dependency = (youth + older) * 100 / working if working else None
    return {
        "youth": pct(youth, total),
        "working": pct(working, total),
        "older": pct(older, total),
        "dependency": dependency,
    }


def paragraph(doc: Document, text: str) -> None:
    if text:
        doc.add_paragraph(text)


def write_general_information(
    doc: Document,
    municipality: dict[str, Any],
    code: str,
    data: dict[str, Any],
    wikipedia: dict[str, Any],
    historical: list[dict[str, Any]],
    historical_context: dict[str, Any],
    area_km2: float | None,
) -> None:
    name = municipality["municipio"]
    province = municipality["provincia"]
    region = municipality["region"]
    basic = data.get("basic") or {}
    households = data.get("households") or {}
    urban = data.get("urban_rural") or {}
    population = basic.get("poblacion_total")
    density = float(population) / float(area_km2) if population and area_km2 else None

    doc.add_heading("1. Información general", level=1)
    paragraph(
        doc,
        f"{name} forma parte de la provincia {province} y se integra en la región de planificación {region}. "
        f"El código geográfico utilizado en este borrador es {code or 'pendiente de confirmación'}. "
        "Esta identificación territorial sirve como clave de enlace entre el Portal de Planificación Municipal, "
        "el Dashboard de Diagnóstico Territorial y los documentos de planificación archivados. [GEO-01]",
    )
    add_data_table(
        doc,
        ["Dato territorial", "Valor", "Fuente / año"],
        [
            ["Municipio", name, "Portal municipal · 2026"],
            ["Provincia", province, "Portal municipal · 2026"],
            ["Región", region, "Clasificador geográfico"],
            ["Código geográfico", code or "Por confirmar", "Clasificador DIGEPRES"],
            ["Superficie cartográfica", f"{fmt_number(area_km2, 1)} km²" if area_km2 else "No disponible", "GeoJSON municipal"],
            ["Densidad estimada", f"{fmt_number(density, 1)} hab./km²" if density else "No disponible", "Censo 2022 + GeoJSON"],
        ],
        [2600, 3000, 3760],
    )

    doc.add_heading("1.1 Origen y formación del municipio", level=2)
    history_written = False
    if historical_context.get("history"):
        history_sentences = sentences(historical_context["history"])[:5]
        if history_sentences:
            paragraph(
                doc,
                f"El PMD anterior del período {historical_context.get('history_period', 'no identificado')} "
                "recoge los siguientes antecedentes sobre la formación del municipio. "
                + " ".join(history_sentences)
                + f" [PMD-ANT, págs. {historical_context.get('history_pages')}]",
            )
            history_written = True
    wiki_sections = wikipedia.get("sections") or split_wikipedia_sections(wikipedia.get("full_extract", ""))
    wiki_history = ""
    if is_municipal_wikipedia(wikipedia) and not history_written:
        preferred_keys = ("historia", "origen", "formacion", "etimologia", "toponimia")
        wiki_history = next(
            (
                value
                for key, value in wiki_sections.items()
                if any(token in base.clean(key) for token in preferred_keys)
            ),
            "",
        )
        if not wiki_history:
            wiki_history = wiki_sections.get("Introducción", "")
    if wiki_history and not history_written:
        useful = sentences(wiki_history)[:6]
        if useful:
            paragraph(
                doc,
                "Como referencia secundaria, la ficha municipal de Wikipedia resume la identidad y evolución "
                "histórica local de la siguiente manera: "
                + " ".join(useful)
                + " [WIKI-01].",
            )
            history_written = True
    if not history_written:
        paragraph(
            doc,
            "No se incorporó una reseña histórica detallada porque las fuentes disponibles no permitieron "
            "identificar un texto verificable. La OMPP debe completar este apartado con la norma de creación, "
            "el archivo municipal o una publicación histórica institucional.",
        )
    add_callout(
        doc,
        "Revisión final requerida",
        "Confirmar nombres propios, fechas, cambios de categoría territorial y norma de creación. "
        "Wikipedia se usa solo como fuente secundaria; cuando existe PMD anterior, sus páginas se indican expresamente.",
        warning=True,
    )

    doc.add_heading("1.2 Perfil demográfico y del poblamiento", level=2)
    if population:
        male = basic.get("poblacion_hombres")
        female = basic.get("poblacion_mujeres")
        variation = basic.get("variacion_pct")
        direction = "creció" if (variation or 0) >= 0 else "disminuyó"
        paragraph(
            doc,
            f"En 2022, {name} registró {fmt_number(population)} habitantes: {fmt_number(male)} hombres "
            f"({fmt_pct(pct(male, population))}) y {fmt_number(female)} mujeres "
            f"({fmt_pct(pct(female, population))}). Frente a los {fmt_number(basic.get('poblacion_2010'))} "
            f"habitantes censados en 2010, la población {direction} {fmt_pct(abs(float(variation)) if variation is not None else None)}. "
            "El cambio intercensal constituye la línea base para dimensionar servicios, infraestructura y demanda de suelo; "
            "no permite atribuir por sí solo causas de migración, fecundidad o movilidad residencial. [DASH-BASE]",
        )
        urban_total = urban.get("urbana")
        rural_total = urban.get("rural")
        paragraph(
            doc,
            f"La distribución territorial muestra {fmt_number(urban_total)} habitantes en zona urbana "
            f"({fmt_pct(pct(urban_total, urban.get('poblacion_total')))}) y {fmt_number(rural_total)} en zona rural "
            f"({fmt_pct(pct(rural_total, urban.get('poblacion_total')))}). Esta relación debe considerarse al organizar "
            "rutas de servicios, mantenimiento vial, equipamientos y mecanismos de atención a comunidades dispersas. [DASH-URB]",
        )
        paragraph(
            doc,
            f"El municipio cuenta con {fmt_number(households.get('hogares_total'))} hogares y "
            f"{fmt_number(households.get('poblacion_en_hogares'))} personas residentes en hogares. "
            f"El promedio es de {fmt_number(households.get('personas_por_hogar'), 2)} personas por hogar. "
            f"El Censo registra {fmt_number(basic.get('viviendas_total'))} viviendas, de las cuales "
            f"{fmt_number(basic.get('viviendas_ocupadas'))} estaban ocupadas y "
            f"{fmt_number(basic.get('viviendas_desocupadas'))} desocupadas. [DASH-HOG] [DASH-BASE]",
        )
    else:
        paragraph(
            doc,
            "El Dashboard aún no dispone de una ficha estadística municipal separada para este territorio. "
            "Por ello no se trasladaron cifras del municipio de origen ni de la provincia. La OMPP debe incorporar "
            "una línea base oficial que corresponda exactamente al nuevo ámbito municipal.",
        )

    doc.add_heading("1.3 Antecedentes de planificación", level=2)
    if historical:
        add_data_table(
            doc,
            ["PMD localizado", "Período", "Condición", "Uso permitido"],
            [
                [
                    f"Plan Municipal de Desarrollo de {name}",
                    f"{row.get('period_start')}-{row.get('period_end')}",
                    "Oficial" if row.get("document_kind") == "official" else "Borrador",
                    "Antecedente histórico; verificar continuidad",
                ]
                for row in sorted(historical, key=lambda item: item.get("period_end") or 0, reverse=True)
            ],
            [3000, 1400, 1500, 3460],
        )
        paragraph(
            doc,
            "Los PMD anteriores permiten reconstruir la trayectoria de problemas, activos, objetivos y proyectos. "
            "En este borrador se utilizan para contextualizar la historia local y señalar temas de continuidad. "
            "No se presume que una obra esté ejecutada, que un problema permanezca vigente ni que una prioridad anterior "
            "haya sido ratificada para 2025–2028 sin verificación de la OMPP. [PMD-ANT]",
        )
    else:
        paragraph(
            doc,
            "No se localizó un PMD anterior en el inventario descargado. Esta ausencia en el repositorio "
            "no demuestra que el documento no exista; la OMPP debe verificar archivo institucional, portal de transparencia "
            "y documentación del Concejo de Regidores.",
        )


def write_diagnostic_narrative(
    doc: Document,
    municipality: dict[str, Any],
    data: dict[str, Any],
    national: dict[str, Any],
    historical_context: dict[str, Any],
) -> list[dict[str, str]]:
    name = municipality["municipio"]
    basic = data.get("basic") or {}
    households = data.get("households") or {}
    urban = data.get("urban_rural") or {}
    living = data.get("living") or {}
    education = data.get("education") or {}
    education_level = data.get("education_level") or {}
    education_offer = data.get("education_offer") or {}
    economy = data.get("economy") or {}
    tic = data.get("tic") or {}
    health = data.get("health") or {}
    pyramid = data.get("pyramid") or []
    findings: list[dict[str, str]] = []

    doc.add_heading("2.1 Dinámica demográfica y hogares", level=2)
    if basic:
        age = age_structure(pyramid)
        paragraph(
            doc,
            f"La estructura por edad de {name} se distribuye aproximadamente en "
            f"{fmt_pct(age['youth'])} de población menor de 15 años, {fmt_pct(age['working'])} entre 15 y 64 años "
            f"y {fmt_pct(age['older'])} de 65 años o más. La relación de dependencia demográfica estimada es de "
            f"{fmt_number(age['dependency'], 1)} personas potencialmente dependientes por cada 100 en edad de trabajar. "
            "Estos valores orientan la demanda relativa de cuidados, espacios públicos, movilidad, formación y servicios sociales, "
            "pero deben analizarse con la distribución territorial y la capacidad real de atención. [DASH-PYR]",
        )
        urban_pct = pct(urban.get("urbana"), urban.get("poblacion_total"))
        paragraph(
            doc,
            f"El grado de urbanización es {fmt_pct(urban_pct)} y el tamaño medio de los hogares es "
            f"{fmt_number(households.get('personas_por_hogar'), 2)} personas. "
            "Una mayor concentración urbana puede facilitar cobertura de rutas y equipamientos, mientras que una proporción rural "
            "relevante exige criterios de accesibilidad y costos diferenciados. La cifra no sustituye un análisis por barrio, sección o paraje. [DASH-URB] [DASH-HOG]",
        )
        findings.append(
            {
                "theme": "Población y territorio",
                "finding": f"Variación intercensal de {fmt_pct(basic.get('variacion_pct'))}; urbanización de {fmt_pct(urban_pct)}.",
                "implication": "Ajustar capacidad y distribución territorial de servicios.",
                "source": "DASH-BASE / DASH-URB",
            }
        )
    else:
        paragraph(
            doc,
            "No existe una serie municipal separada en el Dashboard. El diagnóstico demográfico queda pendiente de una "
            "desagregación oficial específica; no se usaron promedios del territorio de origen como sustituto.",
        )

    doc.add_heading("2.2 Vivienda, agua, saneamiento y servicios básicos", level=2)
    if living:
        water = service_share(living, "agua_uso_domestico", "del_acueducto_dentro_de_la_vivienda")
        national_water = national_service_share(national.get("living") or {}, "agua_uso_domestico", "del_acueducto_dentro_de_la_vivienda")
        toilet = service_share(living, "servicios_sanitarios", "inodoro")
        no_toilet = service_share(living, "servicios_sanitarios", "no_tiene")
        bottled = service_share(living, "agua_para_beber", "botellones")
        paragraph(
            doc,
            f"El {fmt_pct(water)} de los hogares declara recibir agua del acueducto dentro de la vivienda, "
            f"frente a {fmt_pct(national_water)} a escala nacional. Para agua de beber, el uso de botellones representa "
            f"{fmt_pct(bottled)}. La comparación sugiere una condición de acceso físico y un patrón de consumo, pero no informa "
            "continuidad, presión, potabilidad ni costo. Esos cuatro aspectos deben verificarse antes de formular intervenciones. [DASH-VIDA]",
        )
        paragraph(
            doc,
            f"En saneamiento, {fmt_pct(toilet)} de los hogares reporta inodoro y {fmt_pct(no_toilet)} declara no disponer "
            "de servicio sanitario. La línea base permite dimensionar la brecha general, aunque no identifica sectores ni calidad "
            "de disposición final. El levantamiento municipal debe localizar hogares críticos y distinguir competencias del ayuntamiento, "
            "INAPA, salud pública y otros actores. [DASH-VIDA]",
        )
        lighting = (living.get("servicios") or {}).get("alumbrado", {})
        grid_share = pct(
            (lighting.get("categorias") or {}).get("energia_eletrica_del_tendido_publico"),
            lighting.get("total"),
        )
        waste = service_share(living, "eliminacion_basura", "la_recoge_el_ayuntamiento")
        burn = service_share(living, "eliminacion_basura", "la_queman")
        propane = service_share(living, "combustible_cocinar", "gas_propano")
        paragraph(
            doc,
            f"La red pública es la fuente principal de alumbrado para {fmt_pct(grid_share)} de los hogares. "
            f"La recogida de residuos por el ayuntamiento alcanza {fmt_pct(waste)}, mientras {fmt_pct(burn)} declara quemarlos. "
            f"El gas propano es el combustible principal de cocina en {fmt_pct(propane)} de los hogares. "
            "La cobertura de recogida no equivale a frecuencia, regularidad ni manejo ambientalmente adecuado; el plan operativo "
            "de residuos debe contrastar rutas, horarios, disposición final y sectores no servidos. [DASH-VIDA]",
        )
        findings.extend(
            [
                {
                    "theme": "Agua y saneamiento",
                    "finding": f"Agua de acueducto dentro: {fmt_pct(water)}; hogares sin sanitario: {fmt_pct(no_toilet)}.",
                    "implication": "Localizar brechas y validar continuidad/calidad.",
                    "source": "DASH-VIDA",
                },
                {
                    "theme": "Residuos sólidos",
                    "finding": f"Recogida municipal declarada: {fmt_pct(waste)}; quema: {fmt_pct(burn)}.",
                    "implication": "Revisar cobertura, frecuencia y disposición final.",
                    "source": "DASH-VIDA",
                },
            ]
        )
    else:
        paragraph(doc, "No se dispone de registros municipales separados de vivienda y servicios básicos en el Dashboard.")

    doc.add_heading("2.3 Educación y capacidades humanas", level=2)
    if education:
        offer = education_offer.get("niveles") or {}
        total_centers = education_offer.get("centros_total")
        total_students = sum(float((offer.get(level) or {}).get("matricula", 0) or 0) for level in offer)
        paragraph(
            doc,
            f"La oferta educativa asociada registra {fmt_number(total_centers)} centros y "
            f"{fmt_number(total_students)} matrículas agregadas. De ellas, "
            f"{fmt_number((offer.get('inicial_primario') or {}).get('matricula'))} corresponden a inicial/primario, "
            f"{fmt_number((offer.get('secundario') or {}).get('matricula'))} a secundario y "
            f"{fmt_number((offer.get('adultos') or {}).get('matricula'))} a educación de adultos. "
            "La cobertura corresponde a la asociación territorial usada por el Dashboard y debe verificarse cuando el distrito educativo "
            "abarque más de un municipio. [DASH-EDU]",
        )
        levels = education_level.get("nivel") or education_level.get("niveles") or {}
        total_levels = sum(float((row or {}).get("total", 0) or 0) for row in levels.values())
        none = float((levels.get("ninguno") or {}).get("total", 0) or 0)
        secondary_total = sum(
            float((levels.get(key) or {}).get("total", 0) or 0)
            for key in ("secundaria", "secundaria_media", "superior", "universitaria_superior", "postgrado")
        )
        paragraph(
            doc,
            f"Entre la población de referencia de tres años o más, {fmt_pct(pct(none, total_levels))} aparece sin nivel de instrucción "
            f"y {fmt_pct(pct(secondary_total, total_levels))} registra secundaria o un nivel superior. "
            "La distribución educativa ayuda a orientar alfabetización, formación técnica, inclusión digital y articulación con empleo, "
            "pero no debe interpretarse como calidad del aprendizaje. [DASH-EDU-NIVEL]",
        )
        efficiency = ((education.get("anuario") or {}).get("eficiencia") or {})
        primary = efficiency.get("primario") or {}
        secondary = efficiency.get("secundario") or {}
        paragraph(
            doc,
            f"Para el distrito educativo asociado, la promoción es {fmt_pct(primary.get('promocion'))} en primaria y "
            f"{fmt_pct(secondary.get('promocion'))} en secundaria; el abandono reportado en secundaria es "
            f"{fmt_pct(secondary.get('abandono'))}. La OMPP debe coordinar la lectura con MINERD y evitar atribuir al gobierno local "
            "resultados que corresponden al sistema educativo. [DASH-EDU]",
        )
        findings.append(
            {
                "theme": "Educación",
                "finding": f"Abandono secundario del distrito asociado: {fmt_pct(secondary.get('abandono'))}.",
                "implication": "Validar alcance territorial y coordinar permanencia escolar.",
                "source": "DASH-EDU",
            }
        )
    else:
        paragraph(doc, "El Dashboard no ofrece una ficha educativa municipal separada para este territorio.")

    doc.add_heading("2.4 Economía, establecimientos y empleo", level=2)
    dee = economy.get("dee_2024") or {}
    if dee:
        total_establishments = dee.get("total_establishments")
        total_employees = dee.get("total_employees")
        bands = dee.get("employment_size_bands") or []
        micro = next((row for row in bands if row.get("size_band") == "micro_1_10"), {})
        micro_est_share = pct(micro.get("establishments"), total_establishments)
        paragraph(
            doc,
            f"El Directorio de Empresas y Establecimientos 2024 identifica {fmt_number(total_establishments)} establecimientos "
            f"y un empleo estimado de {fmt_number(total_employees, 1)}, equivalente a "
            f"{fmt_number(dee.get('avg_employees_per_establishment'), 1)} empleos por establecimiento. "
            f"Los establecimientos de 1 a 10 personas representan {fmt_pct(micro_est_share)} del total. "
            "Esto caracteriza la estructura formal observada por el DEE; no incluye necesariamente toda la economía informal, "
            "el autoempleo ni la producción agropecuaria familiar. [DASH-ECO]",
        )
        sectors = dee.get("sectors") or []
        top_sectors = [
            f"{row.get('label')} ({fmt_number(row.get('employees'), 1)} empleos estimados)"
            for row in sectors[:3]
        ]
        top_specialization = dee.get("top_specialization") or {}
        paragraph(
            doc,
            "Las actividades con mayor empleo estimado en la fuente son "
            + "; ".join(top_sectors)
            + ". "
            + (
                f"La mayor especialización relativa se observa en {top_specialization.get('label')} "
                f"(cociente de localización {fmt_number(top_specialization.get('lq'), 2)}). "
                if top_specialization.get("label")
                else ""
            )
            + "El cociente de localización compara concentración relativa, no productividad ni competitividad. [DASH-ECO]",
        )
        findings.append(
            {
                "theme": "Economía local",
                "finding": f"{fmt_pct(micro_est_share)} de los establecimientos observados son micro (1–10).",
                "implication": "Diseñar apoyo empresarial acorde con la estructura local y confirmar informalidad.",
                "source": "DASH-ECO",
            }
        )
    else:
        paragraph(doc, "No se dispone de una ficha económica municipal separada en el Dashboard.")

    doc.add_heading("2.5 Salud, conectividad y acceso a información", level=2)
    centers = health.get("centros") or []
    internet = ((tic.get("internet") or {}).get("rate_used"))
    computer = ((tic.get("computer") or {}).get("rate_used"))
    if centers or tic:
        by_type = Counter((row.get("tipo_centro") or "Sin clasificación").title() for row in centers)
        type_summary = ", ".join(f"{label}: {count}" for label, count in by_type.most_common(5)) or "sin registros"
        paragraph(
            doc,
            f"El Dashboard registra {len(centers)} establecimientos de salud en el municipio. La composición principal es: "
            f"{type_summary}. Este conteo informa presencia física, no cartera de servicios, personal, horario, capacidad, calidad "
            "ni accesibilidad efectiva. La planificación municipal debe validar esos elementos con el Servicio Nacional de Salud. [DASH-SALUD]",
        )
        paragraph(
            doc,
            f"El uso de internet alcanza {fmt_pct(float(internet) * 100 if internet is not None else None)} y el de computadora "
            f"{fmt_pct(float(computer) * 100 if computer is not None else None)}. "
            "La diferencia entre conectividad y disponibilidad de equipo es relevante para trámites, educación, información pública "
            "y participación digital. El dato requiere desagregación por edad, sexo y zona para diseñar acciones focalizadas. [DASH-TIC]",
        )
        findings.append(
            {
                "theme": "Conectividad",
                "finding": f"Internet: {fmt_pct(float(internet) * 100 if internet is not None else None)}; computadora: {fmt_pct(float(computer) * 100 if computer is not None else None)}.",
                "implication": "Precisar brechas por grupos y territorio.",
                "source": "DASH-TIC",
            }
        )
    else:
        paragraph(doc, "No se dispone de registros municipales separados de salud y conectividad en el Dashboard.")

    doc.add_heading("2.6 Contexto territorial, ambiental y continuidad", level=2)
    if historical_context.get("history"):
        paragraph(
            doc,
            "El PMD anterior aporta antecedentes históricos y territoriales útiles para comprender la formación del municipio. "
            "Sin embargo, sus descripciones de infraestructura, ambiente, economía o gestión corresponden a otro período. "
            "Antes de reutilizarlas, la OMPP debe clasificarlas como vigentes, modificadas o superadas y registrar la evidencia actual. [PMD-ANT]",
        )
    else:
        paragraph(
            doc,
            "Las fuentes estadísticas utilizadas no describen con suficiente detalle uso del suelo, riesgos, hidrología, residuos, "
            "movilidad o estado de la infraestructura. Este apartado debe completarse con cartografía oficial, instrumentos ambientales "
            "y verificación técnica municipal; no se agregaron afirmaciones de prensa ni datos no comprobables.",
        )
    return findings


def build_foda(data: dict[str, Any], national: dict[str, Any], findings: list[dict[str, str]]) -> dict[str, list[str]]:
    living = data.get("living") or {}
    tic = data.get("tic") or {}
    economy = data.get("economy") or {}
    education = data.get("education") or {}
    strengths: list[str] = []
    weaknesses: list[str] = []
    opportunities = [
        "Uso de la línea base del Dashboard para seguimiento anual y priorización territorial.",
        "Coordinación con instituciones sectoriales cuando la competencia no sea exclusivamente municipal.",
    ]
    threats = [
        "Decisiones basadas en promedios municipales que oculten diferencias entre barrios, secciones y parajes.",
        "Desactualización de diagnósticos anteriores si no se verifican en campo y con registros administrativos.",
    ]

    metrics = [
        (
            "Acceso a agua del acueducto dentro de la vivienda",
            service_share(living, "agua_uso_domestico", "del_acueducto_dentro_de_la_vivienda"),
            national_service_share(national.get("living") or {}, "agua_uso_domestico", "del_acueducto_dentro_de_la_vivienda"),
        ),
        (
            "Recogida de residuos por el ayuntamiento",
            service_share(living, "eliminacion_basura", "la_recoge_el_ayuntamiento"),
            national_service_share(national.get("living") or {}, "eliminacion_basura", "la_recoge_el_ayuntamiento"),
        ),
    ]
    for label, local, country in metrics:
        if local is None or country is None:
            continue
        text = f"{label}: {fmt_pct(local)} (país: {fmt_pct(country)})."
        (strengths if local >= country else weaknesses).append(text)

    internet = ((tic.get("internet") or {}).get("rate_used"))
    national_internet = (((national.get("tic") or {}).get("internet") or {}).get("rate_used"))
    if internet is not None and national_internet is not None:
        text = f"Uso de internet: {fmt_pct(float(internet) * 100)} (país: {fmt_pct(float(national_internet) * 100)})."
        (strengths if internet >= national_internet else weaknesses).append(text)

    dee = economy.get("dee_2024") or {}
    if dee.get("total_establishments"):
        strengths.append(
            f"Base económica formal observable: {fmt_number(dee.get('total_establishments'))} establecimientos en el DEE 2024."
        )
    dropout = (((education.get("anuario") or {}).get("eficiencia") or {}).get("secundario") or {}).get("abandono")
    if dropout is not None and float(dropout) >= 3:
        weaknesses.append(f"Abandono secundario del distrito asociado: {fmt_pct(dropout)}.")
    for finding in findings:
        if len(weaknesses) >= 5:
            break
        candidate = finding.get("finding", "")
        if candidate and candidate not in weaknesses:
            weaknesses.append(candidate)
    if not strengths:
        strengths.append("Activos y capacidades locales pendientes de validación por la OMPP y el CDM.")
    if not weaknesses:
        weaknesses.append("Brechas internas pendientes de localizar y verificar.")
    return {
        "Fortalezas": strengths[:5],
        "Debilidades": weaknesses[:5],
        "Oportunidades": opportunities,
        "Amenazas": threats,
    }


def project_proposals(findings: list[dict[str, str]]) -> list[list[str]]:
    rows: list[list[str]] = []
    mapping = {
        "Agua y saneamiento": (
            "Actualización del diagnóstico operativo de agua y saneamiento",
            "Localizar continuidad, calidad, cobertura y hogares críticos; coordinar responsabilidades sectoriales.",
        ),
        "Residuos sólidos": (
            "Optimización de rutas y control del manejo de residuos",
            "Medir frecuencia, sectores no cubiertos, puntos críticos y disposición final.",
        ),
        "Educación": (
            "Mecanismo local de coordinación para permanencia educativa",
            "Validar datos con el distrito educativo y focalizar factores de abandono dentro de las competencias locales.",
        ),
        "Economía local": (
            "Programa de información y servicios para microempresas",
            "Caracterizar necesidades reales, formalidad, acceso a capacitación y articulación con instituciones competentes.",
        ),
        "Conectividad": (
            "Acceso y alfabetización digital municipal",
            "Identificar grupos con menor acceso y usar espacios municipales existentes antes de definir inversión.",
        ),
        "Población y territorio": (
            "Sistema municipal de información territorial",
            "Desagregar indicadores por sectores y mantener línea base anual con fuente y responsable.",
        ),
    }
    for finding in findings:
        if finding["theme"] not in mapping:
            continue
        title, scope = mapping[finding["theme"]]
        rows.append([title, scope, finding["source"], "Pendiente de validación"])
    if not rows:
        rows.append(
            [
                "Sistema municipal de información territorial",
                "Completar la línea base y localizar brechas antes de formular inversión.",
                "Dashboard / OMPP",
                "Pendiente de validación",
            ]
        )
    return rows[:6]


def add_dashboard_images(doc: Document, paths: list[Path], name: str) -> None:
    doc.add_heading("2. Diagnóstico municipal", level=1)
    paragraph(
        doc,
        "Las cinco láminas siguientes reproducen en Word la lógica de impresión del Dashboard de Diagnóstico Territorial. "
        "Presentan la línea base estadística antes de su interpretación narrativa. Los datos corresponden al territorio "
        "seleccionado y conservan el año y la fuente; una ausencia se muestra como «No disponible» y no se sustituye por estimaciones.",
    )
    add_data_table(
        doc,
        ["Lámina", "Contenido", "Lectura esperada"],
        [
            ["1", "Población, mapa, estructura por edad, hogares y viviendas", "Escala y composición demográfica"],
            ["2", "Agua, saneamiento, alumbrado, combustible y residuos", "Coberturas declaradas por hogares"],
            ["3", "Oferta, nivel de instrucción, infraestructura y eficiencia educativa", "Capacidades y alertas educativas"],
            ["4", "Establecimientos, empleo, tamaño empresarial y sectores CIIU", "Estructura económica formal observada"],
            ["5", "Establecimientos de salud y comparación nacional", "Presencia física y brechas relativas"],
        ],
        [1000, 4300, 4060],
    )
    add_callout(
        doc,
        "Cómo leer las láminas",
        "Cada lámina es una línea base, no una conclusión automática. La interpretación debe conservar tres preguntas: "
        "qué mide el indicador, qué no mide y qué comprobación territorial hace falta antes de decidir.",
    )
    for index, path in enumerate(paths, 1):
        add_page_break(doc)
        doc.add_picture(str(path), width=Inches(6.5))
        picture_paragraph = doc.paragraphs[-1]
        picture_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        picture_paragraph.paragraph_format.space_after = Pt(0)
        picture_paragraph.paragraph_format.keep_together = True
        if index == len(paths):
            add_source_note(
                doc,
                f"Dashboard de Diagnóstico Territorial, ficha de {name}. Láminas generadas el {TODAY.isoformat()}.",
            )


def build_document(
    municipality: dict[str, Any],
    code: str,
    adm2_code: str,
    data: dict[str, Any],
    national: dict[str, Any],
    wikipedia: dict[str, Any],
    historical: list[dict[str, Any]],
    historical_context: dict[str, Any],
    area_km2: float | None,
    dashboard_pages: list[Path],
    output_path: Path,
) -> None:
    name = municipality["municipio"]
    province = municipality["provincia"]
    region = municipality["region"]
    doc = Document()
    configure_document(doc, name, province, code)

    # Editorial cover
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(72)
    spacer.paragraph_format.space_after = Pt(0)
    kicker = doc.add_paragraph()
    kicker.alignment = WD_ALIGN_PARAGRAPH.CENTER
    kicker.paragraph_format.space_after = Pt(18)
    add_run(kicker, "REPÚBLICA DOMINICANA · PLANIFICACIÓN MUNICIPAL", bold=True, color=COLORS["gold"], size=10)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(10)
    add_run(title, "Plan Municipal de Desarrollo", bold=True, color=COLORS["ink"], size=29)
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(5)
    add_run(subtitle, name, bold=True, color=COLORS["teal"], size=20)
    province_p = doc.add_paragraph()
    province_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    province_p.paragraph_format.space_after = Pt(30)
    add_run(province_p, f"{province} · Región {region}", color=COLORS["muted"], size=12)
    status = doc.add_paragraph()
    status.alignment = WD_ALIGN_PARAGRAPH.CENTER
    status.paragraph_format.space_after = Pt(55)
    add_run(status, f"BORRADOR TÉCNICO {PERIOD}", bold=True, color=COLORS["blue"], size=13)
    cover_meta = doc.add_paragraph()
    cover_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(
        cover_meta,
        "Información General y Diagnóstico precompletados para revisión final\n"
        "FODA, Visión y Proyectos sujetos a validación de OMPP y CDM",
        italic=True,
        color=COLORS["muted"],
        size=10,
    )
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_p.paragraph_format.space_before = Pt(22)
    add_run(date_p, f"Versión técnica · {TODAY.strftime('%d-%m-%Y')}", color=COLORS["muted"], size=9)

    add_page_break(doc)
    doc.add_heading("Estado y alcance del borrador", level=1)
    add_callout(
        doc,
        "Qué está listo para revisión final",
        "La identificación territorial, la reseña histórica disponible, la línea base demográfica y el diagnóstico "
        "de servicios, educación, economía, salud y conectividad se presentan ya redactados. La OMPP debe comprobar "
        "cifras, nombres, alcance territorial y cualquier actualización posterior al año de cada fuente.",
    )
    add_callout(
        doc,
        "Qué no se considera aprobado",
        "El documento no acredita consulta ciudadana, reuniones, acuerdos del CDM, aprobación del Concejo de Regidores, "
        "presupuestos, beneficiarios ni compromisos institucionales. FODA, visión y cartera de proyectos son propuestas "
        "para discusión y deben modificarse o eliminarse si no son ratificadas.",
        warning=True,
    )
    add_data_table(
        doc,
        ["Campo", "Valor"],
        [
            ["Estado", "Borrador técnico no aprobado"],
            ["Municipio / provincia", f"{name} / {province}"],
            ["Región", region],
            ["Período", PERIOD],
            ["Código geográfico", code or "Por confirmar"],
            ["Dashboard", "Disponible" if adm2_code else "Sin ficha municipal separada"],
            ["Revisión mínima", "OMPP: datos y redacción · CDM: prioridades y estrategia"],
        ],
        [2600, 6760],
    )
    doc.add_heading("Contenido", level=2)
    contents = [
        "1. Información general: territorio, historia, demografía y antecedentes de planificación.",
        "2. Diagnóstico municipal: cinco láminas del Dashboard y análisis narrativo.",
        "3. Síntesis diagnóstica y FODA preliminar.",
        "4. Marco estratégico y proyectos para validación.",
        "5. Protocolo de revisión, fuentes y trazabilidad.",
    ]
    for item in contents:
        doc.add_paragraph(item, style="List Number")

    add_page_break(doc)
    write_general_information(
        doc,
        municipality,
        code,
        data,
        wikipedia,
        historical,
        historical_context,
        area_km2,
    )

    add_dashboard_images(doc, dashboard_pages, name)
    add_page_break(doc)
    doc.add_heading("Lectura narrativa del diagnóstico", level=1)
    paragraph(
        doc,
        "La lectura siguiente explica qué muestran los indicadores y cuáles son sus límites. "
        "Se diferencia entre dato observado, interpretación técnica y asunto que requiere comprobación municipal.",
    )
    findings = write_diagnostic_narrative(doc, municipality, data, national, historical_context)

    add_page_break(doc)
    doc.add_heading("3. Síntesis diagnóstica", level=1)
    paragraph(
        doc,
        "La síntesis organiza las principales señales cuantitativas para facilitar la revisión técnica. "
        "No constituye una priorización ciudadana ni sustituye la validación territorial.",
    )
    add_data_table(
        doc,
        ["Tema", "Hallazgo verificable", "Implicación para revisar", "Fuente"],
        [
            [row["theme"], row["finding"], row["implication"], row["source"]]
            for row in findings
        ]
        or [["Línea base", "Datos municipales por completar", "Completar fuente oficial", "OMPP"]],
        [1700, 3000, 3100, 1560],
    )

    doc.add_heading("3.1 FODA preliminar", level=2)
    foda = build_foda(data, national, findings)
    add_callout(
        doc,
        "Condición de uso",
        "Este FODA es una lectura técnica inicial. Solo debe pasar a la versión oficial después de que la OMPP "
        "verifique los datos y el CDM confirme fortalezas, debilidades, oportunidades y amenazas.",
        warning=True,
    )
    add_data_table(
        doc,
        ["Fortalezas observadas", "Debilidades observadas"],
        [
            [
                "\n".join(f"• {item}" for item in foda["Fortalezas"]),
                "\n".join(f"• {item}" for item in foda["Debilidades"]),
            ]
        ],
        [4680, 4680],
    )
    add_data_table(
        doc,
        ["Oportunidades por validar", "Amenazas por validar"],
        [
            [
                "\n".join(f"• {item}" for item in foda["Oportunidades"]),
                "\n".join(f"• {item}" for item in foda["Amenazas"]),
            ]
        ],
        [4680, 4680],
    )

    add_page_break(doc)
    doc.add_heading("4. Marco estratégico para validación", level=1)
    paragraph(
        doc,
        "La visión municipal no se completa automáticamente. Debe expresar una situación futura compartida y "
        "ser acordada por las instancias municipales correspondientes. Se propone la siguiente estructura de trabajo:",
    )
    add_data_table(
        doc,
        ["Componente", "Base propuesta", "Decisión requerida"],
        [
            ["Visión municipal", "Redactar en una frase el municipio deseado al cierre del período.", "CDM: acordar texto"],
            ["Eje 1 · Gestión y territorio", "Información, planificación, servicios y ordenamiento.", "OMPP/CDM: confirmar"],
            ["Eje 2 · Desarrollo social", "Servicios básicos, educación, salud, inclusión y convivencia.", "OMPP/CDM: confirmar"],
            ["Eje 3 · Economía local", "Empleo, microempresas, activos productivos y articulación sectorial.", "OMPP/CDM: confirmar"],
            ["Eje 4 · Ambiente y resiliencia", "Riesgos, recursos naturales, residuos y adaptación.", "OMPP/CDM: completar fuentes"],
        ],
        [2100, 4700, 2560],
    )
    doc.add_heading("4.1 Ideas iniciales de proyectos", level=2)
    paragraph(
        doc,
        "Las ideas siguientes derivan de brechas del diagnóstico. No incluyen costo, beneficiarios, cronograma, "
        "aprobación ni fuente de financiamiento.",
    )
    add_data_table(
        doc,
        ["Idea", "Alcance antes de formular", "Evidencia inicial", "Estado"],
        project_proposals(findings),
        [2500, 3900, 1460, 1500],
    )
    doc.add_heading("4.2 Ficha para completar por proyecto", level=2)
    add_data_table(
        doc,
        ["Campo", "Registro municipal"],
        [
            ["Problema y evidencia", ""],
            ["Objetivo, resultado y meta", ""],
            ["Competencia legal", ""],
            ["Responsable y aliados", ""],
            ["Indicador y línea base", ""],
            ["Costo y financiamiento", ""],
            ["Validación OMPP / CDM", ""],
        ],
        [2900, 6460],
    )

    add_page_break(doc)
    doc.add_heading("5. Revisión, fuentes y trazabilidad", level=1)
    doc.add_heading("5.1 Lista de control final", level=2)
    add_data_table(
        doc,
        ["Control", "Responsable", "Evidencia / corrección"],
        [
            ["Confirmar historia, norma de creación y límites", "OMPP", ""],
            ["Confirmar cifras, años, unidades y cobertura", "OMPP / áreas técnicas", ""],
            ["Localizar brechas por barrio, sección o paraje", "OMPP", ""],
            ["Revisar continuidad de PMD anteriores", "OMPP", ""],
            ["Validar síntesis y FODA", "CDM", ""],
            ["Acordar visión, objetivos y proyectos", "CDM / Ayuntamiento", ""],
            ["Documentar aprobación y publicación", "Autoridad competente", ""],
        ],
        [3900, 2400, 3060],
    )
    doc.add_heading("5.2 Fuentes utilizadas", level=2)
    source_rows = [
        ["GEO-01", "Clasificador geográfico y GeoJSON municipal", "2026", "Sin paginación", "Alta"],
        ["DASH-BASE", "X Censo Nacional de Población y Vivienda", "2022", "Dataset municipal", "Alta"],
        ["DASH-PYR", "X Censo: edad y sexo", "2022", "Dataset municipal", "Alta"],
        ["DASH-HOG", "X Censo: hogares y viviendas", "2022", "Dataset municipal", "Alta"],
        ["DASH-URB", "X Censo: población urbana y rural", "2022", "Dataset municipal", "Alta"],
        ["DASH-VIDA", "X Censo: condición de vida y servicios", "2022", "Dataset municipal", "Alta"],
        ["DASH-EDU", "Anuario Estadístico Educativo", "2024", "Distrito asociado", "Media"],
        ["DASH-EDU-NIVEL", "X Censo: nivel de instrucción", "2022", "Dataset municipal", "Alta"],
        ["DASH-ECO", "Directorio de Empresas y Establecimientos", "2024", "Dataset municipal", "Media"],
        ["DASH-SALUD", "Registro de establecimientos SNS mostrado por Dashboard", "s/f", "Dataset municipal", "Media"],
        ["DASH-TIC", "X Censo: tecnologías de información", "2022", "Dataset municipal", "Alta"],
    ]
    if is_municipal_wikipedia(wikipedia):
        source_rows.append(
            [
                "WIKI-01",
                f"Wikipedia en español: {wikipedia.get('title')}",
                wikipedia.get("retrieved_at", TODAY.isoformat()),
                "Sin paginación",
                "Baja / secundaria",
            ]
        )
    for index, row in enumerate(
        sorted(historical, key=lambda item: item.get("period_end") or 0, reverse=True), 1
    ):
        source_rows.append(
            [
                f"PMD-ANT-{index:02d}",
                f"PMD de {name} {row.get('period_start')}-{row.get('period_end')}",
                f"{row.get('period_start')}-{row.get('period_end')}",
                f"{row.get('pages') or 's/p'} páginas; citas específicas en texto",
                "Alta" if str(row.get("validation_status", "")).startswith("verified") else "Media",
            ]
        )
    add_data_table(doc, ["ID", "Documento", "Año", "Página / cobertura", "Confianza"], source_rows, [1300, 3300, 1300, 2200, 1260])
    add_source_note(
        doc,
        "Los datasets del Dashboard no tienen paginación. Cuando se usa texto de un PMD anterior, el rango de páginas "
        "se consigna en el párrafo correspondiente.",
    )
    doc.add_heading("5.3 Registro de correcciones", level=2)
    add_data_table(
        doc,
        ["Sección", "Corrección", "Fuente / página", "Responsable / fecha"],
        [["", "", "", ""] for _ in range(6)],
        [1700, 3300, 2100, 2260],
    )

    doc.core_properties.title = f"PMD de {name} - Borrador técnico sustantivo {PERIOD}"
    doc.core_properties.subject = "Información general y diagnóstico precompletados con trazabilidad"
    doc.core_properties.author = "DDPT - Generación asistida para revisión municipal"
    doc.core_properties.keywords = "PMD, municipio, diagnóstico, OMPP, CDM, trazabilidad, Dashboard"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def first_index(rows: Any) -> dict[str, Any]:
    if isinstance(rows, dict):
        return {base.code_key(key): value for key, value in rows.items()}
    result: dict[str, Any] = {}
    for row in rows or []:
        key = base.code_key(row.get("adm2_code"))
        if key and key not in result:
            result[key] = row
    return result


def build_data_bundle(
    adm2_code: str,
    datasets: dict[str, Any],
    adm2_map_2010: dict[str, str | None],
) -> dict[str, Any]:
    if not adm2_code:
        return {
            "basic": None,
            "households": None,
            "urban_rural": None,
            "living": None,
            "education": None,
            "education_level": None,
            "education_offer": None,
            "economy": None,
            "tic": None,
            "health": None,
            "pyramid": [],
            "pyramid2010": [],
            "household_size": [],
        }
    old_code = adm2_map_2010.get(adm2_code)
    return {
        "basic": datasets["basic"].get(adm2_code),
        "households": datasets["households"].get(adm2_code),
        "urban_rural": datasets["urban_rural"].get(adm2_code),
        "living": datasets["living"].get(adm2_code),
        "education": datasets["education"].get(adm2_code),
        "education_level": datasets["education_level"].get(adm2_code),
        "education_offer": datasets["education_offer"].get(adm2_code),
        "economy": datasets["economy"].get(adm2_code),
        "tic": datasets["tic"].get(adm2_code),
        "health": datasets["health"].get(adm2_code),
        "pyramid": (datasets["pyramid"].get(adm2_code) or {}).get("age_groups", []),
        "pyramid2010": datasets["age2010"].get(base.code_key(old_code), []) if old_code else [],
        "household_size": datasets["household_size"].get(adm2_code, []),
    }


def load_datasets(data_dir: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, str | None]]:
    datasets = {
        "basic": first_index(load_json(data_dir / "indicadores_basicos.json")),
        "households": first_index(load_json(data_dir / "hogares_resumen.json")),
        "urban_rural": first_index(load_json(data_dir / "poblacion_urbana_rural.json")),
        "living": first_index(load_json(data_dir / "condicion_vida.json")),
        "education": first_index(load_json(data_dir / "educacion.json")),
        "education_level": first_index(load_json(data_dir / "educacion_nivel.json")),
        "education_offer": first_index(load_json(data_dir / "educacion_oferta_municipal.json")),
        "economy": first_index(load_json(data_dir / "economia_empleo.json")),
        "tic": first_index(load_json(data_dir / "tic.json")),
        "health": first_index(load_json(data_dir / "salud_establecimientos.json")),
        "pyramid": first_index(load_json(data_dir / "pyramids.json")),
        "age2010": group_index(load_json(data_dir / "edad_sexo_2010.json")),
        "household_size": group_index(load_json(data_dir / "tamano_hogar.json")),
    }
    national = {
        "basic": load_json(data_dir / "national_basic.json"),
        "hogares": load_json(data_dir / "national_hogares.json"),
        "living": load_json(data_dir / "national_condicion_vida.json"),
        "education_level": load_json(data_dir / "national_educacion_nivel.json"),
        "education_offer": load_json(data_dir / "national_educacion_oferta.json"),
        "economy": load_json(data_dir / "national_economia_empleo.json"),
        "tic": load_json(data_dir / "national_tic.json"),
        "health": load_json(data_dir / "national_salud_establecimientos.json"),
    }
    adm2_map_2010 = {
        base.code_key(key): (str(value).zfill(6) if value else None)
        for key, value in load_json(data_dir / "adm2_map_2010.json").items()
    }
    return datasets, national, adm2_map_2010


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workspace", type=Path, required=True)
    parser.add_argument("--portal-repo", type=Path, required=True)
    parser.add_argument("--municipality", help="Generate one municipality as a prototype")
    parser.add_argument("--all", action="store_true", help="Generate all 104 targets")
    parser.add_argument("--web-copy", action="store_true", help="Copy DOCX files to the portal downloads folder")
    parser.add_argument("--skip-wikipedia-refresh", action="store_true")
    parser.add_argument("--resume", action="store_true", help="Keep an existing rich DOCX instead of rebuilding it")
    args = parser.parse_args()

    workspace = args.workspace.resolve()
    portal_repo = args.portal_repo.resolve()
    data_dir = workspace / "tmp" / "dashboard_repo_full_inspect" / "src" / "data"
    output_dir = workspace / "output" / "pmd_borradores_2025_2028"
    docx_dir = output_dir / "docx"
    assets_dir = output_dir / "assets"
    manifest_path = output_dir / "manifest.json"
    wikipedia_cache_path = output_dir / "wikipedia_cache.json"
    historical_cache_path = output_dir / "historical_context_cache.json"
    inventory_path = workspace / "output" / "pmd_historico_fuentes_no_sismap" / "inventory.json"
    classifier_path = workspace / "output" / "clasificador_geografico" / "clasificador_geografico_01-10.txt"
    portal_data_path = portal_repo / "app" / "data" / "municipios.json"
    web_dir = portal_repo / "public" / "downloads" / "pmd-borradores"
    geojson_path = portal_repo / "public" / "data" / "adm2.geojson"

    portal_data = load_json(portal_data_path)
    targets = [
        item
        for item in portal_data
        if not item.get("pmd", {}).get("hasOfficialEvidence")
        and not (item.get("pmd", {}).get("has7_12") or item.get("pmd", {}).get("hasDraft"))
    ]
    if len(targets) != 104:
        raise RuntimeError(f"Expected 104 target municipalities; found {len(targets)}")
    if args.municipality:
        wanted = base.clean(args.municipality)
        targets = [item for item in targets if base.clean(item["municipio"]) == wanted]
        if not targets:
            raise RuntimeError(f"Target municipality not found: {args.municipality}")
    elif not args.all:
        parser.error("Choose --municipality NAME or --all")

    dashboard_index = load_json(data_dir / "municipios_index.json")
    dashboard_by_territory = {
        base.territory_key(item["municipio"], item["provincia"]): base.code_key(item["adm2_code"])
        for item in dashboard_index
    }
    datasets, national, adm2_map_2010 = load_datasets(data_dir)
    geojson = load_json(geojson_path)
    geo_by_territory = {
        base.territory_key(
            feature.get("properties", {}).get("municipio", ""),
            feature.get("properties", {}).get("provincia", ""),
        ): feature
        for feature in geojson.get("features", [])
    }

    historical_rows = load_json(inventory_path).get("rows", [])
    historical_by_territory: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in historical_rows:
        if row.get("status") == "downloaded":
            historical_by_territory[base.territory_key(row["municipality"], row["province"])].append(row)

    classifier = base.read_classifier_codes(classifier_path)
    wikipedia_cache = load_json(wikipedia_cache_path) if wikipedia_cache_path.exists() else {}
    if not args.skip_wikipedia_refresh:
        base.fetch_wikipedia_exact_batches(targets, wikipedia_cache)
        sanitize_wikipedia_cache(wikipedia_cache, targets)
        fetch_full_wikipedia(wikipedia_cache, targets)
        save_json(wikipedia_cache_path, wikipedia_cache)
    historical_cache = load_json(historical_cache_path) if historical_cache_path.exists() else {}

    previous_manifest = load_json(manifest_path).get("municipalities", []) if manifest_path.exists() else []
    previous_by_id = {row["id"]: row for row in previous_manifest}
    generated_rows: list[dict[str, Any]] = []
    docx_dir.mkdir(parents=True, exist_ok=True)
    assets_dir.mkdir(parents=True, exist_ok=True)
    if args.web_copy:
        web_dir.mkdir(parents=True, exist_ok=True)

    for number, item in enumerate(targets, 1):
        key = base.territory_key(item["municipio"], item["provincia"])
        adm2_code = dashboard_by_territory.get(key, "")
        feature = geo_by_territory.get(key)
        map_adm2_code = (
            base.code_key(feature.get("properties", {}).get("adm2_code"))
            if feature
            else adm2_code
        )
        geographic_code = base.resolve_geographic_code(item, adm2_code, classifier)
        if not geographic_code and previous_by_id.get(item["id"]):
            geographic_code = previous_by_id[item["id"]].get("geographic_code", "")
        data = build_data_bundle(adm2_code, datasets, adm2_map_2010)
        historical = historical_by_territory.get(key, [])
        historical_context = extract_historical_context(historical, historical_cache, key)
        wikipedia = wikipedia_cache.get(key, {"status": "not_found"})
        previous = previous_by_id.get(item["id"], {})
        file_name = previous.get("file_name") or (
            f"{geographic_code or f'id-{item['id']:03d}'}_"
            f"PMD_{base.slugify(item['municipio'])}_Borrador_Tecnico_{PERIOD}.docx"
        )
        output_path = docx_dir / file_name
        asset_dir = assets_dir / base.slugify(item["municipio"])
        if not (args.resume and output_path.exists() and all((asset_dir / f"diagnostico-{i}.png").exists() for i in range(1, 6))):
            pages = generate_dashboard_pages(
                item,
                map_adm2_code,
                data,
                national,
                geojson,
                asset_dir,
            )
            area_km2 = (feature or {}).get("properties", {}).get("km2")
            build_document(
                item,
                geographic_code,
                adm2_code,
                data,
                national,
                wikipedia,
                historical,
                historical_context,
                area_km2,
                pages,
                output_path,
            )
        if args.web_copy:
            shutil.copy2(output_path, web_dir / file_name)
        generated_rows.append(
            {
                "id": item["id"],
                "municipio": item["municipio"],
                "provincia": item["provincia"],
                "region": item["region"],
                "geographic_code": geographic_code,
                "dashboard_adm2_code": adm2_code,
                "dashboard_available": bool(adm2_code),
                "historical_pmd_count": len(historical),
                "historical_periods": [
                    f"{row.get('period_start')}-{row.get('period_end')}" for row in historical
                ],
                "wikipedia_status": wikipedia.get("status"),
                "wikipedia_url": wikipedia.get("url", ""),
                "content_version": "2.0-rich-diagnostic",
                "information_general_status": "precompleted",
                "diagnostic_status": "precompleted" if adm2_code else "source-gap",
                "file_name": file_name,
                "relative_url": f"downloads/pmd-borradores/{file_name}",
            }
        )
        print(f"[{number}/{len(targets)}] {item['municipio']} -> {output_path.name}")

    save_json(historical_cache_path, historical_cache)
    if args.municipality:
        generated_ids = {row["id"] for row in generated_rows}
        combined = [row for row in previous_manifest if row["id"] not in generated_ids] + generated_rows
    else:
        combined = generated_rows
    combined = sorted(combined, key=lambda row: row["id"])
    manifest = {
        "generated_at": TODAY.isoformat(),
        "period": PERIOD,
        "content_version": "2.0-rich-diagnostic",
        "target_rule": "No PMD official evidence and no existing 7-12/draft",
        "expected_target_count": 104,
        "generated_count": len(combined),
        "municipalities": combined,
    }
    save_json(manifest_path, manifest)
    save_json(portal_repo / "public" / "data" / "generated-pmd-drafts.json", manifest)

    csv_path = output_dir / "manifest.csv"
    columns = [
        "id",
        "municipio",
        "provincia",
        "region",
        "geographic_code",
        "dashboard_adm2_code",
        "dashboard_available",
        "historical_pmd_count",
        "historical_periods",
        "wikipedia_status",
        "wikipedia_url",
        "content_version",
        "information_general_status",
        "diagnostic_status",
        "file_name",
        "relative_url",
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for row in combined:
            copy = dict(row)
            copy["historical_periods"] = "; ".join(copy.get("historical_periods", []))
            writer.writerow(copy)
    print(
        json.dumps(
            {
                "generated": len(generated_rows),
                "manifest_total": len(combined),
                "output_dir": str(output_dir),
                "web_copy": args.web_copy,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
