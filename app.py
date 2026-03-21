from __future__ import annotations

from math import ceil
from pathlib import Path
from typing import Dict

import streamlit as st

from config import (
    COLUMNS_BY_KEY,
    DEFAULT_SHEET,
    DEFAULT_WORKBOOK,
    FIELD_KEYS,
    LOG_FILE,
    MILESTONE_KEYS,
    MILESTONE_MARK,
    SECTION_LAYOUT,
    TYPE_FREE_OPTION,
)
from excel_handler import ExcelHandler, ExcelHandlerError
from utils import append_history, sanitize_payload, validate_required

PAGE_SIZE = 5

THEMES = {
    "Light": {
        "bg": "#eef4fb",
        "surface": "#ffffff",
        "surface_alt": "#dfeaf7",
        "text": "#12324a",
        "muted": "#5d7891",
        "border": "#bfd2e6",
        "primary": "#4f7ea8",
        "primary_hover": "#436d92",
        "primary_text": "#f7fbff",
    },
    "Dark": {
        "bg": "#0f1a26",
        "surface": "#152435",
        "surface_alt": "#1d3248",
        "text": "#e8f1f8",
        "muted": "#9fb4c7",
        "border": "#294761",
        "primary": "#5f89b3",
        "primary_hover": "#74a0cb",
        "primary_text": "#f7fbff",
    },
}


def inject_theme(theme_name: str) -> None:
    colors = THEMES[theme_name]
    st.markdown(
        f"""
        <style>
        .stApp {{
            background: {colors['bg']};
            color: {colors['text']};
        }}
        [data-testid="stAppViewContainer"] {{
            background: linear-gradient(180deg, {colors['bg']} 0%, {colors['surface_alt']} 100%);
        }}
        [data-testid="stSidebar"] {{
            background: {colors['surface']};
            border-right: 1px solid {colors['border']};
        }}
        [data-testid="stHeader"] {{
            background: transparent;
        }}
        .block-container {{
            padding-top: 1.8rem;
        }}
        h1, h2, h3, h4, h5, h6, p, label, span, div {{
            color: {colors['text']};
        }}
        [data-testid="stForm"], .stAlert, [data-testid="stDataFrame"], .stTextInput, .stTextArea, .stSelectbox {{
            background: rgba(255, 255, 255, 0.02);
            border-radius: 14px;
        }}
        .stButton > button {{
            border-radius: 12px;
            border: 1px solid {colors['border']};
            color: {colors['text']};
            background: {colors['surface']};
            min-height: 2.8rem;
        }}
        .stButton > button:hover {{
            border-color: {colors['primary']};
            color: {colors['primary']};
        }}
        .stButton > button[kind="primary"] {{
            background: {colors['primary']};
            color: {colors['primary_text']};
            border: 1px solid {colors['primary']};
            font-size: 1.05rem;
            font-weight: 700;
            min-height: 4rem;
            box-shadow: 0 10px 24px rgba(0, 0, 0, 0.12);
        }}
        .stButton > button[kind="primary"]:hover {{
            background: {colors['primary_hover']};
            border-color: {colors['primary_hover']};
            color: {colors['primary_text']};
        }}
        .theme-card {{
            background: {colors['surface']};
            border: 1px solid {colors['border']};
            border-radius: 16px;
            padding: 0.8rem 1rem;
            margin-bottom: 1rem;
            color: {colors['text']};
        }}
        .theme-muted {{
            color: {colors['muted']};
            font-size: 0.95rem;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

st.set_page_config(page_title="Requirement Entry", layout="wide")
st.title("Masque de saisie des exigences (Excel)")

workbook_path = Path(st.sidebar.text_input("Fichier Excel", str(DEFAULT_WORKBOOK)))
handler = ExcelHandler(workbook_path)

try:
    sheets = handler.list_sheets()
except ExcelHandlerError as exc:
    st.error(str(exc))
    st.stop()

sheet_index = sheets.index(DEFAULT_SHEET) if DEFAULT_SHEET in sheets else 0
sheet_name = st.sidebar.selectbox("Onglet cible", sheets, index=sheet_index)

if "current_page" not in st.session_state:
    st.session_state.current_page = "exigences"
if "search_page" not in st.session_state:
    st.session_state.search_page = 1
if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "Light"

st.sidebar.markdown(
    '<div class="theme-card"><strong>Thème</strong><div class="theme-muted">Nuances bleues douces pour le confort visuel.</div></div>',
    unsafe_allow_html=True,
)
st.session_state.theme_mode = st.sidebar.radio(
    "Mode d'affichage",
    ["Light", "Dark"],
    index=0 if st.session_state.theme_mode == "Light" else 1,
    horizontal=True,
)
inject_theme(st.session_state.theme_mode)

if st.sidebar.button("Exigences", use_container_width=True):
    st.session_state.current_page = "exigences"
if st.sidebar.button("Archivage", use_container_width=True):
    st.session_state.current_page = "archivage"

try:
    options = handler.get_dynamic_options(sheet_name)
except ExcelHandlerError as exc:
    st.error(str(exc))
    st.stop()

if "form_data" not in st.session_state:
    st.session_state.form_data = {key: "" for key in FIELD_KEYS}
if "preview_data" not in st.session_state:
    st.session_state.preview_data = None
if "selected_row" not in st.session_state:
    st.session_state.selected_row = None


TEXTAREA_KEYS = {"french_resume", "english_resume", "remarks_n", "remarks_p", "remarks_r"}
required = [key for key in FIELD_KEYS if COLUMNS_BY_KEY[key].required and key != "number"]


def is_editing() -> bool:
    return st.session_state.current_page == "edit_requirement" and st.session_state.selected_row is not None


def reset_form() -> None:
    st.session_state.form_data = {key: "" for key in FIELD_KEYS}
    st.session_state.preview_data = None
    st.session_state.selected_row = None


def go_to_add_page() -> None:
    reset_form()
    st.session_state.current_page = "saisie"


def go_to_search_page() -> None:
    st.session_state.preview_data = None
    st.session_state.current_page = "exigences"


def compute_number_preview(type_value: str, fallback_number: str = "") -> str:
    if is_editing():
        current_number = fallback_number or st.session_state.form_data.get("number", "")
        if not current_number:
            return ""
        try:
            return handler.preview_updated_number(current_number, type_value or st.session_state.form_data.get("type", ""))
        except ExcelHandlerError:
            return current_number
    if not type_value.strip():
        return ""
    try:
        return handler.preview_next_number(sheet_name, type_value)
    except ExcelHandlerError:
        return fallback_number


def draw_type_selector(default: str) -> str:
    available_types = []
    for value in options.get("type", []):
        if value not in available_types:
            available_types.append(value)
    selectable_types = [value for value in available_types if value in {"GEN", "AUT"}]
    if "GEN" not in selectable_types:
        selectable_types.insert(0, "GEN")
    if "AUT" not in selectable_types:
        selectable_types.insert(1 if selectable_types else 0, "AUT")

    if default and default not in selectable_types:
        choice_index = len(selectable_types)
        free_default = default
    else:
        choice_index = selectable_types.index(default) if default in selectable_types else 0
        free_default = ""

    selected = st.selectbox(
        "TYPE d'identification",
        options=selectable_types + [TYPE_FREE_OPTION],
        index=choice_index,
    )
    if selected == TYPE_FREE_OPTION:
        return st.text_input("TYPE libre", value=free_default)
    return selected


def draw_input(key: str) -> str:
    label = COLUMNS_BY_KEY[key].label
    default = st.session_state.form_data.get(key, "")

    if key == "number":
        preview_number = compute_number_preview(st.session_state.form_data.get("type", ""), default)
        help_text = "En création : suffixe 00. En modification : seul le dernier nombre est incrémenté automatiquement."
        st.text_input("NUMBER (généré automatiquement)", value=preview_number, disabled=True, help=help_text)
        return preview_number

    if key == "type":
        return draw_type_selector(default)

    if key in MILESTONE_KEYS:
        checked = str(default).strip().upper() in {MILESTONE_MARK, "X", "✗", "TRUE", "1", "YES", "OUI"}
        is_selected = st.checkbox(label, value=checked)
        return MILESTONE_MARK if is_selected else ""

    choices = options.get(key)
    if choices:
        available_choices = list(dict.fromkeys(choices))
        if default and default not in available_choices:
            selected_index = len(available_choices)
        else:
            selected_index = available_choices.index(default) if default in available_choices else 0
        pick = st.selectbox(
            f"{label} ({key})",
            options=available_choices + [TYPE_FREE_OPTION],
            index=selected_index,
        )
        if pick == TYPE_FREE_OPTION:
            return st.text_input(f"{label} libre ({key})", value=default)
        return pick

    if key in TEXTAREA_KEYS:
        return st.text_area(f"{label} ({key})", value=default, height=90)
    return st.text_input(f"{label} ({key})", value=default)


st.caption(
    "Les nouvelles exigences sont toujours insérées en tête de la zone de données (ligne 7),"
    " sans utiliser les lignes vides existantes plus bas dans l'onglet."
)


def render_requirement_form(form_key: str, submit_label: str, preview_label: str) -> tuple[bool, bool, bool, Dict[str, str]]:
    with st.form(form_key):
        collected: Dict[str, str] = {}
        for section, keys in SECTION_LAYOUT.items():
            st.subheader(section)
            if section == "Bloc 5 — Jalons projet":
                milestone_cols = st.columns(len(keys))
                for idx, key in enumerate(keys):
                    with milestone_cols[idx]:
                        collected[key] = draw_input(key)
                continue

            cols = st.columns(2)
            for idx, key in enumerate(keys):
                with cols[idx % 2]:
                    collected[key] = draw_input(key)

        preview_btn = st.form_submit_button(preview_label)
        submit_btn = st.form_submit_button(submit_label)
        clear_btn = st.form_submit_button("Vider le formulaire")
    return preview_btn, submit_btn, clear_btn, collected


def process_form_submission(collected: Dict[str, str]) -> tuple[Dict[str, str], list[str]]:
    cleaned = sanitize_payload(collected)
    cleaned["number"] = compute_number_preview(cleaned.get("type", ""), cleaned.get("number", ""))
    st.session_state.form_data = cleaned
    errors = validate_required(cleaned, required)
    if not cleaned.get("number"):
        errors.append("Le NUMBER n'a pas pu être généré automatiquement. Vérifiez le TYPE.")
    return cleaned, errors


def render_preview() -> None:
    if st.session_state.preview_data:
        st.markdown("### Prévisualisation")
        st.dataframe([st.session_state.preview_data], use_container_width=True)


if st.session_state.current_page == "saisie":
    header_col, plus_col = st.columns([12, 1])
    with header_col:
        st.subheader("Saisie d'une nouvelle exigence")
    with plus_col:
        if st.button("←", help="Retour à la recherche des exigences"):
            go_to_search_page()

    preview_btn, add_btn, clear_btn, collected = render_requirement_form(
        form_key="add_requirement_form",
        submit_label="Ajouter la ligne",
        preview_label="Prévisualiser la ligne avant insertion",
    )

    if clear_btn:
        reset_form()
        st.success("Formulaire vidé.")

    if preview_btn or add_btn:
        cleaned, errors = process_form_submission(collected)
        if errors:
            st.error("\n".join(errors))
        else:
            st.session_state.preview_data = cleaned

    render_preview()

    if add_btn and st.session_state.preview_data:
        try:
            row = handler.add_requirement(sheet_name, st.session_state.preview_data)
            append_history(LOG_FILE, "ADD", row, st.session_state.preview_data.get("number", ""), sheet_name)
            st.success(f"Ligne ajoutée en tête avec succès : {sheet_name}!{row}")
            reset_form()
            st.session_state.current_page = "exigences"
        except ExcelHandlerError as exc:
            st.error(str(exc))

if st.session_state.current_page == "exigences":
    st.subheader("Exigences")
    st.markdown('<div class="theme-muted">Recherche, filtrage et accès rapide à la création d\'une nouvelle exigence.</div>', unsafe_allow_html=True)
    if st.button("Ajouter une exigence", type="primary", use_container_width=True):
        go_to_add_page()

    filter_col1, filter_col2, filter_col3 = st.columns(3)
    with filter_col1:
        edit_keyword = st.text_input("Recherche globale", placeholder="NUMBER, nom FR, nom EN...")
    with filter_col2:
        type_filter_values = [""] + list(dict.fromkeys(options.get("type", [])))
        edit_type_filter = st.selectbox("Filtre TYPE", type_filter_values, format_func=lambda x: x or "Tous")
    with filter_col3:
        applicability_filter_values = [""] + list(dict.fromkeys(options.get("applicability", [])))
        edit_applicability_filter = st.selectbox(
            "Filtre APPLICABILITY",
            applicability_filter_values,
            format_func=lambda x: x or "Tous",
        )

    include_archived = st.checkbox("Inclure les exigences archivées", value=False)

    try:
        rows = handler.list_requirements(
            sheet_name=sheet_name,
            keyword=edit_keyword,
            type_filter=edit_type_filter,
            applicability_filter=edit_applicability_filter,
            include_archived=include_archived,
        )
        total_pages = max(1, ceil(len(rows) / PAGE_SIZE))
        if st.session_state.search_page > total_pages:
            st.session_state.search_page = total_pages

        page_col1, page_col2, page_col3 = st.columns([1, 2, 1])
        with page_col1:
            if st.button("← Page précédente", disabled=st.session_state.search_page <= 1, use_container_width=True):
                st.session_state.search_page -= 1
        with page_col2:
            st.markdown(
                f"<div style='text-align:center; padding-top:0.4rem;'>Page {st.session_state.search_page} / {total_pages}</div>",
                unsafe_allow_html=True,
            )
        with page_col3:
            if st.button("Page suivante →", disabled=st.session_state.search_page >= total_pages, use_container_width=True):
                st.session_state.search_page += 1

        start = (st.session_state.search_page - 1) * PAGE_SIZE
        visible_rows = rows[start : start + PAGE_SIZE]
        st.caption(f"{len(rows)} exigence(s) trouvée(s). Affichage de {len(visible_rows)} résultat(s) sur cette page.")

        if visible_rows:
            for row in visible_rows:
                row_label = f"{row['number']} — {row['french_name']}"
                if st.button(row_label, key=f"edit_row_{row['row']}", use_container_width=True):
                    st.session_state.form_data = handler.load_requirement(sheet_name, row["row"])
                    st.session_state.selected_row = row["row"]
                    st.session_state.preview_data = None
                    st.session_state.current_page = "edit_requirement"
                    st.rerun()
        else:
            st.info("Aucune exigence ne correspond aux filtres courants.")
    except ExcelHandlerError as exc:
        st.error(str(exc))

if st.session_state.current_page == "edit_requirement" and st.session_state.selected_row:
    header_col, back_col = st.columns([12, 1])
    with header_col:
        st.subheader(f"Modification de l'exigence ligne {st.session_state.selected_row}")
    with back_col:
        if st.button("←", help="Retour à la recherche"):
            go_to_search_page()
            st.rerun()

    current_number = st.session_state.form_data.get("number", "")
    next_revision = compute_number_preview(st.session_state.form_data.get("type", ""), current_number)
    st.info(f"Révision suivante : {next_revision}")

    preview_btn, save_btn, clear_btn, collected = render_requirement_form(
        form_key="edit_requirement_form",
        submit_label="Enregistrer les modifications",
        preview_label="Prévisualiser la modification",
    )

    if clear_btn:
        reset_form()
        st.success("Formulaire de modification vidé.")
        go_to_search_page()

    if preview_btn or save_btn:
        cleaned, errors = process_form_submission(collected)
        if errors:
            st.error("\n".join(errors))
        else:
            st.session_state.preview_data = cleaned

    render_preview()

    if save_btn and st.session_state.preview_data:
        updated_number = handler.update_requirement(
            sheet_name,
            st.session_state.selected_row,
            st.session_state.preview_data,
        )
        st.session_state.form_data["number"] = updated_number
        st.session_state.preview_data = None
        append_history(LOG_FILE, "UPDATE", st.session_state.selected_row, updated_number, sheet_name)
        st.success(f"Exigence mise à jour. Nouveau NUMBER : {updated_number}")

if st.session_state.current_page == "archivage":
    st.subheader("Mode archivage sécurisé")
    keyword = st.text_input("Rechercher la ligne à archiver")
    if keyword:
        try:
            rows = handler.search_requirements(sheet_name, keyword)
            if rows:
                labels = [f"Ligne {r['row']} - {r['number']} - {r['french_name']}" for r in rows]
                choice = st.selectbox("Sélection", labels, key="archive_select")
                row = rows[labels.index(choice)]["row"]
                confirm = st.checkbox("Je confirme l'archivage")
                if st.button("Archiver") and confirm:
                    handler.archive_requirement(sheet_name, row)
                    append_history(LOG_FILE, "ARCHIVE", row, rows[labels.index(choice)]["number"], sheet_name)
                    st.success("Exigence archivée.")
            else:
                st.info("Aucun résultat.")
        except ExcelHandlerError as exc:
            st.error(str(exc))
