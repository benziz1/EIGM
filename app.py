from __future__ import annotations
from html import escape
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
import utils
PAGE_SIZE = 5
def requirement_number_base(identifier: str) -> str:
    if hasattr(utils, "requirement_number_base"):
        return utils.requirement_number_base(identifier)
    value = str(identifier or "").strip()
    if not value:
        return ""
    parts = value.split("_")
    if len(parts) >= 4:
        return "_".join(parts[:-1])
    return value
def describe_payload_changes(previous_payload: Dict[str, str], current_payload: Dict[str, str]) -> str:
    if hasattr(utils, "describe_payload_changes"):
        return utils.describe_payload_changes(previous_payload, current_payload)
    changes = []
    for key in FIELD_KEYS:
        if key == "number":
            continue
        before = str(previous_payload.get(key, "") or "").strip()
        after = str(current_payload.get(key, "") or "").strip()
        if before != after:
            label = COLUMNS_BY_KEY[key].label
            changes.append(f"{label}: '{before or '∅'}' → '{after or '∅'}'")
    return " | ".join(changes) if changes else "Aucune différence métier détectée"
def read_history(log_file: Path, limit: int = 20, number_base_filter: str = "") -> list[Dict[str, str]]:
    if hasattr(utils, "read_history"):
        return utils.read_history(log_file, limit=limit, number_base_filter=number_base_filter)
    import csv
    if not log_file.exists():
        return []
    entries: list[Dict[str, str]] = []
    with log_file.open("r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if number_base_filter and row.get("number_base", "") != number_base_filter:
                continue
            entries.append(dict(row))
    entries.sort(key=lambda item: item.get("timestamp", ""), reverse=True)
    return entries[:limit]
def sanitize_payload(payload: Dict[str, str]) -> Dict[str, str]:
    return utils.sanitize_payload(payload)
def validate_required(payload: Dict[str, str], required_keys: list[str]) -> list[str]:
    return utils.validate_required(payload, required_keys)
def append_history(log_file: Path, action: str, row_number: int, identifier: str, sheet_name: str, changes: str = "") -> None:
    try:
        utils.append_history(log_file, action, row_number, identifier, sheet_name, changes=changes)
    except TypeError:
        utils.append_history(log_file, action, row_number, identifier, sheet_name)
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
        [data-baseweb="select"] > div,
        [data-baseweb="select"] span,
        [data-baseweb="select"] input,
        [data-testid="stSelectbox"] div,
        [data-testid="stSelectbox"] span,
        [data-testid="stSelectbox"] svg,
        [data-testid="stSelectbox"] p,
        [data-testid="stSelectbox"] label {{
            color: {colors['text']} !important;
            fill: {colors['text']} !important;
        }}
        [data-baseweb="select"] > div,
        [data-baseweb="popover"],
        [role="listbox"],
        [role="option"] {{
            background: {colors['surface']} !important;
            border-color: {colors['border']} !important;
        }}
        [data-baseweb="popover"] *,
        [role="listbox"] *,
        [role="option"] * {{
            color: {colors['text']} !important;
            fill: {colors['text']} !important;
        }}
        .stFormSubmitButton > button {{
            background: {colors['surface_alt']} !important;
            color: {colors['text']} !important;
            border: 1px solid {colors['border']} !important;
        }}
        .stFormSubmitButton > button:hover {{
            background: {colors['primary']} !important;
            color: {colors['primary_text']} !important;
            border-color: {colors['primary']} !important;
        }}
        [data-testid="stToolbar"] button,
        [data-testid="stToolbar"] svg,
        [data-testid="stElementToolbar"] button,
        [data-testid="stElementToolbar"] svg {{
            color: {colors['text']} !important;
            fill: {colors['text']} !important;
        }}
        [data-baseweb="menu"],
        [role="menu"] {{
            background: {colors['surface']} !important;
            border: 1px solid {colors['border']} !important;
        }}
        [data-baseweb="menu"] *,
        [role="menu"] *,
        [role="menuitem"] *,
        [data-baseweb="menu"] svg,
        [role="menu"] svg {{
            color: {colors['text']} !important;
            fill: {colors['text']} !important;
        }}
        .history-entry {{
            position: relative;
            background: {colors['surface']};
            border: 1px solid {colors['border']};
            border-radius: 12px;
            padding: 0.85rem 1rem;
            margin-bottom: 0.65rem;
        }}
        .history-title {{
            font-weight: 700;
            margin-bottom: 0.2rem;
        }}
        .history-sub {{
            color: {colors['muted']};
            font-size: 0.92rem;
        }}
        .history-tooltip {{
            visibility: hidden;
            opacity: 0;
            transition: opacity 0.2s ease;
            position: absolute;
            left: 1rem;
            right: 1rem;
            top: calc(100% + 0.35rem);
            z-index: 30;
            background: {colors['surface_alt']};
            color: {colors['text']};
            border: 1px solid {colors['border']};
            border-radius: 10px;
            padding: 0.7rem 0.8rem;
            box-shadow: 0 10px 22px rgba(0, 0, 0, 0.18);
            white-space: normal;
        }}
        .history-entry:hover .history-tooltip {{
            visibility: visible;
            opacity: 1;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )
st.set_page_config(page_title="Requirement Entry", layout="wide")
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
    st.session_state.theme_mode = "Dark"
if "theme_toggle" not in st.session_state:
    st.session_state.theme_toggle = st.session_state.theme_mode == "Dark"
inject_theme("Dark" if st.session_state.theme_toggle else "Light")
title_col, theme_col = st.columns([8, 1.4])
with title_col:
    st.title("Masque de saisie des exigences (Excel)")
with theme_col:
    theme_is_dark = st.toggle("Dark", value=st.session_state.theme_toggle, key="theme_toggle_control")
if theme_is_dark != st.session_state.theme_toggle:
    st.session_state.theme_toggle = theme_is_dark
    st.session_state.theme_mode = "Dark" if theme_is_dark else "Light"
    st.rerun()
st.session_state.theme_mode = "Dark" if st.session_state.theme_toggle else "Light"
if st.sidebar.button("Exigences", use_container_width=True):
    st.session_state.current_page = "exigences"
if st.sidebar.button("Archivage", use_container_width=True):
    st.session_state.current_page = "archivage"
if st.sidebar.button("Historique", use_container_width=True):
    st.session_state.current_page = "historique"
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
if "original_form_data" not in st.session_state:
    st.session_state.original_form_data = {key: "" for key in FIELD_KEYS}
TEXTAREA_KEYS = {"french_resume", "english_resume", "remarks_n", "remarks_p", "remarks_r"}
required = [key for key in FIELD_KEYS if COLUMNS_BY_KEY[key].required and key != "number"]
def is_editing() -> bool:
    return st.session_state.current_page == "edit_requirement" and st.session_state.selected_row is not None
def reset_form() -> None:
    st.session_state.form_data = {key: "" for key in FIELD_KEYS}
    st.session_state.original_form_data = {key: "" for key in FIELD_KEYS}
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
def render_history_entries(entries: list[Dict[str, str]]) -> None:
    if not entries:
        st.info("Aucun historique disponible pour le moment.")
        return
    cards = []
    for entry in entries:
        details = entry.get("changes", "") or entry.get("action", "")
        tooltip_body = escape(details, quote=True).replace(" | ", "<br>")
        tooltip_title = escape(details.replace(" | ", "\n"), quote=True)
        cards.append(
            f"""
            <div class="history-entry" title="{tooltip_title}">
                <div class="history-title">{escape(entry.get('action', ''))} — {escape(entry.get('number', ''))}</div>
                <div class="history-sub">{escape(entry.get('timestamp', ''))} · {escape(entry.get('user', ''))} · {escape(entry.get('sheet', ''))}</div>
                <div class="history-tooltip">{tooltip_body}</div>
            </div>
            """
        )
    st.markdown("".join(cards), unsafe_allow_html=True)
    st.caption("Survolez une entrée d'historique pour afficher le détail des changements quand il s'agit d'une mise à jour.")
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
            append_history(LOG_FILE, "ADD", row, st.session_state.preview_data.get("number", ""), sheet_name, changes="Création de l'exigence")
            st.success(f"Ligne ajoutée en tête avec succès : {sheet_name}!{row}")
            reset_form()
            st.session_state.current_page = "exigences"
        except ExcelHandlerError as exc:
            st.error(str(exc))
if st.session_state.current_page == "exigences":
    st.subheader("Exigences")
    st.markdown('<div class="theme-muted">Créez rapidement une nouvelle exigence depuis cette page.</div>', unsafe_allow_html=True)
    if st.button("Ajouter une exigence", type="primary", use_container_width=True):
        go_to_add_page()
    st.markdown("### Recherche et modification d'exigence")
    st.markdown('<div class="theme-muted">Recherchez, filtrez puis modifiez les exigences existantes.</div>', unsafe_allow_html=True)
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
        start = (st.session_state.search_page - 1) * PAGE_SIZE
        visible_rows = rows[start : start + PAGE_SIZE]
        st.caption(f"{len(rows)} exigence(s) trouvée(s). Affichage de {len(visible_rows)} résultat(s) sur cette page.")
        if visible_rows:
            for row in visible_rows:
                row_label = f"{row['number']} — {row['french_name']}"
                if st.button(row_label, key=f"edit_row_{row['row']}", use_container_width=True):
                    st.session_state.form_data = handler.load_requirement(sheet_name, row["row"])
                    st.session_state.original_form_data = dict(st.session_state.form_data)
                    st.session_state.selected_row = row["row"]
                    st.session_state.preview_data = None
                    st.session_state.current_page = "edit_requirement"
                    st.rerun()
        page_col1, page_col2, page_col3 = st.columns([1, 2, 1])
        with page_col1:
            if st.button("← Page précédente", disabled=st.session_state.search_page <= 1, use_container_width=True):
                st.session_state.search_page -= 1
                st.rerun()
        with page_col2:
            st.markdown(
                f"<div style='text-align:center; padding-top:0.4rem;'>Page {st.session_state.search_page} / {total_pages}</div>",
                unsafe_allow_html=True,
            )
        with page_col3:
            if st.button("Page suivante →", disabled=st.session_state.search_page >= total_pages, use_container_width=True):
                st.session_state.search_page += 1
                st.rerun()
        if not visible_rows:
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
    requirement_history = read_history(LOG_FILE, limit=5, number_base_filter=requirement_number_base(current_number))
    st.markdown("### Dernières modifications")
    if requirement_history:
        latest_history = requirement_history[0]
        st.info(
            f"Dernière modification : {latest_history['timestamp']} par {latest_history['user']}\n\n"
            f"{latest_history['changes'] or latest_history['action']}"
        )
        with st.expander("Voir l'historique récent de cette exigence"):
            render_history_entries(requirement_history)
    else:
        st.caption("Aucun historique enregistré pour cette exigence pour le moment.")
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
        updated_payload = dict(st.session_state.preview_data)
        st.session_state.form_data["number"] = updated_number
        changes = describe_payload_changes(st.session_state.original_form_data, updated_payload)
        changes = f"Version: {current_number} → {updated_number} | {changes}"
        append_history(LOG_FILE, "UPDATE", st.session_state.selected_row, updated_number, sheet_name, changes=changes)
        st.session_state.original_form_data = dict(updated_payload)
        st.session_state.original_form_data["number"] = updated_number
        st.session_state.preview_data = None
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
                    append_history(LOG_FILE, "ARCHIVE", row, rows[labels.index(choice)]["number"], sheet_name, changes="Archivage de l'exigence")
                    st.success("Exigence archivée.")
            else:
                st.info("Aucun résultat.")
        except ExcelHandlerError as exc:
            st.error(str(exc))
if st.session_state.current_page == "historique":
    st.subheader("Historique")
    history_entries = read_history(LOG_FILE, limit=30)
    render_history_entries(history_entries)
