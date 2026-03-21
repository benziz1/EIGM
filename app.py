from __future__ import annotations

from pathlib import Path
from typing import Dict

import streamlit as st

from config import (
    COLUMNS_BY_KEY,
    DEFAULT_SHEET,
    DEFAULT_WORKBOOK,
    FIELD_KEYS,
    LOG_FILE,
    SECTION_LAYOUT,
    TYPE_FREE_OPTION,
)
from excel_handler import ExcelHandler, ExcelHandlerError
from utils import append_history, sanitize_payload, validate_required

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
mode = st.sidebar.radio("Mode", ["Ajout", "Édition", "Archivage"], horizontal=True)

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


def compute_number_preview(type_value: str, fallback_number: str = "") -> str:
    if mode == "Édition" and st.session_state.selected_row:
        return fallback_number or st.session_state.form_data.get("number", "")
    if not type_value.strip():
        return ""
    try:
        return handler.preview_next_number(sheet_name, type_value)
    except ExcelHandlerError:
        return fallback_number


current_type = st.session_state.form_data.get("type", "")
number_preview = compute_number_preview(current_type, st.session_state.form_data.get("number", ""))
st.caption(
    "Les nouvelles exigences sont toujours insérées en tête de la zone de données (ligne 7),"
    " sans utiliser les lignes vides existantes plus bas dans l'onglet."
)


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


TEXTAREA_KEYS = {"french_resume", "english_resume", "remarks_n", "remarks_p", "remarks_r"}


def draw_input(key: str) -> str:
    label = COLUMNS_BY_KEY[key].label
    default = st.session_state.form_data.get(key, "")

    if key == "number":
        st.text_input(
            "NUMBER (généré automatiquement)",
            value=compute_number_preview(st.session_state.form_data.get("type", ""), default),
            disabled=True,
        )
        return compute_number_preview(st.session_state.form_data.get("type", ""), default)

    if key == "type":
        return draw_type_selector(default)

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


with st.form("requirement_form"):
    collected: Dict[str, str] = {}
    for section, keys in SECTION_LAYOUT.items():
        st.subheader(section)
        cols = st.columns(2)
        for idx, key in enumerate(keys):
            with cols[idx % 2]:
                collected[key] = draw_input(key)

    preview_btn = st.form_submit_button("Prévisualiser la ligne avant insertion")
    add_btn = st.form_submit_button("Ajouter la ligne")
    clear_btn = st.form_submit_button("Vider le formulaire")

required = [key for key in FIELD_KEYS if COLUMNS_BY_KEY[key].required and key != "number"]

if clear_btn:
    st.session_state.form_data = {key: "" for key in FIELD_KEYS}
    st.session_state.preview_data = None
    st.session_state.selected_row = None
    st.success("Formulaire vidé.")

if preview_btn or add_btn:
    cleaned = sanitize_payload(collected)
    cleaned["number"] = compute_number_preview(cleaned.get("type", ""), cleaned.get("number", ""))
    st.session_state.form_data = cleaned
    errors = validate_required(cleaned, required)
    if not cleaned.get("number"):
        errors.append("Le NUMBER n'a pas pu être généré automatiquement. Vérifiez le TYPE.")
    if errors:
        st.error("\n".join(errors))
    else:
        st.session_state.preview_data = cleaned

if st.session_state.preview_data:
    st.markdown("### Prévisualisation")
    st.dataframe([st.session_state.preview_data], use_container_width=True)

if add_btn and st.session_state.preview_data:
    try:
        row = handler.add_requirement(sheet_name, st.session_state.preview_data)
        append_history(LOG_FILE, "ADD", row, st.session_state.preview_data.get("number", ""), sheet_name)
        st.success(f"Ligne ajoutée en tête avec succès : {sheet_name}!{row}")
        st.session_state.form_data = {key: "" for key in FIELD_KEYS}
        st.session_state.preview_data = None
    except ExcelHandlerError as exc:
        st.error(str(exc))

if mode == "Édition":
    st.markdown("---")
    st.subheader("Mode édition")
    keyword = st.text_input("Rechercher (NUMBER / FRENCH NAME / ENGLISH NAME)")
    if keyword:
        try:
            rows = handler.search_requirements(sheet_name, keyword)
            if not rows:
                st.info("Aucun résultat.")
            else:
                labels = [f"Ligne {r['row']} - {r['number']} - {r['french_name']}" for r in rows]
                choice = st.selectbox("Résultats", labels)
                selected = rows[labels.index(choice)]
                st.session_state.selected_row = selected["row"]
                if st.button("Charger dans le formulaire"):
                    st.session_state.form_data = handler.load_requirement(sheet_name, selected["row"])
                    st.session_state.preview_data = None
                    st.success("Données chargées dans le formulaire principal.")
                if st.button("Sauvegarder mise à jour depuis le formulaire"):
                    payload = sanitize_payload(st.session_state.form_data)
                    handler.update_requirement(sheet_name, selected["row"], payload)
                    append_history(LOG_FILE, "UPDATE", selected["row"], payload.get("number", ""), sheet_name)
                    st.success("Ligne mise à jour.")
        except ExcelHandlerError as exc:
            st.error(str(exc))

if mode == "Archivage":
    st.markdown("---")
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
