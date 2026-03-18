from __future__ import annotations

from pathlib import Path
from typing import Dict

import streamlit as st

from config import COLUMNS_BY_KEY, DEFAULT_SHEET, DEFAULT_WORKBOOK, FIELD_KEYS, LOG_FILE, SECTION_LAYOUT
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

options = handler.get_dynamic_options(sheet_name)

if "form_data" not in st.session_state:
    st.session_state.form_data = {key: "" for key in FIELD_KEYS}
if "preview_data" not in st.session_state:
    st.session_state.preview_data = None
if "selected_row" not in st.session_state:
    st.session_state.selected_row = None


def draw_input(key: str) -> str:
    label = COLUMNS_BY_KEY[key].label
    default = st.session_state.form_data.get(key, "")
    choices = options.get(key)
    if choices:
        pick = st.selectbox(f"{label} ({key})", options=choices + ["<Saisie libre>"], index=0 if default in choices else len(choices))
        if pick == "<Saisie libre>":
            return st.text_input(f"{label} libre ({key})", value=default)
        return pick
    if key in {"french_resume", "english_resume", "remarks_n", "remarks_p", "remarks_r"}:
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

required = [key for key in FIELD_KEYS if COLUMNS_BY_KEY[key].required]

if clear_btn:
    st.session_state.form_data = {key: "" for key in FIELD_KEYS}
    st.session_state.preview_data = None
    st.success("Formulaire vidé.")

if preview_btn or add_btn:
    cleaned = sanitize_payload(collected)
    st.session_state.form_data = cleaned
    errors = validate_required(cleaned, required)
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
        st.success(f"Ligne ajoutée avec succès : {sheet_name}!{row}")
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
