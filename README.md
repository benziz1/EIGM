# Requirement Entry App (Excel Front-End)

Application locale Streamlit pour saisir/mettre à jour des exigences dans un fichier Excel existant (`requirements_db.xlsx`) sans modifier sa structure métier.

## Fonctionnalités

- Interface unique de saisie par blocs fonctionnels.
- Ajout de ligne automatique à partir de la première ligne vide (à partir de la ligne 7).
- Prévisualisation avant insertion.
- Copie du style/format/formules de la ligne précédente vers la nouvelle ligne.
- Incrémentation automatique de la colonne `A (#)`.
- Recopie des formules (dont `Q: DECISION`) via translation de références.
- Validation des champs obligatoires.
- Listes déroulantes dynamiques (valeurs existantes de l'onglet + valeurs par défaut).
- Mode édition (recherche + chargement + sauvegarde).
- Mode archivage sécurisé (marquage `[ARCHIVED]` sur le nom FR).
- Historique CSV (`history_log.csv`) : date, utilisateur, action, onglet, ligne.

## Structure

```bash
.
├── app.py
├── excel_handler.py
├── config.py
├── utils.py
├── requirements.txt
├── README.md
└── requirements_db.xlsx   # à fournir (source de vérité)
```

## Prérequis

- Python 3.10+
- Fichier Excel existant `requirements_db.xlsx` placé à la racine du projet.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

## Utilisation

1. Vérifier dans la barre latérale le chemin du fichier Excel et l'onglet cible (`V00` par défaut).
2. Remplir les champs par blocs.
3. Cliquer **Prévisualiser la ligne avant insertion**.
4. Cliquer **Ajouter la ligne**.
5. En cas de succès, la ligne est écrite dans la première ligne vide et journalisée.

## Comportement Excel

- Les en-têtes (lignes 5-6) ne sont jamais modifiés.
- Les données commencent à la ligne 7.
- La nouvelle ligne réutilise la ligne précédente comme modèle (style + format + formules).
- La formule en Q est recopiée et adaptée automatiquement (translation de référence).
- Les champs saisis utilisateur écrasent uniquement les colonnes métiers prévues.

## Limites connues

- Si le fichier est ouvert dans Excel avec verrouillage exclusif, l'enregistrement peut échouer.
- Les fusions de cellules inter-lignes atypiques peuvent nécessiter une adaptation spécifique.

## Personnalisation rapide

- `config.py` : mapping colonnes, champs obligatoires, options par défaut, sections d'interface.
- `excel_handler.py` : logique d'accès Excel (ajout, recherche, update, archivage).
- `utils.py` : sanitation, validation, journalisation.

