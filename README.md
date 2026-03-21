# Requirement Entry App (Excel Front-End)

Application locale Streamlit pour saisir/mettre à jour des exigences dans un fichier Excel existant (`requirements_db.xlsx`) sans modifier sa structure métier.

## Fonctionnalités

- Interface unique de saisie par blocs fonctionnels, avec thèmes clair/sombre en nuances de bleu.
- Les nouvelles exigences sont insérées **tout en haut de la zone de données** (ligne 7), en décalant les lignes existantes vers le bas.
- Les lignes vides plus bas dans l'onglet sont ignorées pendant l'ajout.
- Le champ `NUMBER` n'est plus saisi manuellement : il est généré automatiquement à partir du `TYPE` sélectionné.
- Le format généré est `REQ_type_XXX_XX`, par exemple `REQ_GEN_012_00` à la création.
- Le `TYPE` d'identification propose `GEN`, `AUT` ou une saisie libre.
- Le champ `IADT` propose maintenant `DEMONSTRATION` en valeur suggérée.
- Le bloc 5 utilise des cases à cocher (une par jalon) qui écrivent une croix `X` dans Excel lorsqu'elles sont cochées.
- Prévisualisation avant insertion.
- Copie du style/format/formules de la ligne précédemment en tête vers la nouvelle ligne.
- Incrémentation automatique de la colonne `A (#)` à partir du maximum déjà présent.
- Recopie des formules (dont `Q: DECISION`) via translation de références.
- Validation des champs obligatoires.
- Listes déroulantes dynamiques (valeurs existantes de l'onglet + valeurs par défaut).
- Mode édition accessible depuis le bandeau latéral gauche, avec recherche filtrée paginée (5 résultats par page), liste cliquable et page de modification dédiée.
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
2. Sélectionner le `TYPE` d'identification (`GEN`, `AUT` ou saisie libre).
3. Vérifier le `NUMBER` généré automatiquement dans le bloc 1.
4. Remplir les autres champs par blocs, puis cocher les jalons du bloc 5 à marquer par une croix dans Excel.
5. Cliquer **Prévisualiser la ligne avant insertion**.
6. Cliquer **Ajouter la ligne**.
7. En cas de succès, la ligne est écrite en tête de la zone de données et journalisée.
8. Depuis le bouton latéral **Exigences**, ouvrir la page de recherche, naviguer dans les pages de 5 résultats, puis cliquer directement sur l'exigence à modifier pour ouvrir sa page dédiée.
9. Utiliser le grand bouton bleu **Ajouter une exigence** dans la rubrique **Exigences** pour créer une nouvelle exigence.
10. Choisir au besoin le thème **Light** ou **Dark** en barre latérale.

## Comportement Excel

- Les en-têtes (lignes 5-6) ne sont jamais modifiés.
- Les données commencent à la ligne 7.
- Chaque nouvel ajout insère une nouvelle ligne en 7 et pousse les anciennes lignes vers le bas.
- La nouvelle ligne réutilise la ligne qui était précédemment en tête comme modèle (style + format + formules).
- La formule en Q est recopiée et adaptée automatiquement (translation de référence).
- Les champs saisis utilisateur écrasent uniquement les colonnes métiers prévues.
- Les jalons `SOR` à `FAI` sont alimentés par cases à cocher et écrivent `X` si cochés, sinon une cellule vide.
- Le `NUMBER` suit la forme `REQ_<TYPE>_<index sur 3 chiffres>_00` à la création.
- Lors d'une modification, le `NUMBER` garde exactement la même base (`REQ_<TYPE>_<index>`) et seuls les deux derniers chiffres sont incrémentés (`00` → `01` → `02`, etc.).

## Limites connues

- Si le fichier est ouvert dans Excel avec verrouillage exclusif, l'enregistrement peut échouer.
- Les fusions de cellules inter-lignes atypiques peuvent nécessiter une adaptation spécifique.

## Personnalisation rapide

- `config.py` : mapping colonnes, champs obligatoires, options par défaut, sections d'interface.
- `excel_handler.py` : logique d'accès Excel (ajout en tête, génération du NUMBER, recherche, update, archivage).
- `utils.py` : sanitation, validation, journalisation.
