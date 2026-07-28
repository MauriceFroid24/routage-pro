# Routage PRO V24.1 — 28/07/2026

Version web/iPhone basée sur la V24 installée. Carte générale en tête, résumé de journée, temps OSRM non gonflés, trafic Google à l’heure du trajet si clé API configurée, analyse IA plus lisible, ouverture PDF dans un nouvel onglet.

# Routage PRO V24 — Import CRM assisté expérimental

Base : V23 stable.

Ajouts :
- l'application accepte un Excel enrichi avec une colonne `details_crm` ;
- les infos CRM collées/récupérées alimentent directement la préparation IA ;
- un robot local expérimental `robot_crm_froid24.py` peut tenter de télécharger l'Excel depuis le CRM Alltoo/FROID24.

## Mise à jour Streamlit
Uploader dans GitHub :
- `app.py`
- `requirements.txt`
- `README.md`

## Robot local Surface
À lancer sur ta Surface, pas sur Streamlit Cloud :

1. Dézipper le dossier.
2. Double-cliquer sur `LANCER_ROBOT_CRM.bat`.
3. Saisir identifiant, mot de passe et date.
4. Le robot tente de télécharger l'Excel dans `exports_crm/`.
5. Importer ensuite cet Excel dans l'application.

Important : c'est expérimental. Le CRM peut changer ses boutons/menus, donc le robot pourra nécessiter des ajustements.
