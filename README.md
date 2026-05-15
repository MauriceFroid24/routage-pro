# Routage PRO V11 — terrain Froid24

Version terrain optimisée iPhone / Windows Surface.

## Nouveautés V11

- Ordre strict par heure de RDV, pas d’optimisation automatique.
- Durée RDV par défaut : 2h30.
- Adresse départ / retour par défaut : 72 avenue des Tourelles, 94490 Ormesson-sur-Marne.
- Fil conducteur terrain : départ base, arrivée RDV, fin prévue, départ max vers RDV suivant, pauses disponibles.
- Détail des trajets étape par étape avec durées au format heure/minute.
- Retour base calculé automatiquement.
- Carte plus lisible : numéro RDV + nom + heure, sans téléphone.
- Tracé des trajets par la route via OSRM quand disponible, plus de simple ligne droite.
- Boutons Waze, Google Maps, Voir maison, Appeler.
- Lien Voir maison corrigé pour éviter l’écran noir.
- PDF enrichi cliquable + export CSV de sauvegarde.

## Format Excel attendu

A numéro RDV · B adresse · C code postal · D date RDV · E heure RDV · J/N nom/prénom · Q téléphone · R ville.


## V11
- Correction du bug `numpy UFuncNoLoopError` au total km.
- Rechargement automatique du dernier Excel importé pendant la journée/session Streamlit.
- Si Streamlit redémarre complètement, réimporte simplement le fichier ou le CSV/PDF sauvegardé.
