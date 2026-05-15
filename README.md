# Routage PRO V13 — terrain iPhone / Surface

## Nouveautés V13

- Mode sombre pour économiser la batterie.
- Calcul automatique dès l'import Excel : aucun bouton à cliquer.
- Géocodage renforcé avec l'API officielle adresse.data.gouv.fr, puis fallback Nominatim.
- Tracé des routes réelles via OSRM gratuit quand les coordonnées sont disponibles.
- Pas de lignes droites de secours : si une route n'est pas disponible, elle n'est pas dessinée.
- Dernier Excel rechargé automatiquement pendant la session/journée.
- Ordre strict par date + heure de RDV.
- RDV par défaut : 150 minutes.
- Retour base inclus.
- PDF enrichi + liens Waze / Google Maps / Voir maison / Appel.

## Format Excel attendu

A : numéro de RDV  
B : adresse_du_prospect  
C : code_postal_du_prospect  
D : date_rendez_vous  
E : debut  
Q : telephone_du_prospect  
R : ville_du_prospect  

Les autres colonnes sont récupérées quand disponibles.
