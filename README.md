# claudito

Application web légère (HTML/CSS/JS vanilla, sans dépendance) pour saisir rapidement des ventes et des achats mensuels.

## Fonctionnalités

- Saisie rapide par mois : jour + montant, navigation clavier (Entrée, flèches haut/bas).
- Multiplicateur optionnel appliqué au montant saisi (ex : `5.5` → `5 500`).
- Persistance locale (`localStorage`) et journal d'audit de toutes les opérations.
- Export/import CSV, avec détection des doublons à l'import.
- Copie au format Excel (TSV) avec mapping de colonnes configurable, y compris des colonnes à valeur constante.

## Utilisation

Ouvrir `index.html` dans un navigateur, ou servir le dossier avec un serveur statique (`npx serve .`).
