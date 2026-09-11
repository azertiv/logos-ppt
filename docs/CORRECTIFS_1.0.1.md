# Correctifs 1.0.1 — 11 septembre 2026

Le retour utilisateur confirme le fonctionnement de la liaison Codex sur le poste Windows, mais signale des défauts de défilement et des réglages encombrés. Cette mise à jour corrige ces deux surfaces et remplace le raccourci Windows par **Ctrl + Alt + P**. Le manifeste passe à **1.0.1.0**, avec une URL versionnée pour la définition des raccourcis. Le compagnon conserve son fonctionnement actuel.

La page occupe maintenant exactement la hauteur du volet. Le titre et la recherche sont dans deux zones fixes ; seule la bibliothèque défile verticalement entre les deux. Le fond est uniforme, les effets de flou des barres ont été retirés et le défilement horizontal est désactivé. Le calcul des cartes utilise la hauteur de la bibliothèque, conserve environ un écran supplémentaire au-dessus et au-dessous et évite de modifier les positions tant que les lignes nécessaires restent identiques. L’arrondi des dimensions ne dépasse plus la largeur disponible. À cinq ou six colonnes, les favoris apparaissent au survol ou au focus pour laisser les petits pictogrammes lisibles.

Les réglages sont organisés en Bibliothèque, Recherche IA, Insertion et Affichage. Les explications sont accessibles par les boutons « i » ; le quota, l’adresse du compagnon et la consommation API sont dans des rubriques repliables. Le titre et la flèche de retour restent fixes. Le filtre de mots-clés est devenu une liste explicite.

## Vérification

Les **33 tests automatiques passent**, y compris les protections existantes du compagnon, les recherches, le raccourci et quatre régressions nouvelles sur le défilement et la largeur des cartes. Le manifeste passe le validateur Microsoft, et la construction du site réussit.

Avec une bibliothèque de **1 740 SVG synthétiques** dans le navigateur, la page reste à 320 × 760 pixels après un défilement de 7 600 pixels dans la bibliothèque : le titre reste à y = 0, la recherche à y = 650 et la largeur défilante reste à 320 pixels. La vérification à 240 pixels et six colonnes ne montre pas de débordement horizontal. Les aides, leur fermeture par Échap, le choix API/Codex, le filtre de mots-clés, le retour à la bibliothèque et l’accès au dernier pictogramme avec la touche Fin ont également été exercés.

Ces vérifications couvrent le code et un navigateur sur le Mac de développement avec Office simulé. Elles ne constituent pas une mesure du rendu WebView2 sur le poste Windows professionnel ni sur les vrais fichiers SVG de l’entreprise.
