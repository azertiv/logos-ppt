# Neurow.Pictos — Atelier Pictos

Complément PowerPoint pour rechercher des pictogrammes SVG dans une bibliothèque ZIP locale, les classer par mots-clés ou par IA, puis les insérer dans une diapositive.

La recherche locale fonctionne sans service IA. La recherche assistée propose au choix la clé API OpenAI existante (GPT‑5.6 Luna), ou un compagnon local utilisant une connexion ChatGPT et le quota Codex. Le choix est explicite : aucun basculement automatique de Codex vers l’API payante.

Le [compagnon Windows prêt à lancer](https://azertiv.github.io/logos-ppt/companion/Atelier-Pictos-Codex-Windows-x64.zip) fonctionne dans la zone de notification, avec démarrage facultatif à l’ouverture de session pour l’utilisateur courant. Son code de liaison est conservé et protégé par Windows. PowerPoint mémorise la liaison, se reconnecte automatiquement et affiche son statut sous le titre.

Le ZIP reste dans le stockage local du complément. Seuls la recherche et les noms/mots-clés des candidats nécessaires sont envoyés au fournisseur IA choisi, jamais les fichiers SVG ni la présentation entière.

## Utilisation et validation

- [Installer et utiliser le compagnon Windows](docs/COMPAGNON_WINDOWS.md)
- [Tester la mise à jour dans PowerPoint Windows](docs/TEST_POWERPOINT_WINDOWS.md)
- [Revue, corrections et limites de validation](docs/AUDIT_2026-09-11.md)
- [Correctifs du défilement, des réglages et du raccourci](docs/CORRECTIFS_1.0.1.md)
- [Fonctionnement et vérification du compagnon 1.1.0](docs/COMPAGNON_1.1.0.md)

Un envoi sur `main` déclenche les tests puis la publication GitHub Pages. Modifier les fichiers uniquement en local ne met pas à jour l’extension déjà installée. Le site fournit aussi `manifest.xml` pour installer la configuration des raccourcis, et `deployment.json` pour identifier la révision publiée.

## Développement

`npm test` exécute les tests sans clé ni appel IA réel. `npm run build:pages` prépare `dist/`. `npm start` démarre le serveur HTTPS de développement existant après installation des dépendances et configuration des certificats de développement. Le compagnon utilise uniquement les modules standard de Node : `npm run companion`. Il accepte `PICTOS_CODEX_PATH` et `PICTOS_DATA_DIR` pour une instance de développement séparée.

Pour fabriquer l’archive : `npm run build:windows` compile le lanceur .NET Framework 4.8, `python3 scripts/fetch-companion-runtimes.py` récupère les fichiers officiels vérifiés selon `bridge/runtime-lock.json`, puis `npm run package:windows` assemble le paquet. Node et Codex sont inclus avec leurs licences. Le poste destinataire n’a besoin d’aucun compilateur. GitHub Pages ne publie le paquet qu’après compilation et vérifications sur Windows.

Seule la bibliothèque défile verticalement, entre le titre et la recherche fixes. La grille monte les lignes visibles et environ un écran de marge de chaque côté, avec navigation par flèches, Début et Fin. Les réglages regroupent Bibliothèque, Recherche IA, Insertion et Affichage, avec des aides « i ». Le raccourci Windows est **Ctrl + Alt + P** ; sa mise à jour nécessite le manifeste **1.0.1.0**.

Le lecteur ZIP fonctionne dans un worker ; JSZip 3.10.1 est distribué avec le site et sa licence MIT. Le repli sur le fil principal permet de continuer si les workers sont indisponibles.
