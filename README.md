# Neurow.Pictos — Atelier Pictos

Complément PowerPoint pour rechercher des pictogrammes SVG dans une bibliothèque ZIP locale, les classer par mots-clés ou par IA, puis les insérer dans une diapositive.

La recherche locale fonctionne sans service IA. La recherche assistée propose au choix la clé API OpenAI existante (GPT‑5.6 Luna), ou un compagnon local utilisant une connexion ChatGPT et le quota Codex. Le choix est explicite : aucun basculement automatique de Codex vers l’API payante.

Le ZIP reste dans le stockage local du complément. Seuls la recherche et les noms/mots-clés des candidats nécessaires sont envoyés au fournisseur IA choisi, jamais les fichiers SVG ni la présentation entière.

## Utilisation et validation

- [Installer et utiliser le compagnon Windows](docs/COMPAGNON_WINDOWS.md)
- [Tester la mise à jour dans PowerPoint Windows](docs/TEST_POWERPOINT_WINDOWS.md)
- [Revue, corrections et limites de validation](docs/AUDIT_2026-09-11.md)

Un envoi sur `main` déclenche les tests puis la publication GitHub Pages. Modifier les fichiers uniquement en local ne met pas à jour l’extension déjà installée. Le site fournit aussi `manifest.xml` pour installer la configuration des raccourcis, et `deployment.json` pour identifier la révision publiée.

## Développement

`npm test` exécute les tests sans clé ni appel IA réel. `npm run build:pages` prépare `dist/`. `npm start` démarre le serveur HTTPS de développement existant après installation des dépendances et configuration des certificats de développement. Le compagnon utilise uniquement les modules standard de Node : `npm run companion`. Il accepte `PICTOS_CODEX_PATH` et `PICTOS_DATA_DIR` pour une instance de développement séparée.

`npm run package:windows` fabrique l’archive portable Windows x64 à partir des archives officielles préalablement téléchargées dans `.tmp/downloads`, vérifiées selon `bridge/runtime-lock.json`. Node et Codex sont inclus, ainsi que leurs licences. Aucun compilateur ni installation système n’est nécessaire sur le poste destinataire ; la politique de l’entreprise doit autoriser l’exécution des fichiers.

La grille ne monte que les lignes visibles et deux lignes de marge, avec navigation par flèches, Début et Fin. Le lecteur ZIP fonctionne dans un worker ; JSZip 3.10.1 est distribué avec le site et sa licence MIT. Le repli sur le fil principal permet de continuer si les workers sont indisponibles.
