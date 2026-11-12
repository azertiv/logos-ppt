# Compagnon 1.2.0 — notes de réunion pour Syllens

Le compagnon sert aussi l’extension navigateur **Syllens**, qui transcrit les sous-titres de Google Meet et Microsoft Teams. Il utilise la même connexion ChatGPT, le même quota Codex et le même code de liaison que PowerPoint, sans clé API.

## Ce qui est ajouté

- `POST /v1/meeting/summarize` — synthèse glissante. Le corps contient les métadonnées de la réunion, un contexte facultatif (ordre du jour, projet), la synthèse et les points clés précédents, et **uniquement les nouvelles lignes** de transcription. La réponse est la synthèse et les points clés à jour, avec des identifiants stables.
- `POST /v1/meeting/expand` — un point clé et la transcription ; la réponse contient 2 à 5 sous-points.
- `GET /health` annonce `capabilities: ["pictos-search", "meeting-notes"]` et la version `1.2.0`. L’extension s’en sert pour signaler un compagnon trop ancien.

Le contrat est dans `bridge/meeting-tasks.js` : schémas de sortie stricts, validation des entrées (transcription ≤ 120 000 caractères, contexte ≤ 4 000, etc.) et des réponses, consignes contre l’injection (la transcription est une donnée, jamais une instruction).

## Ce qui ne change pas

- Même isolement que la recherche de pictogrammes : fil Codex **éphémère**, bac à sable en lecture seule, aucun outil, aucun accès web, aucun fichier, historique non conservé.
- Un seul appel Codex à la fois, partagé entre PowerPoint et Syllens. Une demande concurrente reçoit `409 BUSY` ; l’extension réessaie d’elle-même.
- Code de liaison obligatoire sur `/v1/*`. Parmi les extensions de navigateur, seule l’origine de Syllens est acceptée (`chrome-extension://ncoilhmgdphmchbcdlajbjnfkbmbfogk`, identifiant fixé par la clé de son manifeste). `PICTOS_ALLOWED_ORIGINS` permet d’en ajouter une autre, par exemple après publication sur le Chrome Web Store.
- Limites : 128 Ko par demande pour la recherche, 1 Mo pour les routes de réunion. Délai : 90 s pour la recherche, 180 s pour une synthèse.

## Vérification

`npm test` couvre les deux routes avec le faux Codex : capacités annoncées, origine et code de liaison, fil éphémère en lecture seule, schéma imposé, conservation des identifiants d’une synthèse à la suivante, développement d’un point clé, limites de taille, réponse invalide refusée et partage de l’unique créneau Codex avec la recherche.
