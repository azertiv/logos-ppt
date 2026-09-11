# Atelier Pictos — tester Codex sur Windows

Ce compagnon ajoute une option Codex à Atelier Pictos. Il ne remplace pas la recherche locale ni le mode avec clé API. Le paquet fourni vise Windows x64 (Intel/AMD). Il inclut Node.js 24.21.0 et Codex 0.154.0, déjà compilés par leurs éditeurs. Vous n’avez besoin ni d’installer Node.js, ni de compiler, ni d’utiliser PowerShell.

## Premier démarrage, sans PowerPoint

1. Copiez l’archive sur le poste Windows et décompressez **tout le dossier** dans un emplacement où vous pouvez écrire. Ne lancez pas les fichiers depuis l’intérieur du ZIP.
2. Ouvrez `Verifier.cmd`. Ce diagnostic vérifie le serveur local et le démarrage de Codex. Il ne lance aucune recherche IA et ne consomme pas de quota d’inférence. « chatgptConnected: false » est normal avant la première connexion.
3. Ouvrez `Demarrer.cmd`. Gardez sa fenêtre ouverte. Une page s’ouvre dans le navigateur ; sinon, copiez l’adresse affichée dans la fenêtre. Cette adresse contient un code privé de liaison, valable uniquement pendant cette session locale.
4. Dans la page, cliquez sur **Se connecter avec ChatGPT**, puis sur le lien proposé. Connectez-vous dans la page officielle et revenez au compagnon. La page vérifie automatiquement l’issue de la connexion pendant dix minutes. Vous pouvez aussi cliquer sur Actualiser.
5. Choisissez un modèle parmi ceux proposés par votre compte. Essayez « ambition » avec **Rechercher avec Codex**. Le test envoie à Codex cette recherche et une courte liste d’exemples, en deux opérations qui utilisent votre quota. Les résultats de démonstration sont des noms de pictogrammes, pas votre bibliothèque d’entreprise.
6. Pour arrêter le compagnon, faites Ctrl+C dans sa fenêtre. Fermer seulement l’onglet du navigateur ne l’arrête pas.

`Test-Recherche.cmd` permet également de vérifier une recherche réelle après connexion. Il utilise le même compte local du compagnon et effectue deux opérations Codex. Ne l’exécutez pas en même temps qu’une connexion ou une déconnexion dans la page. Le diagnostic n’ouvre pas PowerPoint et ne tente pas d’installer un complément.

Les diagnostics enregistrent `diagnostic.json` dans `%LOCALAPPDATA%\AtelierPictos` (ou le dossier AtelierPictos de votre profil si cette variable manque). Ce rapport ne contient ni adresse e-mail, ni clé API, ni code de liaison. Il indique la version, les étapes réussies et, si demandé, le résultat de la recherche de démonstration. Un nouveau diagnostic remplace le précédent.

## Dans le complément PowerPoint installé sur votre PC

Une fois **la version contenant cette fonctionnalité** installée par une voie autorisée : ouvrez les Réglages, puis Recherche IA. Choisissez **Abonnement ChatGPT via Codex**. Copiez le code de liaison depuis le compagnon, puis cliquez sur **Vérifier la connexion**. L’adresse par défaut est `http://127.0.0.1:43129`. Activez ensuite le bouton AI dans la recherche.

Le code change à chaque redémarrage du compagnon. Il reste seulement dans la session du navigateur du complément. Une connexion ChatGPT réussie, elle, est conservée par Codex dans le dossier propre au compagnon. Le code de liaison n’est pas votre mot de passe ChatGPT et ne doit pas être communiqué à d’autres personnes.

Pour utiliser GPT‑5.6 Luna via l’API, choisissez **Clé API OpenAI** : votre clé et son suivi de consommation sont conservés. Le mode Codex ne déclenche jamais de secours payant par API. En cas d’échec, la recherche locale demeure disponible. Les pourcentages affichés correspondent aux limites du compte Codex, pas à un crédit d’API ni à un quota réservé au complément.

## Ce que le poste doit autoriser

Le lancement est conçu pour un utilisateur standard : aucune installation système, aucun service Windows, aucune modification du registre, des politiques PowerShell ou du pare-feu, aucun ajout de certificat. La présence d’un exécutable de préparation de sandbox dans les fichiers officiels Codex n’entraîne pas son installation : le compagnon demande le mode Windows sans élévation.

Une entreprise peut toutefois bloquer les exécutables portables, les scripts `.cmd`, la connexion ChatGPT ou l’accès au réseau local. Si une protection bloque le lancement, conservez son message et passez par le service informatique ; ce paquet ne contourne pas ces restrictions. La publication sur Microsoft Marketplace ne garantit pas à elle seule qu’une entreprise autorisera l’installation du complément ou sa liaison locale.

Un test réussi dans le navigateur prouve le fonctionnement du compagnon sur ce poste, pas celui de la liaison depuis PowerPoint. Office peut imposer des contraintes supplémentaires pour les accès locaux. Si c’est le cas, le mode API existant reste utilisable indépendamment du compagnon. Il faudra examiner le message exact avant de modifier l’architecture.

## Données et limites

Seuls les concepts recherchés et les noms/mots-clés des candidats nécessaires sont transmis à Codex. Le complément conserve le choix et l’insertion. Le compagnon ne fournit aucune route pour lancer une commande arbitraire, lire un fichier ou modifier la présentation. Les outils de commande sont désactivés et les demandes de permissions sont refusées.

La connexion de ce compagnon est séparée de celle de l’application Codex. **Déconnecter le compte** dans la page retire cette connexion locale. Les fichiers de compte restent gérés par Codex ; ne partagez pas le dossier `%LOCALAPPDATA%\AtelierPictos` ni son sous-dossier `codex`.

La commande App Server reste expérimentale selon la documentation OpenAI. Ce paquet est une version à valider sur le poste Windows et dans Office avant toute diffusion commerciale.

Le protocole détaillé de mise à jour et de test du raccourci est fourni dans `docs/TEST_POWERPOINT_WINDOWS.md` du projet. Le nouveau manifeste seul ne publie pas la nouvelle interface web.
