# Compagnon 1.1.0 — fonctionnement discret et liaison persistante

Le lanceur Windows est un exécutable .NET Framework 4.8 de type Windows GUI, sans formulaire visible. Il affiche une icône dans la zone de notification. Le menu propose le statut ChatGPT, l’ouverture volontaire de la page de liaison, la copie du code, le démarrage à l’ouverture de session, la relance et l’arrêt. Le navigateur ne s’ouvre pas au lancement.

Le démarrage automatique est désactivé initialement. Son activation écrit uniquement la valeur `Neurow.Pictos` dans `HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run`. Le chemin est cité et contrôlé ; la désactivation retire uniquement cette valeur. Un mutex par utilisateur et session évite les instances concurrentes. Le manifeste du lanceur demande `asInvoker`.

Le code aléatoire de 256 bits est conservé dans `pairing.dat`, sous le dossier de données du compagnon. Windows DPAPI le protège avec la portée `CurrentUser`. La création et le renouvellement utilisent une écriture atomique. Un fichier existant mais illisible n’est jamais remplacé silencieusement. Le code est transmis à Node par son entrée standard privée, sans argument de commande ni variable d’environnement. La page web s’ouvre avec un ticket à usage unique valable 60 secondes ; le code permanent n’apparaît pas dans les arguments du navigateur.

Un Job Object Windows possède Node et ses descendants Codex. La fermeture du lanceur arrête également les descendants, même lors d’un arrêt brutal. Le lancement habituel reste sans console ; les relances automatiques après un échec sont limitées. Une ancienne version occupant déjà le port n’est pas tuée : son arrêt initial est expliqué dans le guide.

PowerPoint conserve la liaison dans le stockage local de son profil, vérifie automatiquement la connexion et distingue un code incorrect, un compagnon indisponible et un compte ChatGPT à connecter. Un échec de stockage est signalé. Le contrôle périodique ne relance pas de recherche et n’annule pas une recherche en cours. Il utilise une lecture du compte, sans récupérer les modèles ni les quotas à chaque passage. La compatibilité avec l’ancien compagnon en console est conservée.

La page web se limite à la connexion ChatGPT, à l’actualisation et à la copie du code. La déconnexion et la connexion par code sont repliées. Les exemples de recherche, choix du modèle et panneaux de consommation ont été retirés de cette page.

## Vérification et publication

Les 45 tests JavaScript couvrent notamment deux vrais lancements du serveur géré avec le même code, son arrêt à la fermeture du canal parent, les tickets à usage unique et leur expiration, la persistance côté complément, la reconnexion et l’absence d’inférence lors des vérifications de statut. Les interfaces ont été exercées dans le navigateur avec une bibliothèque synthétique et un compagnon simulé.

Le workflow `windows-companion.yml` compile le lanceur, exécute quatre vérifications natives sur Windows (DPAPI, entrée de démarrage utilisateur, lancement caché, arrêt des descendants), puis les tests JavaScript. Après assemblage du ZIP, il lance deux fois l’application Windows complète dans un profil de test isolé, vérifie l’absence de fenêtre principale, une lecture du compte par le vrai runtime Codex sans connexion ni inférence, la conservation du fichier protégé et l’arrêt du serveur. GitHub Pages dépend du succès de ces étapes avant de publier l’archive.

Ces vérifications ne remplacent pas l’observation de l’icône, de PowerPoint WebView2 et des autorisations d’exécution sur le poste professionnel. Le lanceur n’a pas de signature commerciale. Le manifeste Office reste `1.0.1.0`, car les commandes Office n’ont pas changé.

## Références

Les mécanismes Windows suivent la documentation Microsoft : [icône de notification](https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/notifyicon-component-overview-windows-forms), [démarrage à l’ouverture de session](https://learn.microsoft.com/en-us/windows/win32/setupapi/run-and-runonce-registry-keys), [protection liée à l’utilisateur](https://learn.microsoft.com/en-us/dotnet/api/system.security.cryptography.dataprotectionscope), [versions de .NET Framework intégrées à Windows](https://learn.microsoft.com/en-us/dotnet/framework/install/versions-and-dependencies).
