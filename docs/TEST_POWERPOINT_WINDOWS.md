# Mise à jour et tests PowerPoint Windows

État au 11 septembre 2026 : la mise à jour est publiée par le workflow GitHub Pages après les tests ; son résultat se vérifie dans les Actions du dépôt. Le manifeste courant est `manifest.xml` version **1.0.2.0**. Le compagnon a réussi une recherche réelle sur le Mac de développement, et l’utilisateur a confirmé le fonctionnement de l’ancienne liaison Codex sur Windows. Le nouveau compagnon en arrière-plan n’a pas encore été essayé par l’utilisateur. Le poste Windows professionnel et ses 1 740 vrais SVG n’ont pas été directement accessibles aux essais du développeur.

## Rendre cette version accessible au poste

La version déjà installée pointe vers `https://azertiv.github.io/logos-ppt/taskpane.html`. Les corrections web arrivent à son actualisation ; la correction du moteur partagé nécessite le [manifeste courant](https://azertiv.github.io/logos-ppt/manifest.xml), version **1.0.2.0**. Il retire `TaskpaneId`, incompatible avec la configuration recommandée du moteur partagé, et versionne l’URL du volet. La publication GitHub Pages a été autorisée par l’utilisateur ; aucune publication Microsoft Marketplace n’a été demandée.

Le même identifiant de complément et la même origine web sont conservés pour préserver autant que possible les préférences et le ZIP déjà stockés. Une suppression/réinstallation, une politique d’entreprise ou l’effacement du cache Office peuvent toutefois effacer le stockage local : conserver le ZIP original avant ces manipulations.

Après mise à jour, les repères visibles sont le bouton **← Retour** dans les réglages, le curseur **1 à 6** dans la barre de recherche et **gpt-5.6-luna** dans le suivi API. Le raccourci exige aussi la mise à jour du manifeste, même si l’interface web s’est déjà actualisée.

## Bibliothèque et affichage

Importer le ZIP habituel, vérifier le total **1 740**, puis fermer et rouvrir le volet. La bibliothèque doit revenir sans réimport. Tester les densités 1, 3 et 6, le défilement jusqu’en bas, une recherche précise, une recherche sans résultat, les favoris et le tri des récents. Une nouvelle recherche doit revenir aux premiers résultats. Les ombres ne doivent plus former un rectangle à la limite de la grille.

Ouvrir les réglages et descendre jusqu’aux options IA : **Retour** doit rester disponible. Tester également Échap, Tab et les flèches dans la grille. Sur une carte sélectionnée, Entrée ou Espace insère le logo ; sur son étoile, ces touches gèrent le favori.

## Recherche IA

Pour le compagnon, suivre `COMPAGNON_WINDOWS.md` sur **ce PC Windows** : l’instance ouverte sur le Mac ne remplace pas le compagnon local du PC. Coller le code de liaison, vérifier la connexion, choisir Luna si proposé et activer AI. Saisir « Supplier Questionnaire » d’un seul mouvement : l’appel démarre après deux secondes d’inactivité. Tester aussi Entrée avant les deux secondes, puis une saisie suivie d’un clic hors du champ. Entrée suivie d’un clic hors du champ ne doit pas doubler la même recherche. Une requête déjà mise en cache doit réutiliser son résultat.

Vérifier séparément le mode API avec la clé habituelle. Son compteur doit évoluer uniquement pour les appels API, pas pour Codex. Arrêter le compagnon puis rechercher en mode Codex : un message doit expliquer l’indisponibilité, la recherche locale doit rester accessible et aucun appel API de secours ne doit être déclenché. Une annulation interrompt le travail restant, mais ne garantit pas l’annulation du quota déjà consommé côté fournisseur.

## Raccourci depuis une diapositive

La documentation Microsoft indique la prise en charge PowerPoint Windows à partir de Microsoft 365 **version 2601, build 19628.20150** ; le canal Monthly Enterprise requiert **2604, build 19929.20172**. Vérifier la version réelle dans Fichier → Compte → À propos de PowerPoint. Le support effectif doit être essayé sur le poste, notamment si l’entreprise contrôle les mises à jour.

Mettre à jour le manifeste installé vers **1.0.2.0**, puis ouvrir une fois le complément dans la présentation. Dans **Réglages → Insertion**, vérifier l’état lu auprès de PowerPoint. Une version non compatible, une action absente, un conflit et une erreur de lecture doivent être distingués. Une ancienne combinaison personnalisée doit s’afficher telle quelle ; **Rétablir le raccourci** ne change que l’action Atelier Pictos, sur demande explicite. La simple ouverture des réglages ne modifie aucune association.

Sélectionner le texte « Supplier Questionnaire » dans une diapositive et appuyer sur **Ctrl + Alt + P**, ou sur la combinaison réellement affichée. Le volet doit afficher la recherche puis ajouter son premier pictogramme. Le texte source doit rester intact, y compris si l’option habituelle « Remplacer la sélection » est cochée. Répéter avec le bouton **Insérer depuis la sélection** : il utilise la même lecture Office, la même recherche et les mêmes protections, sans dépendre de l’enregistrement du raccourci.

Fermer le volet et répéter. Sur les versions compatibles, le moteur partagé reste chargé. L’option des réglages **Charger le complément à l’ouverture de cette présentation** permet ensuite de demander son chargement lors de la prochaine ouverture de ce document. Ce choix est propre à la présentation ; il ne rend pas compatible une version de PowerPoint qui ne l’est pas.

PowerPoint peut signaler un conflit avec son raccourci existant et demander quelle action utiliser. Choisir l’action Atelier Pictos pour cet essai. Si le raccourci ne se déclenche pas, vérifier le nouveau manifeste, la version d’Office et les préférences des raccourcis de compléments avant toute autre modification.

Enfin, lancer une recherche puis changer de diapositive pendant l’attente : le complément doit refuser l’insertion automatique et inviter à relancer le raccourci. Une recherche en échec, une sélection vide ou une bibliothèque absente ne doit rien insérer. Une sélection valide reste visible dans le champ de recherche même lorsque l’insertion est bloquée par une bibliothèque absente.

Le champ et son formulaire déclarent `autocomplete="off"`. Vérifier séparément les suggestions « Saved Info » dans le WebView2 réel : elles proviennent d’Edge et leur affichage ne prouve pas l’exécution de l’action du complément. Le correctif ne change pas les réglages globaux ni les données enregistrées du navigateur.

## Informations utiles si un essai échoue

Conserver le message exact, la version complète de PowerPoint et l’étape concernée. Le menu de l’icône affiche aussi l’état du compagnon. Les développeurs peuvent exécuter `bridge/diagnose.js` avec le runtime Node fourni pour créer un diagnostic explicite. Ne pas partager la clé API, le code de liaison ni le dossier de connexion Codex.

Références : [raccourcis Office](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts), [moteur partagé](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime), [chargement à l’ouverture](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/run-code-on-document-open).
