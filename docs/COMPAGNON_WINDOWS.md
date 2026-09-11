# Neurow.Pictos — compagnon Windows

Le compagnon fonctionne discrètement près de l’horloge Windows. Il utilise votre compte ChatGPT pour les recherches Codex du complément PowerPoint. Aucun terminal ni navigateur ne s’ouvre au lancement.

## Mise en place

1. Téléchargez [l’archive Windows](https://azertiv.github.io/logos-ppt/companion/Atelier-Pictos-Codex-Windows-x64.zip) et décompressez tout le dossier dans un emplacement stable de votre profil. Par exemple : `%LOCALAPPDATA%\AtelierPictos\Application`.
2. Si l’ancienne version tourne encore dans un terminal, arrêtez-la une fois avec Ctrl+C, puis fermez cette fenêtre.
3. Lancez **Neurow.Pictos.exe**. L’icône « P » apparaît près de l’horloge, éventuellement dans le menu des icônes masquées **^**.
4. Cliquez sur l’icône et choisissez **Connexion et liaison PowerPoint…**. Si nécessaire, cliquez sur **Se connecter avec ChatGPT** et terminez la connexion sur la page officielle.
5. Depuis l’icône ou la page du compagnon, choisissez **Copier le code de liaison**. Dans PowerPoint : **Réglages → Recherche IA → Abonnement ChatGPT · Codex**. Collez le code et cliquez sur **Enregistrer / vérifier**.
6. Pour retrouver le compagnon à chaque ouverture de session, cochez **Démarrer avec Windows · cet utilisateur** dans le menu de l’icône.

Vous pouvez fermer la page du navigateur. La liaison est conservée dans votre profil PowerPoint et le code reste identique au prochain lancement du compagnon. La mise à jour du compagnon réutilise le dossier de connexion ChatGPT existant par défaut ; elle ne supprime pas votre bibliothèque PowerPoint. Le manifeste Office reste en version **1.0.1.0** pour **Ctrl + Alt + P**.

## Au quotidien

Le menu de l’icône indique si ChatGPT est connecté, à connecter, ou si le compagnon rencontre un problème. Il permet de copier le code, d’ouvrir la page de connexion, de relancer le compagnon et de le quitter. Un second lancement ne crée pas une seconde instance.

PowerPoint affiche aussi le statut sous le titre du volet. Il vérifie automatiquement la liaison lorsqu’il est ouvert et la retrouve après un redémarrage du compagnon. Ces vérifications ne lancent aucune recherche et ne consomment aucun quota d’inférence. Le choix du modèle reste dans les réglages PowerPoint.

Le démarrage automatique concerne uniquement votre utilisateur. Vous pouvez le désactiver en décochant la même option. Gardez le dossier du programme au même emplacement ; si vous le déplacez, réactivez cette option depuis son nouvel emplacement. Windows ou la politique de l’entreprise peuvent différer ou interdire ce démarrage.

## Votre code et votre compte

Le code est enregistré dans `%LOCALAPPDATA%\AtelierPictos\pairing.dat`, protégé par Windows pour votre utilisateur. Il ne s’agit pas de votre mot de passe ChatGPT. Dans le complément, il est conservé dans le stockage local du profil Office. Ne partagez ni ce code ni le dossier de connexion `%LOCALAPPDATA%\AtelierPictos\codex`.

**Options de liaison → Créer un nouveau code…**, dans le menu de l’icône, permet de renouveler volontairement la liaison. Il faudra alors recopier ce nouveau code dans PowerPoint. **Oublier la liaison**, dans les réglages PowerPoint, efface uniquement la liaison de ce profil ; **Déconnecter ChatGPT**, dans les options de la page du compagnon, déconnecte le compte du compagnon.

## Compatibilité

Le paquet vise Windows x64 avec .NET Framework 4.8 ou supérieur. Il contient le lanceur déjà compilé ainsi que Node.js et Codex : aucun compilateur, installation système ou droit administrateur n’est nécessaire pour ce fonctionnement. L’option de démarrage écrit uniquement l’entrée du compagnon dans le registre de votre utilisateur. Le paquet ne modifie ni le pare-feu ni les politiques PowerShell.

L’exécutable du lanceur n’est pas signé avec un certificat commercial. Si le poste professionnel bloque son exécution ou la liaison locale avec Office, l’autorisation relève du service informatique ; le compagnon ne contourne pas ces restrictions. Le mode API et la recherche locale du complément restent disponibles indépendamment.
