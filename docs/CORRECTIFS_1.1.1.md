# Correctif raccourci et saisie — 1.1.1

L’utilisateur rapporte que **Ctrl + Alt + P** ne déclenche aucune insertion sur son PowerPoint Windows **version 2607**. Cette version dépasse les minima documentés par Microsoft ; son ancienneté n’explique donc pas le problème. Le pop-up **Saved Info** dans le champ de recherche correspond au remplissage automatique d’Edge. L’association réellement enregistrée et le manifeste installé sur le poste ne sont pas directement accessibles ; la cause locale du non-déclenchement ne peut donc pas être affirmée.

## Corrections

Le manifeste **1.0.2.0** retire `TaskpaneId` de l’action `ShowTaskpane`. Microsoft précise que cet élément ne doit pas être présent dans un hôte utilisant un moteur partagé. L’URL du volet et ses fichiers principaux sont versionnés pour charger la mise à jour lors du remplacement du manifeste. L’identifiant du complément et son origine web restent identiques.

Le diagnostic ne confond plus `SharedRuntime 1.1` avec la prise en charge des raccourcis. Il vérifie `KeyboardShortcuts 1.1`, puis lit `Office.actions.getShortcuts()` : association absente, conflit, combinaison enregistrée et échec de lecture sont distingués. Une combinaison personnalisée est affichée telle quelle ; l’utilisateur peut rétablir uniquement celle d’Atelier Pictos. L’association du gestionnaire peut être retentée après `Office.onReady` si Office n’était pas disponible au chargement. Le chargement automatique du complément est présenté séparément et ne promet plus un raccourci compatible.

Le bouton **Insérer depuis la sélection** appelle la même action que le raccourci, sans exiger un moteur partagé lorsque le volet est déjà visible. La lecture conserve le texte avant le changement de focus ; les erreurs restent visibles et une lecture bloquée libère l’action après cinq secondes. Une bibliothèque absente bloque l’insertion mais n’empêche plus l’affichage de la sélection capturée.

Le champ de recherche et son formulaire déclarent `autocomplete="off"` avec un nom et un libellé de recherche explicites. Il s’agit d’une demande adressée au navigateur, pas d’une garantie contre toute politique de remplissage d’Edge. Les réglages et données du navigateur ne sont pas modifiés.

## Vérification et limites

Les **50 tests automatisés** passent, dont cinq régressions sur la prise en charge réelle, les différents états d’association, le rétablissement limité à cette action, l’utilisation sans moteur partagé et la sélection visible lorsque la bibliothèque manque. La construction du site réussit. Le diagnostic est inspecté dans un aperçu étroit avec Office simulé. Le manifeste est validé avec l’outil Microsoft.

Ces vérifications ne reproduisent pas le clavier de PowerPoint Windows ni les suggestions Edge sur le poste professionnel. La version 2607 est confirmée par l’utilisateur ; la confirmation du correctif nécessite le manifeste installé et l’état désormais affiché dans **Réglages → Insertion**. Le correctif du complément ne nécessite pas de changer de compagnon Codex.

## Références primaires vérifiées le 11 septembre 2026

- [Raccourcis Office : versions minimales et configuration](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)
- [KeyboardShortcuts 1.1 et clients compatibles](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets)
- [Moteur partagé : absence de TaskpaneId dans ShowTaskpane](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime#good-practice-avoid-multiple-task-panes)
- [Office.actions : lecture des associations, conflits et rétablissement](https://learn.microsoft.com/en-us/javascript/api/office/office.actions)
- [Remplissage automatique et données de Microsoft Edge](https://learn.microsoft.com/en-us/legal/microsoft-edge/privacy)
