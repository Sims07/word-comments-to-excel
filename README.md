# FDL Extractor — Macro VBA Word

Contrairement à l'outil navigateur (`fdl-extractor.js`), cette macro tourne **directement dans Word** et récupère le numéro de page **exact**, calculé par le moteur de mise en page de Word lui-même (`Range.Information(wdActiveEndPageNumber)`) — le même calcul que celui utilisé à l'impression. Aucune estimation, aucun export PDF intermédiaire : ça fonctionne quelle que soit la complexité de la mise en forme (tableaux, images, styles, en-têtes/pieds de page...).

**Limite à connaître** : cette macro nécessite Word **Desktop** (Windows ou Mac) avec les macros autorisées, et Excel installé sur le même poste (pour créer le classeur de sortie). Elle ne fonctionne pas dans Word en ligne (Word for the Web), qui ne supporte pas les macros VBA.

## Installation (à faire une fois)

1. Ouvrir le document Word à traiter.
2. Ouvrir l'éditeur VBA : `Alt` + `F11` (ou onglet **Développeur** > **Visual Basic**).
   - Si l'onglet **Développeur** n'est pas visible : **Fichier > Options > Personnaliser le ruban**, cocher **Développeur**.
3. Dans l'éditeur VBA, clic droit sur le document dans l'arborescence (panneau de gauche) > **Insérer > Module**.
4. Ouvrir le fichier `FDLExtractor.bas` fourni avec un éditeur de texte, copier tout son contenu, et le coller dans le module vide créé à l'étape 3.
   - Alternative : **Fichier > Importer un fichier...** dans l'éditeur VBA, puis sélectionner directement `FDLExtractor.bas`.
5. Fermer l'éditeur VBA (`Alt` + `Q` ou la croix).
6. Enregistrer le document en `.docm` (Word avec macros) si tu veux conserver la macro dans ce fichier précis, **ou** l'installer une fois pour toutes dans ton modèle global (`Normal.dotm`) pour qu'elle soit disponible dans tous les documents Word (voir plus bas).

### Rendre la macro disponible dans tous les documents (recommandé)

Pour ne pas avoir à réinstaller la macro à chaque document :
1. Dans l'éditeur VBA, avec `Normal.dotm` sélectionné dans l'arborescence (au lieu du document courant), suivre les mêmes étapes 3-4 ci-dessus.
2. Enregistrer : Word demande de confirmer la modification de `Normal.dotm` en quittant, accepter.
3. La macro `ExporterCommentairesFDL` sera alors disponible dans **tous** les documents Word ouverts sur ce poste.

### Ajouter un bouton (optionnel, plus confortable)

**Ruban** : clic droit sur le ruban > **Personnaliser le ruban** > créer un nouveau groupe dans un onglet > **Choisir les commandes dans : Macros** > ajouter `ExporterCommentairesFDL` au groupe.

**Barre d'outils Accès rapide** : clic droit sur la barre en haut > **Personnaliser la barre d'outils Accès rapide** > **Macros** > ajouter `ExporterCommentairesFDL`.

## Utilisation

1. Ouvrir le document Word contenant les commentaires à traiter.
2. **Développeur > Macros** (ou `Alt` + `F8`), sélectionner `ExporterCommentairesFDL`, cliquer **Exécuter** (ou le bouton ajouté à l'étape précédente).
3. Si Word affiche un avertissement de sécurité sur les macros, autoriser l'exécution (macro locale, pas de contenu téléchargé).
4. Excel s'ouvre automatiquement avec un nouveau classeur contenant l'onglet **Detail** rempli.

## Colonnes générées

Identiques à l'outil navigateur : `RefFDL`, `Relecteur`, `Date ouverture remarque`, `Priorité`, `JIRA`, `BO / FO`, `Version`, `Page, Paragraphe`, `Contenu - Si applicable`, `Remarque`.

- **Page, Paragraphe** : numéro de page **exact** (calculé par Word). Le numéro de paragraphe n'est pas rempli (reste manuel, comme dans l'outil navigateur).
- **Priorité** : déduite d'un tag `#0`…`#9` dans le texte du commentaire (tag retiré une fois converti), comme dans l'outil navigateur.
- **BO / FO** et **Version** : déduits du nom du fichier Word ouvert (même logique que l'outil navigateur : contient "BO"/"FO", et motif `VXX` pour la version).
- **JIRA** et **Contenu - Si applicable** : laissés vides, à compléter manuellement.

Les commentaires sont triés par page puis par position dans le document (comme dans un export PDF classique des commentaires).

## Dépannage

### "Erreur de compilation : Erreur de syntaxe" en ouvrant le module

C'est presque toujours un problème de copier-coller, pas un vrai problème dans le code :

- **Guillemets courbes** : si le code a été collé via Word (ou tout endroit avec correction automatique), les guillemets droits `"` peuvent être transformés en guillemets courbes `" "`, que VBA refuse. Solution : supprimer entièrement le module existant, puis utiliser **Fichier > Importer un fichier...** dans l'éditeur VBA pour charger directement `FDLExtractor.bas` (lecture du fichier brut, aucune transformation possible), plutôt qu'un copier-coller.
- **Ligne coupée** : si une ligne a été tronquée pendant le copier-coller (fin de ligne manquante), le symptôme est le même. Là encore, l'import direct du `.bas` évite ce risque.

### "La macro est introuvable ou a été désactivée..." en cliquant sur un bouton du ruban/QAT

Si ce message apparaît alors que la macro s'exécute correctement via `Alt+F8`, le bouton pointe vers une référence obsolète (créé avant que le module soit finalisé, ou double présence document + Normal.dotm créant une ambiguïté). Supprimer le bouton (clic droit dans **Personnaliser le ruban** > **Supprimer**) et le recréer une fois la macro confirmée fonctionnelle via `Alt+F8`.

### Repartir sur une base propre

En cas de doute, le plus sûr est de tout supprimer et recommencer proprement :
1. `Alt+F11`, supprimer le(s) module(s) `FDLExtractor` existant(s) (clic droit > **Supprimer FDLExtractor**, répondre "Non" si Word propose de l'exporter avant).
2. Clic droit sur **Normal (NormalProject)** > **Insérer > Module**.
3. **Fichier > Importer un fichier...** et sélectionner `FDLExtractor.bas` directement (pas de copier-coller).
4. Fermer l'éditeur VBA, quitter Word en acceptant d'enregistrer les modifications de `Normal.dotm`.
5. Tester via `Alt+F8` avant de recréer un bouton.

## Sécurité / macros


Cette macro ne fait que lire le document ouvert et écrire dans un nouveau classeur Excel — elle n'accède à aucune ressource réseau, ne modifie pas le document Word source, et ne contient aucun code téléchargé. Si ta politique d'entreprise bloque toutes les macros, il faudra passer par l'outil navigateur (`fdl-extractor.js`) à la place, en acceptant sa limite (page estimée, pas exacte).
