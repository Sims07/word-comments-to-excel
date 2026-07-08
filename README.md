# FDL Extractor

Petit outil console (aucune installation, aucun serveur) qui transforme soit un PDF **« Liste des marques de révision et commentaires »** (export Word), soit directement un fichier **Word (.docx)**, en fichier **Excel de Fiche De Lecture (FDL)**.

## Utilisation

1. Ouvrir https://sims07.github.io/word-comments-to-excel/install.html
   
## Fichiers du projet

- `fdl-extractor.js` — le script à coller dans la console (ou à charger via le favori).
- `install.html` — génère un favori (bookmarklet) pour installer l'outil sans repasser par la console à chaque fois.
- `README.md` — ce document.

## Installation en favori (bookmarklet) — recommandé pour un usage récurrent

`install.html` génère un bouton à glisser dans la barre des favoris du navigateur, qui lance directement l'outil sur n'importe quelle page, sans repasser par la console à chaque fois.

1. Héberger `fdl-extractor.js` sur GitHub (ou ailleurs) et ajuster la constante `SCRIPT_URL` en haut du `<script>` d'`install.html` pour qu'elle pointe vers l'URL brute (raw) du fichier.
2. Ouvrir `install.html` dans le navigateur.
3. Afficher la barre des favoris (`Ctrl + Maj + B` / `Cmd + Maj + B`).
4. Glisser-déposer le bouton vert **📋 FDL Extractor** dans la barre des favoris.
5. Sur n'importe quelle page, cliquer sur ce favori pour ouvrir le panneau de l'outil.

Le favori récupère toujours la dernière version de `fdl-extractor.js` à chaque clic (via `fetch`), donc pas besoin de refaire cette manipulation après une mise à jour du script — seul le contenu du fichier source change.

## Utilisation ponctuelle (sans favori)

1. Ouvrir n'importe quelle page dans le navigateur (peu importe laquelle).
2. Ouvrir la console développeur (F12, onglet *Console*).
3. Coller le contenu de `fdl-extractor.js` et valider (Entrée).
4. Un panneau s'affiche en haut à droite avec deux boutons : **Choisir un PDF** ou **Importer un Word (.docx)**.
5. Selon le bouton choisi, l'outil extrait les commentaires soit du PDF (via `pdf.js`, éventuellement OCR), soit directement du fichier Word (via lecture structurée du `.docx`), et remplit un tableau éditable.
6. Compléter/ajuster les champs qui ne peuvent pas être déduits automatiquement (voir ci-dessous), puis cliquer **Exporter en Excel** : un fichier `<nom_du_fichier>_FDL.xlsx` est téléchargé.

Rien n'est envoyé sur un serveur : tout le traitement (lecture du PDF ou du Word, génération du xlsx) se fait dans le navigateur, avec `pdf.js`/`SheetJS` (et `Tesseract.js` en mode OCR) chargés depuis un CDN. Le mode Word n'a besoin d'aucune librairie externe dans le cas courant (voir plus bas).

## Deux modes d'entrée : PDF ou Word (.docx)

| | **Choisir un PDF** | **Importer un Word (.docx)** |
|---|---|---|
| Source | Export PDF "Liste des marques de révision et commentaires" | Fichier `.docx` original |
| Fiabilité du texte | Dépend du calque texte du PDF (risque d'encodage mal mappé selon la police, cf. section dédiée) | Fiable à 100 % : les commentaires sont des données structurées, pas de rendu de police à interpréter |
| Vitesse | Proportionnelle au nombre de pages (extraction texte) ou au temps d'OCR | Quasi instantanée quelle que soit la taille du fichier : seule l'entrée `word/comments.xml` est lue, jamais le reste du document (texte, images…) |
| Colonne "Page" | Renseignée (numéro de page issu du PDF, exact car figé au rendu) | **Estimée** (préfixe `≈`) à partir des sauts de page/section explicites du document — voir limite ci-dessous |
| Mode OCR | Disponible (case à cocher) | Non applicable (pas de calque à lire) |

Dans les deux cas, **BO/FO** et **Version** sont déduits du nom du fichier uploadé, et la **Priorité** est déduite d'un tag `#0`…`#9` dans le texte du commentaire (cf. tableau des colonnes ci-dessous).

## Format PDF attendu

Chaque commentaire est repéré par une ligne de titre du type :

```
Page 11 : Commenté [JD1]John DOE (Site A)28/11/2025 10:23
```

suivie d'une ou plusieurs lignes de texte (le commentaire), jusqu'à la ligne de titre suivante.

## Format Word (.docx) attendu

L'outil lit directement l'entrée `word/comments.xml` à l'intérieur du `.docx` (un fichier Word est une archive zip). Contrairement à une lecture via une librairie de dézippage générique, l'outil ne lit que la table centrale de l'archive puis les quelques kilo-octets de cette seule entrée : le reste du document (texte principal, images, media...) n'est jamais chargé ni décompressé, donc le temps d'extraction ne dépend quasiment pas de la taille du fichier. Il récupère, pour chaque commentaire : l'auteur (`w:author`), la date (`w:date`), et le texte (concaténation des runs `w:t` de chaque paragraphe du commentaire). Aucune installation d'extension Word n'est requise — c'est un simple fichier `.docx` exporté ou enregistré normalement.

La décompression utilise l'API native du navigateur (`DecompressionStream`), disponible sans aucun réseau sur les navigateurs récents (Chrome/Edge, Firefox 100+, Safari 16.4+). Sur un navigateur plus ancien qui ne la supporterait pas, l'outil se replie automatiquement sur une petite librairie (`pako`) chargée depuis un CDN.

## Estimation du numéro de page en mode Word (approximatif)

Word ne stocke **aucun numéro de page figé** dans le `.docx` : la pagination réelle n'est calculée par Word qu'au moment du rendu/impression (c'est justement ce que fait l'export PDF, d'où sa précision). Il est donc structurellement impossible de retrouver un numéro de page exact à partir du seul fichier source.

L'outil fournit malgré tout une **estimation** en comptant, dans `word/document.xml`, les deux seuls indices de pagination présents dans le XML :
- les sauts de page manuels (`<w:br w:type="page"/>`) ;
- les fins de section (`<w:sectPr>`), qui déclenchent presque toujours un saut de page.

Chaque commentaire se voit ainsi attribuer le numéro de la page en cours à l'endroit où il est ancré dans le document. Cette estimation est affichée avec un préfixe `≈` et une cellule surlignée en bleu clair, pour bien la distinguer d'un numéro exact.

**Limite importante** : cette méthode ne peut pas détecter les sauts de page *automatiques* dus au simple débordement du texte (un paragraphe trop long qui pousse la suite à la page suivante) — c'est le cas le plus fréquent dans un document Word classique. L'estimation sera donc fiable sur des documents structurés avec des sauts explicites (un chapitre par page/section, par exemple), mais peu ou pas fiable sur un document au fil de l'eau sans sauts manuels. Dans ce cas, mieux vaut repartir du PDF si le numéro de page exact est important pour la relecture.

## Détection du numéro de version (nom de fichier)

Que ce soit pour un PDF ou un `.docx`, le numéro de version est déduit du nom du fichier uploadé s'il contient un motif `VXX` (ex. `..._V15_1_...` → version `15`, `Rapport_V3.docx` → version `3`). Si aucun motif de ce type n'est trouvé, le champ Version reste vide (à compléter à la main).

## Colonnes générées

| Colonne | Origine |
|---|---|
| RefFDL | Numéro incrémental automatique (1, 2, 3…) |
| Relecteur | Extrait du titre (PDF) ou de l'auteur du commentaire (Word) |
| Date ouverture remarque | Extraite de l'horodatage (PDF) ou de la date du commentaire (Word) |
| Priorité | Auto-détectée si la remarque contient un tag `#0`…`#9` (ex. `#0 blocage sur...` → P0) ; sinon **manuel** (champ avec suggestions P0/P1/P2/P3). Le tag est retiré du texte de la remarque une fois converti. |
| JIRA | **Manuel** |
| BO / FO | Auto-détecté depuis le nom du fichier uploadé (contient "BO" → BO, "FO" → FO, sinon NA) — modifiable |
| Version | Auto-détecté depuis le nom du fichier uploadé s'il contient un motif `VXX` (voir ci-dessus) — sinon **manuel** |
| Page, Paragraphe | PDF : numéro de page exact (figé au rendu). Word : numéro de page **estimé** (préfixé `≈`, cellule surlignée en bleu clair), à partir des sauts de page/section du document ; le numéro de paragraphe reste à compléter manuellement dans les deux cas |
| Contenu - Si applicable | **Manuel** (liste déroulante Esthétique/Question) |
| Remarque | Texte du commentaire, extrait du PDF ou du Word |

Les colonnes réservées aux répondants (Statut, Échanges, Version de prise en compte, Date fermeture remarque, Validation relecteur) ne sont volontairement pas générées par l'outil : elles sont renseignées ensuite dans le fichier FDL final.

Le tableau est entièrement éditable avant export : possibilité de corriger une valeur extraite, supprimer une ligne (🗑) ou en ajouter une manuellement (+ Ajouter une ligne).

## Problèmes d'encodage de police (caractères type coréen/CJK)

Certains PDF générés par Word intègrent une police subsettée dont pdf.js ne parvient pas à retrouver le bon mappage Unicode ; le texte extrait affiche alors des caractères CJK/Hangul incohérents à la place d'accents ou de mots entiers. Ce n'est pas un bug de cet outil mais une limite de l'extraction de texte par calque PDF face à certaines polices.

Deux garde-fous sont intégrés :
1. **Détection automatique** : toute ligne contenant des caractères improbables en français (CJK/Hangul) est surlignée en orange dans le tableau, avec un bandeau d'avertissement en haut du panneau.
2. **Mode OCR** (case à cocher dans l'en-tête) : au lieu de lire le calque texte du PDF, l'outil rend chaque page en image et utilise l'OCR (Tesseract.js, moteur français) pour relire le texte visuellement — cela contourne le problème de police mais est plus lent (quelques secondes par page) et nécessite un accès internet vers le CDN de Tesseract.js.

Si des lignes restent surlignées même après le mode OCR, il faut les corriger manuellement (cellules éditables). Autre option si le fichier Word source est disponible : utiliser directement **Importer un Word (.docx)**, qui lit les commentaires comme des données structurées et n'est donc jamais concerné par ce problème.

## Limites connues

- Le format de la ligne de titre est reconnu par un motif fixe (`Page X : Commenté [ref] Auteur (Lieu) date`). Si Word génère un format légèrement différent (ordre des champs, absence de parenthèses, etc.), l'extraction peut échouer sur certaines lignes — dans ce cas, ajouter la ligne manuellement via « + Ajouter une ligne ».
- La date est reconnue au format `JJ/MM/AAAA` (éventuellement suivie d'une heure `HH:MM`).
- Le numéro de paragraphe et la version du document reviewé ne sont pas présents dans le PDF de commentaires et restent à saisir à la main.
- Le mode OCR dépend d'un accès réseau vers le CDN de Tesseract.js ; en environnement fermé sans accès internet externe, seul le mode texte standard fonctionnera.
- Mode Word : fonctionne sans aucun réseau sur les navigateurs récents (voir ci-dessus). Sur un navigateur ancien sans `DecompressionStream`, un accès au CDN de `pako` est nécessaire en repli. L'ordre des commentaires suit l'ordre du fichier `word/comments.xml` (généralement l'ordre de lecture, mais non garanti à 100 % selon la façon dont Word a réorganisé le document). Limite technique : les très rares archives zip "zip64" (fichiers de plusieurs Go) ne sont pas gérées.

## Design visuel

Le panneau reprend les couleurs de l'identité publique de l'Assurance Maladie (bleu institutionnel `#2A61B1`, vert santé en accent pour l'action principale). Il s'agit d'une approximation basée sur l'identité de marque publique — je n'ai pas accès à la charte graphique interne officielle (fichier de design tokens, typographie exacte, etc.). Si tu as cette charte, je peux caler les couleurs/police exactement dessus.

## Changelog

### v1.8
- Mode Word : ajout d'une **estimation** du numéro de page (préfixée `≈`), calculée à partir des sauts de page manuels et des fins de section trouvés dans `word/document.xml` — lit une deuxième entrée ciblée de l'archive (toujours sans charger le reste du document). Clairement documentée comme approximative : Word ne stocke pas de numéro de page figé, contrairement au PDF.

### v1.7
- **Correctif de performance majeur** sur le mode Word : remplacement de JSZip (qui indexait toute l'archive) par un lecteur zip minimal fait maison, qui ne lit que la table centrale de l'archive puis les octets de la seule entrée `word/comments.xml` — le reste du document (texte, images) n'est jamais touché. L'extraction passe ainsi de potentiellement plusieurs minutes sur un gros fichier à quasi instantanée, quelle que soit sa taille.
- Décompression via l'API native `DecompressionStream` du navigateur (aucun réseau nécessaire dans le cas courant), avec repli automatique sur `pako` (CDN) pour les navigateurs plus anciens.
- Message de statut plus précis pendant l'extraction Word (taille du fichier, étapes), et détection d'un éventuel blocage réseau silencieux (message d'alerte après 12s sans résultat).

### v1.6
- Nouveau mode d'entrée : **Importer un Word (.docx)**, qui lit les commentaires directement dans le fichier source (via JSZip + `word/comments.xml`) — plus fiable que le PDF puisqu'aucun rendu de police n'entre en jeu.
- **Version** désormais auto-détectée depuis le nom du fichier uploadé (motif `VXX`), pour le PDF comme pour le Word ; correction de la regex de détection pour gérer les noms de fichiers avec underscores (ex. `..._V15_1_...`), qu'elle ne reconnaissait pas auparavant.
- Factorisation de la détection BO/FO + Version dans une fonction commune aux deux modes.

### v1.5 (contribution utilisateur)
- Regex du titre PDF assouplie pour gérer les relecteurs sans lieu entre parenthèses, les auteurs et dates collés sans espace, et les références de commentaire complexes.

### v1.4
- Ajout d'`install.html` : génère un favori (bookmarklet) pour installer l'outil en un glisser-déposer, sur le modèle des autres projets (chargement dynamique du script depuis GitHub, sans copier-coller manuel dans la console).

### v1.3
- Bouton de sélection de fichier redessiné (fini le `<input type="file">` natif du navigateur) : bouton pilule assorti au reste du design, avec icône et libellé qui affiche le nom du fichier une fois choisi.

### v1.2
- Lignes zébrées + surbrillance au survol pour scanner plus facilement les longues listes.
- Séparateur visuel bleu quand le numéro de page change d'une ligne à l'autre (regroupement visuel par page).
- Largeurs de colonnes réajustées (Relecteur plus large, Priorité/JIRA/Version plus étroites, Remarque prend l'espace restant).
- Barre de statistiques (total, plage de pages, nb sans priorité, nb à vérifier) + champ de recherche/filtre rapide (page, relecteur, texte).
- Colonne RefFDL figée au défilement horizontal ; zone de texte "Remarque" s'agrandit au focus pour relire plus confortablement.

### v1.1
- Priorité auto-détectée depuis un tag `#0`…`#9` dans la remarque (tag retiré du texte une fois converti).
- Détection et surlignage des lignes probablement mal encodées (police PDF non standard).
- Mode OCR (Tesseract.js) en recours, activable par case à cocher, pour contourner les polices mal encodées.
- Refonte visuelle aux couleurs de l'Assurance Maladie (bleu/vert), coins arrondis, bandeau d'avertissement.

### v1.0
- Première version : extraction PDF (pdf.js) → tableau éditable → export Excel (SheetJS), colonnes conformes au modèle de FDL fourni.
