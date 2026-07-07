# FDL Extractor

Petit outil console (aucune installation, aucun serveur) qui transforme un PDF **« Liste des marques de révision et commentaires »** (export Word) en fichier **Excel de Fiche De Lecture (FDL)**.

## Utilisation

1. Ouvrir n'importe quelle page dans le navigateur (peu importe laquelle).
2. Ouvrir la console développeur (F12, onglet *Console*).
3. Coller le contenu de `fdl-extractor.js` et valider (Entrée).
4. Un panneau s'affiche en haut à droite : cliquer sur **Choisir un fichier**, sélectionner le PDF des commentaires.
5. L'outil charge `pdf.js` et `SheetJS` depuis un CDN, extrait le texte du PDF et remplit un tableau éditable.
6. Compléter/ajuster les champs qui ne peuvent pas être déduits automatiquement du PDF (voir ci-dessous), puis cliquer **Exporter en Excel** : un fichier `<nom_du_pdf>_FDL.xlsx` est téléchargé.

Rien n'est envoyé sur un serveur : tout le traitement (lecture du PDF, génération du xlsx) se fait dans le navigateur.

## Format PDF attendu

Chaque commentaire est repéré par une ligne de titre du type :

```
Page 11 : Commenté [JD1]John DOE (Site A)28/11/2025 10:23
```

suivie d'une ou plusieurs lignes de texte (le commentaire), jusqu'à la ligne de titre suivante.

## Colonnes générées

| Colonne | Origine |
|---|---|
| RefFDL | Numéro incrémental automatique (1, 2, 3…) |
| Relecteur | Extrait du titre (nom/prénom de l'auteur) |
| Date ouverture remarque | Extraite de l'horodatage du titre |
| Priorité | Auto-détectée si la remarque contient un tag `#0`…`#9` (ex. `#0 blocage sur...` → P0) ; sinon **manuel** (champ avec suggestions P0/P1/P2/P3). Le tag est retiré du texte de la remarque une fois converti. |
| JIRA | **Manuel** |
| BO / FO | Auto-détecté depuis le nom du fichier PDF (contient "BO" → BO, "FO" → FO, sinon NA) — modifiable |
| Version | **Manuel** (non déductible du PDF) |
| Page, Paragraphe | Numéro de page pré-rempli ; le numéro de paragraphe est à compléter manuellement |
| Contenu - Si applicable | **Manuel** (liste déroulante Esthétique/Question) |
| Remarque | Texte du commentaire, extrait du PDF |

Les colonnes réservées aux répondants (Statut, Échanges, Version de prise en compte, Date fermeture remarque, Validation relecteur) ne sont volontairement pas générées par l'outil : elles sont renseignées ensuite dans le fichier FDL final.

Le tableau est entièrement éditable avant export : possibilité de corriger une valeur extraite, supprimer une ligne (🗑) ou en ajouter une manuellement (+ Ajouter une ligne).

## Problèmes d'encodage de police (caractères type coréen/CJK)

Certains PDF générés par Word intègrent une police subsettée dont pdf.js ne parvient pas à retrouver le bon mappage Unicode ; le texte extrait affiche alors des caractères CJK/Hangul incohérents à la place d'accents ou de mots entiers. Ce n'est pas un bug de cet outil mais une limite de l'extraction de texte par calque PDF face à certaines polices.

Deux garde-fous sont intégrés :
1. **Détection automatique** : toute ligne contenant des caractères improbables en français (CJK/Hangul) est surlignée en orange dans le tableau, avec un bandeau d'avertissement en haut du panneau.
2. **Mode OCR** (case à cocher dans l'en-tête) : au lieu de lire le calque texte du PDF, l'outil rend chaque page en image et utilise l'OCR (Tesseract.js, moteur français) pour relire le texte visuellement — cela contourne le problème de police mais est plus lent (quelques secondes par page) et nécessite un accès internet vers le CDN de Tesseract.js.

Si des lignes restent surlignées même après le mode OCR, il faut les corriger manuellement (cellules éditables).

## Limites connues

- Le format de la ligne de titre est reconnu par un motif fixe (`Page X : Commenté [ref] Auteur (Lieu) date`). Si Word génère un format légèrement différent (ordre des champs, absence de parenthèses, etc.), l'extraction peut échouer sur certaines lignes — dans ce cas, ajouter la ligne manuellement via « + Ajouter une ligne ».
- La date est reconnue au format `JJ/MM/AAAA` (éventuellement suivie d'une heure `HH:MM`).
- Le numéro de paragraphe et la version du document reviewé ne sont pas présents dans le PDF de commentaires et restent à saisir à la main.
- Le mode OCR dépend d'un accès réseau vers le CDN de Tesseract.js ; en environnement fermé sans accès internet externe, seul le mode texte standard fonctionnera.

## Design visuel

Le panneau reprend les couleurs de l'identité publique de l'Assurance Maladie (bleu institutionnel `#2A61B1`, vert santé en accent pour l'action principale). Il s'agit d'une approximation basée sur l'identité de marque publique — je n'ai pas accès à la charte graphique interne officielle (fichier de design tokens, typographie exacte, etc.). Si tu as cette charte, je peux caler les couleurs/police exactement dessus.

## Changelog

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
