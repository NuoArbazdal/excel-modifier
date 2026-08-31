# Excel Modifier V3

Interface volontairement minimale :

1. Déposer un fichier `.xls` ou `.xlsx`.
2. Le site traite automatiquement le fichier.
3. Télécharger le résultat en `.xlsx`.

Les anciens fichiers `.xls` sont convertis automatiquement côté serveur avec LibreOffice.
L'utilisateur n'a aucune conversion à effectuer.

## Fichiers obligatoires à la racine du dépôt GitHub

- `app.py`
- `requirements.txt`
- `packages.txt`
- `modele_visuel.xlsx`
- `README.md`

`packages.txt` installe `libreoffice-calc` sur Streamlit Community Cloud pour permettre
la conversion automatique des anciens classeurs `.xls`.
\n## V4\nCorrection de la copie des styles entre le modèle et le fichier utilisateur.\n
## V5
Correction du remplissage bleu des barres Gantt : couleur ARGB opaque FF00A9E0.

## V6
Aucun volet figé dans la feuille Planning dynamique.

## V7
Barres Gantt visibles immédiatement et dynamiques. Le remplissage initial est écrit dans le fichier et les règles dynamiques référencent directement les durées des feuilles sources. Aucun volet figé.

## V8
Correction des formules de mise en forme conditionnelle : les barres bleues sont calculées directement depuis les durées source. 1 jour = 2 demi-cases. Une modification d'une durée redimensionne la barre et décale automatiquement les tâches suivantes.

## V9
Barres dynamiques basées sur la colonne C de Planning dynamique (même feuille). 1 jour = 2 demi-cases. Les références inter-feuilles ont été retirées des règles de couleur.

## V10
Recalcul automatique du classeur côté serveur avec LibreOffice avant téléchargement. Les valeurs des formules sont enregistrées dans le fichier, afin que les barres conditionnelles soient visibles immédiatement tout en restant dynamiques.

## V11
Détection assouplie : un titre contenant CHANTIER suffit désormais. Le mot LOT n'est plus obligatoire, ce qui permet de traiter des fichiers comme TEST 2.xls.
