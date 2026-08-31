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
