# Excel Modifier V2

Cette version applique automatiquement les règles des deux prompts.

## Fonctionnement

1. L'utilisateur dépose un fichier `.xlsx`.
2. Le site détecte les blocs contenant :
   - un titre CHANTIER / LOT ;
   - une colonne de durée en jours ;
   - les catégories et les désignations.
3. Aucune feuille existante n'est modifiée.
4. Une nouvelle feuille `Planning dynamique` est créée.
5. Si elle existe déjà, le site crée `Planning dynamique 2`, puis 3, etc.
6. Les textes et durées du planning sont liés par des formules aux cellules sources.
7. Les barres utilisent des demi-journées en interne et se recalculent avec les durées.
8. `modele_visuel.xlsx` sert uniquement de référence de mise en forme.

## Fichiers à mettre sur GitHub

Les 4 fichiers suivants doivent être à la racine du dépôt :

- `app.py`
- `requirements.txt`
- `modele_visuel.xlsx`
- `README.md`

Streamlit redéploiera automatiquement après le commit.
