# Site Excel basique

Fonctionnement :

- tu déposes un fichier `.xlsx`
- le traitement démarre automatiquement
- le fichier modifié apparaît directement
- tu le récupères avec le bouton de téléchargement

Pour lancer le site :

```bash
pip install -r requirements.txt
streamlit run app.py
```

La fonction à modifier plus tard est :

```python
def modifier_excel(fichier_bytes: bytes) -> bytes:
```

C'est là qu'on intégrera ta logique finale ou ton prompt.
