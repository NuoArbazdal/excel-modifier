import io
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="Modifier un fichier Excel",
    page_icon="📄",
    layout="centered"
)

st.markdown(
    """
    <style>
        .block-container {
            max-width: 720px;
            padding-top: 4rem;
        }
        .title {
            text-align: center;
            font-size: 2rem;
            font-weight: 700;
            margin-bottom: .5rem;
        }
        .subtitle {
            text-align: center;
            color: #777;
            margin-bottom: 2rem;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="title">Dépose ton fichier Excel</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Le fichier sera traité automatiquement puis tu pourras récupérer la version modifiée.</div>',
    unsafe_allow_html=True
)

def modifier_excel(fichier_bytes: bytes) -> bytes:
    """
    Fonction temporaire.
    On remplacera son contenu plus tard par ton prompt / ta logique réelle.
    """
    entree = io.BytesIO(fichier_bytes)
    wb = load_workbook(entree)

    nom_feuille = "Feuille ajoutée"

    if nom_feuille in wb.sheetnames:
        del wb[nom_feuille]

    ws = wb.create_sheet(nom_feuille)
    ws["A1"] = "Cette feuille sera remplacée par la logique de ton prompt."

    sortie = io.BytesIO()
    wb.save(sortie)
    sortie.seek(0)

    return sortie.getvalue()

fichier = st.file_uploader(
    "Dépose ton document ici",
    type=["xlsx"],
    label_visibility="collapsed"
)

if fichier is not None:
    try:
        resultat = modifier_excel(fichier.getvalue())

        nom_sortie = fichier.name.replace(".xlsx", "_modifie.xlsx")

        st.success("Fichier modifié avec succès.")

        st.download_button(
            "Récupérer le fichier modifié",
            data=resultat,
            file_name=nom_sortie,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Impossible de traiter ce fichier : {e}")
