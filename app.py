import base64
import html
import os
from pathlib import Path

import streamlit as st

# Adresse du hub EtudeApp (modifiable sans toucher au code : variable d'environnement HUB_URL).
HUB_URL = os.environ.get("HUB_URL", "https://etude-app.fr")
if not HUB_URL.lower().startswith(("https://", "http://")):
    HUB_URL = "https://etude-app.fr"

ASSETS = Path(__file__).with_name("assets")

st.set_page_config(
    page_title="Planning dynamique Excel",
    page_icon="📄",
    layout="centered",
)


def _illustration() -> str:
    try:
        data = (ASSETS / "visuel-planning.png").read_bytes()
    except OSError:
        return ""
    return "data:image/png;base64," + base64.b64encode(data).decode("ascii")


CSS = """
<style>
@font-face{font-family:Inter;font-weight:400;font-display:swap;src:url(app/static/inter-400.woff) format('woff')}
@font-face{font-family:Inter;font-weight:600;font-display:swap;src:url(app/static/inter-600.woff) format('woff')}
@font-face{font-family:Inter;font-weight:700;font-display:swap;src:url(app/static/inter-700.woff) format('woff')}
@font-face{font-family:Inter;font-weight:800;font-display:swap;src:url(app/static/inter-800.woff) format('woff')}
html,body,[data-testid="stApp"],[data-testid="stApp"] *{font-family:Inter,"Segoe UI",system-ui,sans-serif}
[data-testid="stApp"]{background:#f8fbfe;color:#0b1736}
#MainMenu,footer,[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stHeader"],[data-testid="stStatusWidget"]{display:none!important}
[data-testid="stMainBlockContainer"],.block-container{max-width:680px!important;padding:128px 24px 56px!important}

/* En-tête : identique au hub */
.ea-header{position:fixed;top:0;left:0;right:0;height:80px;background:#fff;border-bottom:1px solid #dfe8f4;z-index:1000}
.ea-header-in{height:100%;width:min(1413px,calc(100% - 48px));margin:0 auto;display:flex;align-items:center;justify-content:space-between}
.ea-brand{display:flex;align-items:center;gap:20px;font-size:26px;font-weight:700;letter-spacing:-.5px;color:#0b1736;text-decoration:none}
.ea-mark{width:48px;height:48px;border-radius:12px;background:#1d63e2;color:#fff;display:grid;place-items:center;font-size:22px;font-weight:700;letter-spacing:0}
.ea-hub{display:flex;align-items:center;gap:10px;background:#fff;color:#0b1736;border:1px solid #dfe8f4;border-radius:12px;padding:13px 22px;font-size:17px;font-weight:700;text-decoration:none;box-shadow:0 2px 6px #1e468c0d;transition:background .15s,border-color .15s}
.ea-hub:hover{background:#f3f7fd;border-color:#c9daf2;color:#0b5be0}
.ea-hub svg{width:20px;height:20px;color:#1d63e2}

/* Carte d'introduction : même construction que les cartes du hub */
.ea-card{container-type:inline-size;background:#fff;border:1px solid #dde8f6;border-radius:18px;overflow:hidden;box-shadow:0 14px 30px #285aa012;margin-bottom:28px}
.ea-hero{aspect-ratio:453/174;background:#eef5fd;overflow:hidden}
.ea-hero img{display:block;width:101.8cqw;max-width:none;margin-left:-1.8cqw;height:100%;object-fit:cover;object-position:left bottom}
.ea-card-body{padding:0 6.4cqw 6.4cqw}
.ea-tile{width:18.3cqw;height:18.1cqw;margin:-11.9cqw 0 0;border:.45cqw solid #f9fbff;border-radius:4.9cqw;background:linear-gradient(#f7fafe,#f0f6fd);box-shadow:0 1.3cqw 3.5cqw #285aaa1a;color:#1d63e2;position:relative;display:grid;place-items:center}
.ea-tile svg{width:8.4cqw;height:8.4cqw;display:block}
.ea-card h1{font-size:34px;font-weight:800;letter-spacing:-.035em;line-height:1.15;margin:26px 0 0;padding:0;color:#0b1736}
.ea-card p{font-size:17px;line-height:1.5;color:#5f7090;margin:10px 0 0}

/* Choix du type de planning */
[data-testid="stRadio"]>label,[data-testid="stRadio"]>div:first-child p{font-size:14px;font-weight:600;color:#5f7090}
[data-testid="stRadio"] [role="radiogroup"]{gap:12px;flex-wrap:wrap}
[data-testid="stRadio"] label[data-baseweb="radio"]{background:#fff;border:1.5px solid #dde8f6;border-radius:14px;padding:12px 18px;margin:0;transition:border-color .15s,background .15s}
[data-testid="stRadio"] label[data-baseweb="radio"]:hover{border-color:#b9d2f7}
[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked){border-color:#0b5be0;background:#eef5fd}
[data-testid="stRadio"] label[data-baseweb="radio"] p{font-size:16px;font-weight:600;color:#0b1736}

/* Zone de dépôt */
[data-testid="stFileUploaderDropzone"]{background:#fff;border:2px dashed #b9d2f7;border-radius:18px;padding:34px 24px;transition:border-color .15s,background .15s}
[data-testid="stFileUploaderDropzone"]:hover{border-color:#0b5be0;background:#f6faff}
[data-testid="stFileUploaderDropzoneInstructions"]>div>span{font-size:0!important}
[data-testid="stFileUploaderDropzoneInstructions"]>div>span::before{content:"Dépose ton fichier Excel ici";font-size:18px;font-weight:700;color:#0b1736}
[data-testid="stFileUploaderDropzoneInstructions"] small{font-size:0!important}
[data-testid="stFileUploaderDropzoneInstructions"] small::before{content:"Fichiers .xls ou .xlsx";font-size:14px;color:#5f7090}
[data-testid="stFileUploaderDropzone"] button{font-size:0!important;background:#fff;border:1px solid #dfe8f4;border-radius:12px;padding:10px 20px;box-shadow:0 2px 6px #1e468c0d}
[data-testid="stFileUploaderDropzone"] button::after{content:"Parcourir";font-size:15px;font-weight:700;color:#0b1736}
[data-testid="stFileUploaderDropzone"] button:hover{border-color:#0b5be0}
[data-testid="stFileUploaderDropzone"] button:hover::after{color:#0b5be0}

/* Résultat */
[data-testid="stAlert"]{border-radius:14px;border:1px solid #dde8f6;font-size:16px}
[data-testid="stDownloadButton"] button{background:#0b5be0;color:#fff;border:0;border-radius:14px;padding:16px 24px;font-size:17px;font-weight:700;box-shadow:0 10px 22px #0b5be033;transition:background .15s}
[data-testid="stDownloadButton"] button p{color:#fff;font-size:17px;font-weight:700}
[data-testid="stDownloadButton"] button:hover{background:#0a4fc4;color:#fff}

@media(max-width:600px){
 .ea-brand{font-size:21px;gap:12px}.ea-mark{width:38px;height:38px;font-size:18px}
 .ea-hub{padding:10px 14px;font-size:14px}
 [data-testid="stMainBlockContainer"],.block-container{padding:104px 16px 40px!important}
 .ea-card h1{font-size:26px}.ea-card p{font-size:15px}
}
</style>
"""

HUB_ICON = (
    '<svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">'
    '<rect x="3" y="3" width="7" height="7" rx="1.6"/><rect x="14" y="3" width="7" height="7" rx="1.6"/>'
    '<rect x="3" y="14" width="7" height="7" rx="1.6"/><rect x="14" y="14" width="7" height="7" rx="1.6"/></svg>'
)

CALENDAR_ICON = (
    '<svg viewBox="0 0 40 40" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" '
    'stroke-linejoin="round" aria-hidden="true"><path d="M12 3v7M28 3v7M9 7h22a5 5 0 0 1 5 5v21a5 5 0 0 1-5 5H9'
    'a5 5 0 0 1-5-5V12a5 5 0 0 1 5-5ZM4 17h32M13 24h.01M20 24h.01M27 24h.01M13 31h.01M20 31h.01"/></svg>'
)


def header_html() -> str:
    hub = html.escape(HUB_URL, quote=True)
    return (
        '<div class="ea-header"><div class="ea-header-in">'
        f'<a class="ea-brand" href="{hub}" target="_top" title="Retour au hub"><span class="ea-mark">E</span><span>EtudeApp</span></a>'
        f'<a class="ea-hub" href="{hub}" target="_top">{HUB_ICON}<span>Hub</span></a>'
        "</div></div>"
    )


def intro_html() -> str:
    image = _illustration()
    hero = f'<img src="{image}" alt="">' if image else ""
    return (
        '<div class="ea-card">'
        f'<div class="ea-hero">{hero}</div>'
        '<div class="ea-card-body">'
        f'<div class="ea-tile">{CALENDAR_ICON}</div>'
        "<h1>Planning dynamique</h1>"
        "<p>Choisis le type de planning, puis dépose ton fichier Excel.</p>"
        "</div></div>"
    )


st.markdown(CSS, unsafe_allow_html=True)
st.markdown(header_html(), unsafe_allow_html=True)
st.markdown(intro_html(), unsafe_allow_html=True)

mode = st.radio(
    "Type de planning",
    ["Planning classique", "Planning avec phases"],
    horizontal=True,
)

uploaded = st.file_uploader(
    "Dépose ton document ici",
    type=["xls", "xlsx"],
    label_visibility="collapsed",
)

if uploaded is not None:
    try:
        if mode == "Planning avec phases":
            from planning_phases import process_file
        else:
            from planning_classique import process_file

        with st.spinner("Création du planning dynamique..."):
            result = process_file(uploaded.getvalue(), uploaded.name)

        lower_name = uploaded.name.lower()
        if lower_name.endswith(".xlsx"):
            output_name = uploaded.name[:-5] + "_planning.xlsx"
        elif lower_name.endswith(".xls"):
            output_name = uploaded.name[:-4] + "_planning.xlsx"
        else:
            output_name = "fichier_planning.xlsx"

        st.success("Planning créé avec succès.")
        st.download_button(
            "Récupérer le fichier modifié",
            data=result,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as exc:
        st.error(str(exc))
