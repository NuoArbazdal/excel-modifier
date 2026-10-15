import streamlit as st

st.set_page_config(
    page_title="Planning dynamique | EtudeApp",
    page_icon="📊",
    layout="centered",
)

st.markdown("""
<style>
:root { --blue:#1762e9; --navy:#101e38; --line:#dce5f1; }
.stApp { background:#f7f9fc; color:var(--navy); }
header[data-testid="stHeader"] { background:transparent; }
.block-container { max-width:920px; padding-top:2.3rem; padding-bottom:4rem; }
#MainMenu, footer { visibility:hidden; }
.hub-top { display:flex; align-items:center; gap:12px; padding:12px 0 19px; border-bottom:1px solid var(--line); font-weight:750; font-size:19px; }
.hub-mark { width:35px;height:35px;display:inline-flex;align-items:center;justify-content:center;border-radius:9px;background:#1762e9;color:white;font-size:16px; }
.hub-back { display:inline-block; margin:20px 0 18px; color:#1762e9; text-decoration:none; font-weight:600; font-size:14px; }
.hub-title { font-size:32px; font-weight:800;letter-spacing:-.8px; margin:0 0 4px; }
.hub-sub { color:#62718a; font-size:15px; margin-bottom:25px; }
.hub-steps { display:flex;justify-content:center;align-items:center;gap:22px; flex-wrap:wrap; color:#77859b; font-size:13px; font-weight:600; padding:20px 5px; background:#fff; border:1px solid var(--line);border-radius:14px 14px 0 0; }
.hub-steps b {color:#1762e9;}
[data-testid="stRadio"] { background:white;border-left:1px solid var(--line);border-right:1px solid var(--line);padding:12px 25px 0; }
[data-testid="stFileUploader"] { background:white;border:1px solid var(--line);border-top:0;border-radius:0 0 14px 14px;padding:18px 24px 24px; }
[data-testid="stFileUploaderDropzone"] {min-height:230px;display:flex;align-items:center;justify-content:center;text-align:center;border:1.5px dashed #bad0f4;border-radius:12px;background:#fafcff;}
[data-testid="stFileUploaderDropzone"] > div { width:100%;justify-content:center; }
[data-testid="stFileUploaderDropzone"] button {background:#fff!important;color:#1762e9!important;border:1px solid #b8d0fa!important;border-radius:9px!important;}
[data-testid="stFileUploaderFile"] {justify-content:center;}
[data-testid="stAlert"] {border-radius:12px;}
.stDownloadButton > button {background:#eef4ff!important;color:#1762e9!important;border:1px solid #c9dafb!important;border-radius:10px!important;box-shadow:none!important;min-height:48px;font-weight:650!important;}
.stDownloadButton > button:hover {background:#e2edff!important;border-color:#a8c6fa!important;}
@media(max-width:640px) {.block-container{padding-top:1.5rem}.hub-title{font-size:27px}.hub-steps{gap:10px;font-size:11px}[data-testid="stFileUploader"]{padding:12px}}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="hub-top"><span class="hub-mark">E</span>EtudeApp</div>', unsafe_allow_html=True)
st.markdown('<a class="hub-back" href="javascript:history.back()">← Retour au hub</a>', unsafe_allow_html=True)
st.markdown('<div class="hub-title">Planning dynamique</div><div class="hub-sub">Transformez votre fichier Excel en planning dynamique, automatiquement.</div>', unsafe_allow_html=True)
st.markdown('<div class="hub-steps"><b>① Déposer le fichier</b><span>② Traitement automatique</span><span>③ Télécharger le résultat</span></div>', unsafe_allow_html=True)

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

        st.success(f"Votre fichier est prêt : {output_name}")
        st.download_button(
            "Récupérer le fichier modifié",
            data=result,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as exc:
        st.error(str(exc))
