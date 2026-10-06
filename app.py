import streamlit as st

st.set_page_config(
    page_title="Planning dynamique Excel",
    page_icon="📄",
    layout="centered",
)

st.markdown(
    """
    <style>
      .block-container {max-width: 760px; padding-top: 4rem;}
      .main-title {text-align:center; font-size:2rem; font-weight:700; margin-bottom:.45rem;}
      .sub {text-align:center; color:#777; margin-bottom:2rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">Dépose ton fichier Excel</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub">Choisis le type de planning, puis dépose ton fichier.</div>',
    unsafe_allow_html=True,
)

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
