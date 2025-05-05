
import streamlit as st
from datetime import datetime
import tempfile
import os
import shutil
from rapport_generator import  (
    COMMERCIAUX_CIBLES,
    PARTIES,
    charger_donnees,
    creer_rapport
)

st.set_page_config(page_title="Générateur de rapports commerciaux", layout="centered")

st.title("📊 Générateur de rapports commerciaux")

uploaded_file = st.file_uploader("📁 Importer le fichier Excel", type=["xlsx"])
uploaded_logo = st.file_uploader("🖼️ Importer le logo (facultatif)", type=["png", "jpg", "jpeg"])

col1, col2 = st.columns(2)
with col1:
    mois = st.selectbox("📅 Mois", list(range(1, 13)), index=datetime.now().month - 1)
with col2:
    annee = st.selectbox("📆 Année", list(range(2022, 2026)), index=3)

col3, col4 = st.columns(2)
with col3:
    jour_debut = st.number_input("📍 Jour de début", min_value=1, max_value=31, value=1)
with col4:
    jour_fin = st.number_input("📍 Jour de fin", min_value=1, max_value=31, value=31)

if uploaded_file:
    if st.button("🚀 Générer les rapports"):
        with st.spinner("Génération des rapports en cours..."):

            with tempfile.TemporaryDirectory() as temp_dir:
                excel_path = os.path.join(temp_dir, "data.xlsx")
                with open(excel_path, "wb") as f:
                    f.write(uploaded_file.read())

                logo_path = None
                if uploaded_logo:
                    logo_path = os.path.join(temp_dir, uploaded_logo.name)
                    with open(logo_path, "wb") as f:
                        f.write(uploaded_logo.read())

                output_dir = os.path.join(temp_dir, "rapports")
                img_dir = os.path.join(temp_dir, "images")
                os.makedirs(output_dir, exist_ok=True)
                os.makedirs(img_dir, exist_ok=True)

                data = charger_donnees(excel_path, mois, annee, jour_debut, jour_fin)
                if data:
                    commerciaux = list(data[next(iter(data))].keys())
                    for com in commerciaux:
                        creer_rapport(com, data, mois, annee, output_dir, excel_path, logo_path, img_dir)

                    zip_path = shutil.make_archive(os.path.join(temp_dir, "Rapports_Commerciaux"), 'zip', output_dir)
                    st.success("✅ Rapport généré avec succès.")
                    st.download_button("📥 Télécharger le fichier ZIP", open(zip_path, "rb"), file_name="Rapports_Commerciaux.zip")
                else:
                    st.warning("Aucune donnée trouvée pour les filtres choisis.")
