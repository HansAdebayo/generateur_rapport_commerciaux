
import streamlit as st
from datetime import datetime
import tempfile
import os
import shutil
from rapport_generator import (
    sanitize_filename, normalize, detect_column, convert_mois_to_int,
    PARTIES, COMMERCIAUX_CIBLES, charger_donnees, creer_rapport
)
import pandas as pd

st.set_page_config(page_title="Générateur de rapports commerciaux", layout="centered")

st.title("📊 Générateur de rapports commerciaux")

uploaded_file = st.file_uploader("📁 Importer le fichier Excel", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    mois = st.selectbox("📅 Mois", list(range(1, 13)), index=datetime.now().month - 1)
with col2:
    annee = st.selectbox("📆 Année", list(range(2022, 2026)), index=3)

if uploaded_file:
    if st.button("🚀 Générer les rapports"):
        with st.spinner("Génération des rapports en cours..."):

            # Créer répertoire temporaire
            temp_dir = tempfile.mkdtemp()
            excel_path = os.path.join(temp_dir, "data.xlsx")
            with open(excel_path, "wb") as f:
                f.write(uploaded_file.read())

            # Répertoire de sortie
            output_dir = os.path.join(temp_dir, "rapports")
            os.makedirs(output_dir, exist_ok=True)

            # Charger les données et générer les rapports
            data = charger_donnees(excel_path, mois, annee)
            if data:
                commerciaux = list(data[next(iter(data))].keys())
                for com in commerciaux:
                    creer_rapport(com, data, mois, annee, output_dir)

                # Zipper les fichiers
                zip_path = shutil.make_archive(os.path.join(temp_dir, "Rapports_Commerciaux"), 'zip', output_dir)
                st.success("✅ Rapport généré avec succès.")
                st.download_button("📥 Télécharger le fichier ZIP", open(zip_path, "rb"), file_name="Rapports_Commerciaux.zip")
            else:
                st.warning("Aucune donnée trouvée pour les filtres choisis.")
