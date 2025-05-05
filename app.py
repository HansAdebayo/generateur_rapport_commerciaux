import streamlit as st
import tempfile
import os
import shutil
from datetime import datetime

from rapport_generator import (
    COMMERCIAUX_CIBLES,
    PARTIES,
    charger_donnees,
    creer_rapport
)

# === CONFIG PAR DÉFAUT ===
logo_path = "logo_watt.png"

# === UI ===
st.title("📊 Générateur de rapports commerciaux")

uploaded_file = st.file_uploader("📂 Importer le fichier Excel", type=["xlsx"])

annee = st.selectbox("📅 Choisir l'année", list(range(2023, 2026)), index=2)
mois = st.selectbox("🗓️ Choisir le mois", list(range(1, 13)), index=datetime.now().month - 1)
jour_debut = st.number_input("📆 Jour de début", min_value=1, max_value=31, value=1)
jour_fin = st.number_input("📆 Jour de fin", min_value=1, max_value=31, value=31)

if uploaded_file and jour_debut <= jour_fin:
    with tempfile.TemporaryDirectory() as temp_dir:
        excel_path = os.path.join(temp_dir, uploaded_file.name)
        with open(excel_path, "wb") as f:
            f.write(uploaded_file.read())

        output_dir = os.path.join(temp_dir, "rapports")
        os.makedirs(output_dir, exist_ok=True)

        st.info("📑 Traitement en cours...")

        data = charger_donnees(excel_path, mois, annee, jour_debut, jour_fin)

        if data:
            commerciaux = list(data[next(iter(data))].keys())
            for com in commerciaux:
                creer_rapport(com, data, mois, annee, output_dir, excel_path, logo_path, temp_dir)

            zip_path = shutil.make_archive(os.path.join(temp_dir, "rapports_commerciaux"), 'zip', output_dir)

            with open(zip_path, "rb") as f:
                st.success("✅ Rapport généré avec succès.")
                st.download_button("📥 Télécharger le fichier ZIP", f, file_name="rapports_commerciaux.zip")

        else:
            st.warning("⚠️ Aucune donnée trouvée pour les filtres choisis.")
