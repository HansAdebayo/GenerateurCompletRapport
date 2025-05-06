
import streamlit as st
from datetime import datetime
import tempfile
import os
import shutil
import pandas as pd
from rapport_generator import charger_donnees, creer_rapport, detect_column, sanitize_filename
from rdv_generator import load_rdv_data

COMMERCIAUX_CIBLES = ['Sandra', 'Ophélie', 'Arthur', 'Grégoire', 'Tania']

st.set_page_config(page_title="📊 Rapports complets commerciaux", layout="centered")
st.title("📊 Générateur de rapports commerciaux complets (avec RDV)")

uploaded_rapport = st.file_uploader("📁 Fichier Excel principal (rapports)", type=["xlsx"])
uploaded_rdv = st.file_uploader("📁 Fichier Excel RDV", type=["xlsx"])
uploaded_logo = st.file_uploader("🖼️ Logo (facultatif)", type=["png", "jpg", "jpeg"])

selected_commerciaux = st.multiselect(
    "👤 Choisir les commerciaux",
    options=COMMERCIAUX_CIBLES,
    default=COMMERCIAUX_CIBLES
)

col1, col2 = st.columns(2)
with col1:
    mois = st.selectbox("📅 Mois", list(range(1, 13)), index=datetime.now().month - 1)
with col2:
    annee = st.selectbox("📆 Année", list(range(2022, 2030)), index=3)

col3, col4 = st.columns(2)
with col3:
    jour_debut = st.number_input("📍 Jour de début", min_value=1, max_value=31, value=1)
with col4:
    jour_fin = st.number_input("📍 Jour de fin", min_value=1, max_value=31, value=31)

if uploaded_rapport and uploaded_rdv and selected_commerciaux:
    if st.button("🚀 Générer les rapports fusionnés"):
        with st.spinner("📄 Génération des rapports en cours..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                rapport_path = os.path.join(temp_dir, "rapport.xlsx")
                rdv_path = os.path.join(temp_dir, "rdv.xlsx")

                with open(rapport_path, "wb") as f:
                    f.write(uploaded_rapport.read())
                with open(rdv_path, "wb") as f:
                    f.write(uploaded_rdv.read())

                logo_path = None
                if uploaded_logo:
                    logo_path = os.path.join(temp_dir, uploaded_logo.name)
                    with open(logo_path, "wb") as f:
                        f.write(uploaded_logo.read())

                output_dir = os.path.join(temp_dir, "rapports")
                img_dir = os.path.join(temp_dir, "images")

                data_by_part = charger_donnees(rapport_path, mois, annee, jour_debut, jour_fin)
                rdv_data = load_rdv_data(rdv_path, jour_debut, jour_fin, mois, annee)

                for com in selected_commerciaux:
                    rdv_df = None
                    for nom in rdv_data:
                        if nom.lower().startswith(com.lower()):
                            rdv_df = rdv_data[nom]
                            break
                    creer_rapport(com, data_by_part, mois, annee, jour_debut, jour_fin, output_dir, rapport_path, logo_path, img_dir, rdv_df)

                zip_path = shutil.make_archive(os.path.join(temp_dir, "Rapports_Commerciaux_RDV"), 'zip', output_dir)
                st.success("✅ Rapports générés avec succès.")
                st.download_button("📥 Télécharger les rapports fusionnés", open(zip_path, "rb"), file_name="Rapports_Commerciaux_RDV.zip")
