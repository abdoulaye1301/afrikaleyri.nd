import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl import load_workbook
import io

st.set_page_config(
    page_title="Ingénieur NDAO", layout="wide", page_icon="ndao abdoulaye.png"
)
profil = Image.open("Logo Afrika Leyri.png")
st.logo(profil)

st.title("Éditeur Excel avec plusieurs feuilles")
# Upload du fichier Excel
Chargement = st.sidebar.file_uploader(" 📁 Charger un fichier Excel", type=["xlsx"])

if Chargement:
    # Lire toutes les feuilles
    xls = pd.ExcelFile(Chargement)
    feuilles = xls.sheet_names

    # Choisir une feuille à modifier
    feuille_selectionnee = st.sidebar.selectbox(
        "Choisissez une feuille à éditer :", feuilles
    )

    # Charger la feuille sélectionnée
    donnee = pd.read_excel(xls, sheet_name=feuille_selectionnee)
    # Définir les chemins des fichiers source et destination
    donnee["Date"] = donnee["Date"].dt.date
    donnee["Prix Total"] = donnee["Quantites"] * donnee["Prix_Unitaire"]
    # donnee["Mois"] = donnee["Date"].dt.month

    # Choix de l’onglet
    # Définir les bornes du slider
    min_date = min(donnee["Date"])
    max_date = max(donnee["Date"])

    # Slider Streamlit pour filtrer une plage de dates
    start_date, end_date = st.slider(
        "Sélectionnez une plage de dates",
        min_value=min_date,
        max_value=max_date,
        value=(min_date, max_date),  # valeur par défaut (tout)
        format="YYYY/MM/DD"
    )

    # Filtrer les données selon la plage sélectionnée
    donnee = donnee[(donnee["Date"] >= start_date) & (donnee["Date"] <= end_date)]

    # Afficher les résultats
    st.write(f"Résultats entre {start_date} et {end_date} :")

    menu = st.sidebar.selectbox("Navigation", ["Kamlac", "Opération"])

    if menu == "Kamlac":
        st.subheader("Contenu de la feuille sélectionnée :")
        st.dataframe(donnee)
        operation="Kamlac"
    elif menu == "Opération":
        operation = st.sidebar.selectbox(
            "Type d'opération", ("Commande", "Livraison", "Aucune")
        )
        donnee = donnee[donnee["Operation"] == operation]
        if operation == "Aucune":
            nomcol = donnee.columns.tolist()
            nomcol.remove("Prix_Unitaire")
            nomcol.remove("Quantites")
            nomcol.remove("Produit")
            nomcol.remove("Prix Total")
            st.dataframe(donnee[nomcol])
        else:
            st.dataframe(donnee)
    else:
        st.write(
            "La colonne Opération ne se trouve pas dans les colonnes selectionnées"
        )

    donnee_agre = (
        donnee.groupby(["Date", "Prenom_Nom_RZ", "secteur","Produit"])
        .agg({"Quantites": "sum", "Prix Total": "sum"})
        .reset_index()
    )
    
    st.subheader("Regroupement des ventes et ordonnées par Date et Prénom du RZ")
    donnee_agre = donnee_agre.rename(
        columns={
            "Quantites": "Quantités",
            "Prix Total": "Prix Total",
        }
    )
    donnee_ordre = donnee_agre.sort_values(by=["Date", "Prenom_Nom_RZ"], ascending=False)
    #donnee_agre["Date"] = donnee_agre["Date"].dt.strftime("%d/%m/%Y")
    st.dataframe(donnee_ordre)
    nom_nouvelle_feuille = st.sidebar.text_input("Nom de la feuille :",value=operation)
    if st.button("Sauvegarder"):
        # Définir le nom sous lequel la feuille sera enregistrée dans le fichier de destination
        if nom_nouvelle_feuille.strip() == "":
            st.warning(
                "Veuillez renseigner le nom de la feuille dans la barre de naviagation."
            )
        else:
            # Charger le fichier original dans openpyxl
            memorise_nouvelle_feuille = io.BytesIO(Chargement.getvalue())
            wb = load_workbook(memorise_nouvelle_feuille)

            # Supprimer la feuille si elle existe déjà (et n'est pas la seule)
            if nom_nouvelle_feuille in wb.sheetnames:
                if len(wb.sheetnames) > 1:
                    del wb[nom_nouvelle_feuille]
                else:
                    st.error("Impossible de supprimer la seule feuille visible.")
                    st.stop()

            # Copie de toutes les feuilles existantes dans un nouveau Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                # Copier les anciennes feuilles
                for feuille in wb.sheetnames:
                    data = pd.read_excel(memorise_nouvelle_feuille, sheet_name=feuille)
                    data.to_excel(writer, sheet_name=feuille, index=False)

                # Ajouter la feuille modifiée
                donnee.to_excel(writer, sheet_name=nom_nouvelle_feuille, index=False)
                donnee_ordre.to_excel(writer, sheet_name=f"Récapitulatif des {nom_nouvelle_feuille}", index=False)
            


            st.success("✅ Fichier modifié avec succès.")

            # Bouton de téléchargement
            st.download_button(
                label="📥 Télécharger",
                data=output.getvalue(),
                file_name="KAMLAC_RZ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Veuillez charger un fichier pour commencer.")
