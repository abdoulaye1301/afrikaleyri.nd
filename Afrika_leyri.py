import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl import load_workbook
import io
import matplotlib.pyplot as plt
from io import BytesIO
from matplotlib.patches import Rectangle
from matplotlib.patches import FancyBboxPatch
import textwrap

st.set_page_config(
    page_title="Ingénieur NDAO", layout="wide", page_icon="ndao abdoulaye.png"
)
profil = Image.open("Logo Afrika Leyri.png")
st.logo(profil)

st.title("INFORMATION SUR LES DONNÉES DES RZ")
def analyse_donnee(Chargement):
    #st.sidebar.file_uploader(" 📁 Charger un fichier Excel", type=["xlsx"])


    # Lire toutes les feuilles
    # xls = pd.ExcelFile(Chargement)
    #feuilles = Chargement.sheet_names

    # Choisir une feuille à modifier
    #feuille_selectionnee = st.sidebar.selectbox(
    #   "Choisissez une feuille à éditer :", feuilles
    #)

    # Charger la feuille sélectionnée
    donnee = Chargement.copy()
    # Définir les chemins des fichiers source et destination
    donnee["Date"] = donnee["Date"].dt.date
    donnee=donnee.sort_values(by="Date", ascending=False)
    donnee["Prix Total"] = donnee["Quantites"] * donnee["Prix_Unitaire"]
    # donnee["Mois"] = donnee["Date"].dt.month

    # Choix de l’onglet
    # Définir les bornes du slider
    date1 = donnee["Date"].unique().tolist()

    # Slider Streamlit pour filtrer une plage de dates
    col=st.columns(5)

    # Lire toutes les feuilles
    # xls = pd.ExcelFile(Chargement)
    #feuilles = Chargement.sheet_names

    # Choisir une feuille à modifier
    #feuille_selectionnee = st.sidebar.selectbox(
    #   "Choisissez une feuille à éditer :", feuilles
    #)

    # Charger la feuille sélectionnée
    donnee = Chargement.copy()
    # Définir les chemins des fichiers source et destination
    donnee["Date"] = donnee["Date"].dt.date
    donnee=donnee.sort_values(by="Date", ascending=False)
    donnee["Prix Total"] = donnee["Quantites"] * donnee["Prix_Unitaire"]
    # donnee["Mois"] = donnee["Date"].dt.month

    # Choix de l’onglet
    # Définir les bornes du slider
    date1 = donnee["Date"].unique().tolist()

    # Slider Streamlit pour filtrer une plage de dates
    col=st.columns(5)



    menu = st.sidebar.selectbox("Choisissez une option", ["Données", "Opération"])

    if menu == "Données":
        
        start_date=col[1].selectbox("Début",date1)
        end_date=col[3].selectbox("Fin",date1)


        # Filtrer les données selon la plage sélectionnée
        donnee = donnee[(donnee["Date"] >= start_date) & (donnee["Date"] <= end_date)]
        st.subheader("Contenu de la feuille sélectionnée :")
        st.dataframe(donnee)
        operation="Données"
    elif menu == "Opération":
        # Sélectionner la plage de dates
        end_date=col[2].selectbox("Selectionnez une date",date1)


        # Filtrer les données selon la plage sélectionnée
        donnee = donnee[(donnee["Date"] == end_date)]
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
        # st.dataframe(donnee[nomcol])
        #else:
        # st.write("Choisissez une opération pour continuer")
    else:
        st.write(
            "La colonne Opération ne se trouve pas dans les colonnes selectionnées"
        )
    if menu == "Opération" and operation == "Livraison":
        donnee_agre = (
            donnee.groupby(["Date", "Prenom_Nom_RZ", "secteur","Produit"])
            .agg({"Quantites": "sum", "Prix Total": "sum"})
            .reset_index()
        )

        donnee_agre = donnee_agre.rename(
        columns={
            "Quantites": "Quantités",
            "Prix Total": "Prix Total",
        }
        )
        donnee_ordre = donnee_agre.sort_values(by=["Date", "Prenom_Nom_RZ"], ascending=False)
    elif menu == "Opération" and operation == "Commande":
        donnee_agre = (
            donnee.groupby(["Prenom_Nom_RZ","secteur","Produit"])
            .agg({"Quantites": "sum"})
            .reset_index()
        )

        donnee_agre = donnee_agre.rename(
        columns={
            "Quantites": "Quantités",
            "Prenom_Nom_RZ": "RZ",
            "secteur": "Secteur"
        }
        )
        donnee_ordre = donnee_agre.sort_values(by=["RZ"], ascending=False)


    # 🔧 Fonction pour créer l'image avec les infos en haut
    def generate_png_report(df,date_max):
        fig, ax = plt.subplots(figsize=(12, len(df) * 0.6+1.5))
        ax.axis('off')
        # ✅ Texte commentaire à droite du cadre
        # retour à la ligne pour le commentaire
        #df["Commentaire"] = df["Commentaire"].apply(
            # lambda x: (textwrap.wrap(x, width=30)) if isinstance(x, str) else x)
            #textwrap.fill(commentaire, width=45)
        # Dimensions du rectangle d’en-tête (valeurs relatives à l’axe)
        header_x = 0.001    # gauche
        header_y = 0.85    # position verticale bas du bloc
        header_width = 0.996
        header_height = 0.12

        # ✅ Dessiner le rectangle d'encadrement
        #rect = Rectangle((header_x, header_y), header_width, header_height,
        #               transform=ax.transAxes,
        #              fill=False, color='black', linewidth=1.5)
        #ax.add_patch(rect)
        # En-tête
        plt.text(0.45, 0.9, f"{operation} du {date_max}", ha='center', fontsize=14, transform=ax.transAxes, weight='bold')
        # Tableau matplotlib
        table = ax.table(cellText=df.values,
                        colLabels=df.columns,
                        cellLoc='center',
                        loc='center')

        table.scale(1, 1.5)
        buffer = BytesIO()
        plt.savefig(buffer, format='png', bbox_inches='tight', dpi=200)
        plt.close()
        buffer.seek(0)
        return buffer



    #donnee_agre["Date"] = donnee_agre["Date"].dt.strftime("%d/%m/%Y")

    if menu == "Données":
        # Afficher les résultats
        #st.write(f"Résultats entre {start_date} et {end_date} :")
        nom_nouvelle_feuille = operation
        if st.button("Sauvegarder"):
            # Définir le nom sous lequel la feuille sera enregistrée dans le fichier de destination
            if nom_nouvelle_feuille.strip() == "":
                st.warning(
                    "Veuillez renseigner le nom de la feuille dans la barre de naviagation."
                )
            else:
                # Copie de toutes les feuilles existantes dans un nouveau Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    donnee.to_excel(writer, index=False, sheet_name='Feuille1')
                output.seek(0)  # repositionner au début du fichier
                memorise_nouvelle_feuille = output
                    # Ajouter la feuille modifiée
                    #donnee.to_excel(writer, sheet_name=nom_nouvelle_feuille, index=False)
                    #donnee_ordre.to_excel(writer, sheet_name=f"Récapitulatif des {nom_nouvelle_feuille}", index=False)
                


                st.success("✅ Fichier enregistré avec succès.")

                # Bouton de téléchargement
                st.download_button(
                    label="📥 Télécharger",
                    data=memorise_nouvelle_feuille,
                    file_name="KAMLAC_RZ.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


    # Afficher le tableau récapitulatif
    if menu == "Opération" :
        #st.subheader("Regroupement des ventes et ordonnées par Date et Prénom du RZ")
        st.dataframe(donnee_ordre)
    if operation == "Commande":
        png_bytes = generate_png_report(donnee_ordre, date_max=end_date)
        # ✅ Afficher l'aperçu de l'image directement dans l'interface
        #st.image(png_bytes, caption="", use_container_width=True)
        #png_bytes = generate_png_report(donnee_ordr[(donnee_ordr["TATA"] == prom)])
        st.download_button(
            label="📥 Télécharger le rapport en PNG",
            data=png_bytes,
            file_name=f"{operation}_du__{end_date}.png",
            mime="image/png"
        )

#===========================BARRE DE NAVIGATION===========================
navi=st.sidebar.radio("Navigation",["Analyse","Traitement"])
# Chargement du fichier
if navi=="Traitement":
    uploaded_file = st.file_uploader("📁 Choisir un fichier (CSV ou Excel)", type=["csv", "xlsx"])

    if uploaded_file is not None:
        try:
            # Vérifier le type de fichier
            if uploaded_file.name.endswith(".csv"):
                Chargement = pd.read_csv(uploaded_file)
            else:
                Chargement = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name="KAMLAC")
            st.success("✅ Fichier chargé avec succès !")
            # Afficher les données et permettre l'édition
            analyse_donnee(Chargement)
        except Exception as e:
            st.error(f"Erreur lors du chargement du fichier : {e}")
# Upload du fichier Excel
elif navi=="Analyse":
    Chargement = pd.read_excel("Tableau_bord_23_au_26_07.xlsx", engine='openpyxl', sheet_name="KAMLAC")

    # Afficher les données et permettre l'édition
    analyse_donnee(Chargement)