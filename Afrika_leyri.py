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
# Upload du fichier Excel
Chargement = pd.read_excel("Tableau_bord_23_au_26_07.xlsx", engine='openpyxl', sheet_name="KAMLAC")
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
donnee["Prix Total"] = donnee["Quantites"] * donnee["Prix_Unitaire"]
# donnee["Mois"] = donnee["Date"].dt.month

# Choix de l’onglet
# Définir les bornes du slider
date1 = donnee["Date"].unique().tolist()

# Slider Streamlit pour filtrer une plage de dates
col=st.columns(5)
start_date=col[1].selectbox("Début",date1)
end_date=col[3].selectbox("Fin",date1)


# Filtrer les données selon la plage sélectionnée
donnee = donnee[(donnee["Date"] >= start_date) & (donnee["Date"] <= end_date)]

# Afficher les résultats
st.write(f"Résultats entre {start_date} et {end_date} :")

menu = st.sidebar.selectbox("Choisissez une option", ["Données", "Opération"])

if menu == "Données":
    st.subheader("Contenu de la feuille sélectionnée :")
    st.dataframe(donnee)
    operation="Données"
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
       # st.dataframe(donnee[nomcol])
    else:
        st.write("Choisissez une opération pour continuer")
else:
    st.write(
        "La colonne Opération ne se trouve pas dans les colonnes selectionnées"
    )
if menu == "Opération" or operation == "Livraison":
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
elif menu == "Opération" or operation == "Commande":
    donnee_agre = (
        donnee.groupby(["Date", "Prenom_Nom_RZ", "zone","Produit"])
        .agg({"Quantites": "sum"})
        .reset_index()
    )

    donnee_agre = donnee_agre.rename(
    columns={
        "Quantites": "Quantités",
    }
    )

if menu == "Opération" :
    donnee_ordre = donnee_agre.sort_values(by=["Date", "Prenom_Nom_RZ"], ascending=False)


# 🔧 Fonction pour créer l'image avec les infos en haut
def generate_png_report(df, date_min,date_max):
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
    rect = Rectangle((header_x, header_y), header_width, header_height,
                    transform=ax.transAxes,
                    fill=False, color='black', linewidth=1.5)
    ax.add_patch(rect)
    # En-tête
    plt.text(0.45, 0.9, f"{operation} du {date_min} au {date_max}", ha='center', fontsize=14, transform=ax.transAxes, weight='bold')
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
            #donnee_ordre.to_excel(writer, sheet_name=f"Récapitulatif des {nom_nouvelle_feuille}", index=False)
        


        st.success("✅ Fichier enregistré avec succès.")

        # Bouton de téléchargement
        st.download_button(
            label="📥 Télécharger",
            data=output.getvalue(),
            file_name="KAMLAC_RZ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# Afficher le tableau récapitulatif

st.subheader("Regroupement des ventes et ordonnées par Date et Prénom du RZ")
st.dataframe(donnee_ordre)
if operation == "Commande":
    png_bytes = generate_png_report(donnee_ordre, date_min=start_date, date_max=end_date)
    # ✅ Afficher l'aperçu de l'image directement dans l'interface
    #st.image(png_bytes, caption="", use_container_width=True)
    #png_bytes = generate_png_report(donnee_ordr[(donnee_ordr["TATA"] == prom)])
    st.download_button(
        label="📥 Télécharger le rapport en PNG",
        data=png_bytes,
        file_name=f"{operation}_du_{start_date}_au_{end_date}.png",
        mime="image/png"
    )