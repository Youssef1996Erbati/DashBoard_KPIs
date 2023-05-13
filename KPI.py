# import openpyxl
# import openpyxl.utils
import pandas as pd
# import numpy as np
import streamlit as st
import plotly.express as px
#from .data import df
# import io
# from openpyxl.chart import BarChart


def dashboard():
    st.set_page_config(page_title="Synthese des KPIs", page_icon=":bar_chart:", layout="wide", initial_sidebar_state="auto")
    # ---- SIDEBAR ----
    st.sidebar.image("TDF.ico",width = 50)
    fichier = st.sidebar.file_uploader("Upload a data set in EXCEL format", type="xlsx")
    if fichier is not None:
        df=pd.read_excel(fichier)
        st.sidebar.header("Please Filter Here:")
        region = st.sidebar.selectbox(
            "Selectionner la région:",df["Région"].dropna().unique()
        )
        equipe = st.sidebar.multiselect(
            "Selectionner l'équipe:",
            options=df["Entité actuelle de l'intervenant"].dropna().unique(),
            default=df["Entité actuelle de l'intervenant"].dropna().unique()
        )
        Annee = st.sidebar.multiselect(
            "Selectionner l'année:",
            options=df["Année"].dropna().unique(),
            default=df["Année"].dropna().unique()
        )
        Mois = st.sidebar.multiselect(
            "Selectionner le mois:",
            options=df["Mois"].dropna().unique(),
            default=df["Mois"].dropna().unique()
        )
        # ---- MAINPAGE ----
        st.title(":bar_chart: Synthese KPI de la région SE")
        st.header("La productivité au Pôle Opérations se mesure avec les indicateurs suivants :")
        st.markdown('<h2 style="color: blue;">Taux de Rendement global (TRG)</h2>', unsafe_allow_html=True)
        st.subheader("Ratio entre le temps sur site que passe le technicien et le temps disponible")
        df_TRG = pd.read_excel(fichier)
        
        
        # Calcul de TRG

        # Créer les colonnes calculées
        df_TRG['TTE'] = df_TRG['Somme TTE'].where(df_TRG['Activité '].isin(["3- Towerco – Contribution au déploiement", "4- Towerco Infra - Préventif et planifiable", "5- Towerco - Correctif", "6- AV - Contribution au déploiement et VDR", "7- AV - Préventive", "8- AV - Correctif"]))
        df_TRG['TTE_HJM'] = df_TRG['TTE HJ'].where(df_TRG['Activité '] == "9- Management")
        df_TRG['TTE_HJ'] = df_TRG['TTE HJ'].where(df_TRG['Activité '].isin(["3- Towerco – Contribution au déploiement", "4- Towerco Infra - Préventif et planifiable", "5- Towerco - Correctif", "6- AV - Contribution au déploiement et VDR", "7- AV - Préventive", "8- AV - Correctif", "12- Support à la production"]))
        df_TRG['Trajet_Aller'] = df_TRG['Trajet Aller'].where(df_TRG['Activité '].isin(["3- Towerco – Contribution au déploiement", "4- Towerco Infra - Préventif et planifiable", "5- Towerco - Correctif", "6- AV - Contribution au déploiement et VDR", "7- AV - Préventive", "8- AV - Correctif"]))

        # Grouper les données par Année, Mois et Mnemonique intervenant et calculer les sommes
        resultat_TRG = df_TRG.groupby(['Année', 'Mois', "Entité actuelle de l'intervenant","Mnemonique intervenant"])[['TTE','TTE_HJM', 'TTE_HJ', 'Trajet_Aller']].sum().reset_index()

        resultat_TRG['TRG'] = (resultat_TRG['TTE'] - resultat_TRG['Trajet_Aller']) / (7.8 * (resultat_TRG['TTE_HJ'] + resultat_TRG['TTE_HJM']))
        resultat_TRG['TRG'] = pd.Series(["{0:.2f}%".format(val * 100) for val in resultat_TRG['TRG']], index = resultat_TRG.index)

        resultat_TRG1 = resultat_TRG[resultat_TRG["Entité actuelle de l'intervenant"].isin(equipe) & resultat_TRG["Année"].isin(Annee) & resultat_TRG["Mois"].isin(Mois)]
        resultat_TRG1['TRG'] = pd.to_numeric(resultat_TRG1['TRG'].str.replace('%', '').replace('nan', '0'))
        TCD_TRG = pd.pivot_table(resultat_TRG1, values=['TTE', 'TTE_HJM', 'TTE_HJ', 'Trajet_Aller'], index="Entité actuelle de l'intervenant",aggfunc=sum)
        TCD_TRG['TRG'] = (TCD_TRG['TTE'] - TCD_TRG['Trajet_Aller']) / (7.8 * (TCD_TRG['TTE_HJ'] + TCD_TRG['TTE_HJM']))
        TCD_TRG['TRG'] = pd.Series(["{0:.2f}%".format(val * 100) for val in TCD_TRG['TRG']], index = TCD_TRG.index)

        # Calcul du TRG total
        total_TTE = TCD_TRG['TTE'].sum()
        total_Trajet_Aller = TCD_TRG['Trajet_Aller'].sum()
        total_TTE_HJ = TCD_TRG['TTE_HJ'].sum()
        total_TTE_HJM = TCD_TRG['TTE_HJM'].sum()

        TRG_total = (total_TTE - total_Trajet_Aller) / (7.8 * (total_TTE_HJ + total_TTE_HJM))

        TRG_total = "{0:.2f}%".format(TRG_total * 100)


        col1, col2 = st.columns(2)

        with col1:
            st.markdown('<div style="border: 2px solid blue; padding: 10px; text-align: center; font-weight: bold;">Le TRG total de la région est :</div>', unsafe_allow_html=True)

        with col2:
            st.markdown(f'<div style="border: 2px solid blue; padding: 10px; text-align: center; font-weight: bold;">{TRG_total}</div>', unsafe_allow_html=True)

        fig_TRG=px.bar(TCD_TRG,TCD_TRG.index, y='TRG')

        col1, col2 = st.columns(2)

        col1.subheader("Synthese de TRG par Equipe")
        col1.write(TCD_TRG)

        col2.subheader("Représentation graphique du TRG par Equipe")
        col2.plotly_chart(fig_TRG)
        # ------------------------------------------------------------------------------------------------------------
        st.markdown('<h2 style="color: blue;">Tenue des gabarits </h2>', unsafe_allow_html=True)
        st.subheader("Temps passé sur site par rapport au temps gabarit standard d’une année par rapport à la précédente")
        # Calcul de réalisé Gammé
        # Charger le fichier Excel dans un DataFrame
        df = pd.read_excel(fichier)
        df_RG = pd.read_excel(fichier)
        # Exclusion des valeurs dans la colonne "Gamme - BT"
        excluded_values = ["MANAGEMENT", "Z_TACHE_REFERENT", "VISITE_MEDICALE", "COMPAGNONNAGE",
                        "Z_POINT_RELAIS", "FORMATION", "Z_COMPTE_RENDU_1H_1U"]

        df_RG = df_RG[~df_RG['Gamme - BT'].isin(excluded_values)]
        df_RG['Gamme - BT'] = df_RG['Gamme - BT'].astype(str)  # Convertir la colonne en type string
        df_RG = df_RG[~df_RG['Gamme - BT'].str.startswith("R_")]
        df_RG = df_RG.dropna(subset=['Gamme - BT'])
        # Exclusion des valeurs dans la colonne "Numéro du BT"
        df_RG['Numéro du BT'] = df_RG['Numéro du BT'].astype(str)  # Convertir la colonne en type string
        df_RG = df_RG[~df_RG['Numéro du BT'].str.contains("_C|_P")]
        resultat_RG = df_RG[df_RG["Entité actuelle de l'intervenant"].isin(equipe) & df_RG["Année"].isin(Annee) & df_RG["Mois"].isin(Mois)]
        TCD_RG = pd.pivot_table(resultat_RG, values=['Somme TTE', 'Trajet Aller', 'Durée de prestation'], index="Entité actuelle de l'intervenant",aggfunc=sum)
        TCD_RG['Réalisé_Gammé'] = (TCD_RG['Somme TTE'] - TCD_RG['Trajet Aller']) / TCD_RG['Durée de prestation']
        TCD_RG['Réalisé_Gammé'] = pd.Series(["{0:.2f}%".format(val * 100) for val in TCD_RG['Réalisé_Gammé']], index = TCD_RG.index)
        # Calcul du TRG total
        total_STTE = TCD_RG['Somme TTE'].sum()
        total_TAller = TCD_RG['Trajet Aller'].sum()
        total_Presta = TCD_RG['Durée de prestation'].sum()

        RG_total = (total_STTE - total_TAller) / total_Presta

        RG_total = "{0:.2f}%".format(RG_total * 100)


        col1, col2 = st.columns(2)

        with col1:
            st.markdown('<div style="border: 2px solid blue; padding: 10px; text-align: center; font-weight: bold;">Le Réalisé Gammé total de la région est :</div>', unsafe_allow_html=True)

        with col2:
            st.markdown(f'<div style="border: 2px solid blue; padding: 10px; text-align: center; font-weight: bold;">{RG_total}</div>', unsafe_allow_html=True)

        fig_RG=px.bar(TCD_RG,TCD_RG.index, y='Réalisé_Gammé')

        col1, col2 = st.columns(2)

        col1.subheader("Synthese de Réalisé Gammé par Equipe")
        col1.write(TCD_RG)

        col2.subheader("Représentation graphique du Réalisé Gammé par Equipe")
        col2.plotly_chart(fig_RG)
        #------------------------------------ Test ----------------------------------------
        # Combiner les deux tables croisées dynamiques en utilisant la méthode merge de pandas
        table_combinee =pd.merge(TCD_TRG, TCD_RG, on="Entité actuelle de l'intervenant")
        table_KPI=table_combinee[["TRG","Réalisé_Gammé"]]
        st.write(table_KPI)
        #------------------------------------ Export de la Synthese ----------------------------------------
        @st.cache
        def generer_fichier_excel():
            # Création du DataFrame pour la feuille "TRG"
            df_f1 = pd.DataFrame(TCD_TRG)  # Remplacez TCD_TRG par vos données réelles
            df_f1_fig = pd.DataFrame(fig_TRG)  # Remplacez fig_TRG par vos données réelles
            
            # Création du DataFrame pour la feuille "Réalisé Gammé"
            df_f2 = pd.DataFrame(TCD_RG)  # Remplacez TCD_RG par vos données réelles
            df_f2_fig = pd.DataFrame(fig_RG)  # Remplacez fig_RG par vos données réelles
            # Enregistrement du fichier Excel
            with pd.ExcelWriter('Synthèse_KPIs.xlsx') as writer:
                df_f1.to_excel(writer, sheet_name='TRG', index=False)
                df_f1_fig.to_excel(writer, sheet_name='TRG', startrow=len(df_f1) + 2, index=False, header=False)
                df_f2.to_excel(writer, sheet_name='Réalisé Gammé', index=False)
                df_f2_fig.to_excel(writer, sheet_name='Réalisé Gammé', startrow=len(df_f2) + 2, index=False, header=False)
        # Code Streamlit pour créer le bouton et appeler la fonction
        st.sidebar.header("Exporter la Synthèse en format Excel")
        if st.sidebar.button('Exporter la Synthèse'):
            generer_fichier_excel()
            st.sidebar.success('Le fichier Excel a été généré avec succès.')
    
    
if __name__ == "__main__":
    dashboard()
    
    