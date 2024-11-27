import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import plotly.express as px

@st.cache_data
def charger_donnees(fichier):
    return pd.read_excel(fichier)

st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")

st.title("📊 Analyse des interventions des opérateurs")

fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        col_prenom_nom = st.selectbox("Choisissez la colonne pour 'Prénom et nom'", df_principal.columns)
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        
        operateurs = df_principal[col_prenom_nom].unique()
        operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année", "Total"]
        periode_selectionnee = st.selectbox("Choisissez une période", periodes)
        
        date_min = pd.to_datetime(df_principal[col_date]).min().date()
        date_max = pd.to_datetime(df_principal[col_date]).max().date()
        debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
    
    if st.button("Analyser"):
        df_principal[col_date] = pd.to_datetime(df_principal[col_date])
        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        # Filtrer les données pour le graphique
        df_graph = df_principal[df_principal[col_date].dt.date >= debut_periode]

        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)
        
        repetitions_graph = df_graph[df_graph[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        
        # Calcul des répétitions pour le tableau (toutes les dates)
        repetitions_tableau = df_principal[df_principal[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        
        with col2:
            if periode_selectionnee != "Total":
                fig = px.bar(repetitions_graph, x=periode_selectionnee, y='Repetitions', color=col_prenom_nom, barmode='group',
                             title=f"Répétitions par {periode_selectionnee.lower()} pour les opérateurs sélectionnés (à partir de {debut_periode})")
            else:
                fig = px.bar(repetitions_graph, x=col_prenom_nom, y='Repetitions',
                             title=f"Total des répétitions pour les opérateurs sélectionnés (à partir de {debut_periode})")
            
            st.plotly_chart(fig)
        
        st.subheader(f"Tableau des répétitions par {periode_selectionnee.lower()} (toutes les dates)")
        
        colonnes_affichage = [col_prenom_nom, periode_selectionnee, 'Repetitions'] if periode_selectionnee != "Total" else [col_prenom_nom, 'Repetitions']
        tableau_affichage = repetitions_tableau[colonnes_affichage]
        
        st.dataframe(tableau_affichage, use_container_width=True)
        
        # Tirage au sort
        st.subheader("Tirage au sort de deux lignes par opérateur")
        df_filtre = df_principal[df_principal[col_date].dt.date >= debut_periode]
        for operateur in operateurs_selectionnes:
            st.write(f"Tirage pour {operateur}:")
            df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
            lignes_tirees = df_operateur.sample(n=min(2, len(df_operateur)))
            if not lignes_tirees.empty:
                st.dataframe(lignes_tirees, use_container_width=True)
            else:
                st.write("Pas de données disponibles pour cet opérateur dans la période sélectionnée.")
            st.write("---")

    if st.checkbox("Afficher toutes les données"):
        st.dataframe(df_principal)
