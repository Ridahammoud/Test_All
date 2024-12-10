import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from io import BytesIO
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import plotly.express as px

# Fonction de chargement des données
@st.cache_data
def charger_donnees(fichier):
    return pd.read_excel(fichier)


# Ajouter une colonne pour les équipes
def assign_team(name):
    if name in team_1_Christian:
        return "Team 1 Christian"
    elif name in team_2_Hakim:
        return "Team 2 Hakim"
    else:
        return "Non assigné"


# Fonction pour appliquer des styles aux moyennes
def style_moyennes(df, top_n=3, bottom_n=5):
    moyenne_totale = df['Repetitions'].mean()

    df_top = df.nlargest(top_n, 'Repetitions')
    df_bottom = df.nsmallest(bottom_n, 'Repetitions')

    def apply_styles(row):
        if row.name in df_top.index:
            return ['background-color: gold; color: black'] * len(row)  # Top 3
        elif row.name in df_bottom.index:
            return ['background-color: lightcoral; color: white'] * len(row)  # Flop 5
        elif row['Repetitions'] > moyenne_totale:
            return ['background-color: lightgreen'] * len(row)  # Supérieur à la moyenne
        else:
            return ['background-color: lightpink'] * len(row)  # Inférieur à la moyenne

    styled_df = df.style.apply(apply_styles, axis=1)
    return styled_df

# Fonction pour le tirage au sort
def tirage_au_sort(df, col_prenom_nom, col_date, debut_periode, fin_periode):
    df_filtre = df[(df[col_date].dt.date >= debut_periode) & (df[col_date].dt.date <= fin_periode)]
    resultats = []
    for operateur in df[col_prenom_nom].unique():
        df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
        if not df_operateur.empty:
            lignes_tirees = df_operateur.sample(n=min(2, len(df_operateur)))
            resultats.append((operateur, lignes_tirees))
    return resultats



# Configuration de la page Streamlit
st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")
st.title("📊 Analyse des interventions des opérateurs")

fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)

    col1, col2 = st.columns([2, 3])

    with col1:
        # Validation des colonnes requises
        required_columns = ["Prénom et nom", "Date et Heure début d'intervention", "Équipement", "Localisation", "Technique", "Opérationnel", "Photo"]
        missing_columns = [col for col in required_columns if col not in df_principal.columns]
        if missing_columns:
            st.error(f"Les colonnes suivantes sont manquantes : {', '.join(missing_columns)}")
            st.stop()
        
    # Conversion des dates
    col_prenom_nom = "Prénom et nom"
    col_date = "Date et Heure début d'intervention"
    df_principal[col_date] = pd.to_datetime(df_principal[col_date], errors='coerce')
    df_principal = df_principal.dropna(subset=[col_date])

    # Interface utilisateur
    operateurs = df_principal[col_prenom_nom].unique().tolist()
    operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
    
    periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année"]
    periode_selectionnee = st.selectbox("Choisissez une période", periodes)
    
    date_min = df_principal[col_date].min()
    date_max = df_principal[col_date].max()
    debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
    fin_periode = st.date_input("Fin de la période", min_value=debut_periode, max_value=date_max, value=date_max)

    if st.button("Analyser"):
        # Ajout de colonnes pour les périodes
        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        df_graph = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]
        
        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)

        repetitions_graph = df_graph.groupby(groupby_cols).size().reset_index(name='Repetitions')

        with col2:
            # graphique principal
            fig = px.bar(repetitions_graph = df_graph.groupby(groupby_cols).size().reset_index(name="Nombre de Rapports d'intervention"))
            for operateur in operateurs_selectionnes:
                df_operateur = repetitions_graph[repetitions_graph[col_prenom_nom] == operateur]
                fig.add_trace(go.Bar(x=df_operateur[periode_selectionnee],
                                     y=df_operateur['Repetitions'],
                                     name=operateur,
                                     text=df_operateur['Repetitions'],
                                     textposition='inside',
                                     hovertemplate='%{y}'))

                fig.update_layout(title=f"Nombre de rapports d'intervention (du {debut_periode} au {fin_periode})",
                              xaxis_title=periode_selectionnee,
                              yaxis_title="Répetitions",
                              template="plotly_dark")
                st.plotly_chart(fig)
            
        # Calcul des moyennes par opérateur et par période
        moyennes_par_periode = repetitions_graph.groupby([periode_selectionnee, col_prenom_nom])['Repetitions'].mean().reset_index()
        moyennes_par_operateur = moyennes_par_periode.groupby(['Prénom et nom'])['Repetitions'].mean().reset_index()
        moyenne_globale = moyennes_par_periode['Repetitions'].mean()

                # Graphique des moyennes avec moyenne globale
        fig1 = go.Figure()

        colors = px.colors.qualitative.Set1

        for i, operateur in enumerate(operateurs_selectionnes):
            df_operateur_moyenne = moyennes_par_periode[moyennes_par_periode[col_prenom_nom] == operateur]
            fig1.add_trace(go.Scatter(
                x=df_operateur_moyenne[periode_selectionnee],
                y=df_operateur_moyenne['Repetitions'],
                mode='lines+markers',
                name=operateur,
                line=dict(color=colors[i % len(colors)]),
                text=df_operateur_moyenne['Repetitions'],
                textposition='top center'
            ))

        # Ligne de moyenne globale
        fig1.add_trace(go.Scatter(
            x=moyennes_par_periode[periode_selectionnee].unique(),
            y=[moyenne_globale] * len(moyennes_par_periode[periode_selectionnee].unique()),
            mode='lines',
            name='Moyenne Globale',
            line=dict(color='red', dash='dash'),
            hoverinfo='skip'
        ))

        fig1.update_layout(
            title=f"Moyenne des rapports d'interventions par opérateur ({periode_selectionnee}) avec ligne de moyenne globale",
            xaxis_title=periode_selectionnee,
            yaxis_title="Moyenne des rapports d'interventions",
            template="plotly_dark"
        )

        st.plotly_chart(fig1)

        # Affichage des tableaux
        col3, col4 = st.columns([2, 3])
        with col3:
            st.write("### Tableau des Moyennes par période et par opérateur")
            styled_df = style_moyennes(moyennes_par_operateur)
            st.dataframe(styled_df, use_container_width=True)
            st.markdown("""
            #### Légende :
            - <span style="display: inline-block; width: 12px; height: 12px; background-color: gold; margin-right: 5px; border: 1px solid black;"></span> **Top 3**
            - <span style="display: inline-block; width: 12px; height: 12px; background-color: lightgreen; margin-right: 5px; border: 1px solid black;"></span> **Supérieur à la moyenne**
            - <span style="display: inline-block; width: 12px; height: 12px; background-color: lightpink; margin-right: 5px; border: 1px solid black;"></span> **Inférieur à la moyenne**
            - <span style="display: inline-block; width: 12px; height: 12px; background-color: lightcoral; margin-right: 5px; border: 1px solid black;"></span> **Flop 5**
            """, unsafe_allow_html=True)
            styled_df = style_moyennes(repetitions_graph)
            st.dataframe(styled_df, use_container_width=True)
        with col4:
            st.write("### Tableau des rapports d'intervention par période et par opérateur")
            st.dataframe(repetitions_tableau, use_container_width=True)

# Tirage au sort
        st.subheader("Tirage au sort de deux lignes par opérateur")
        df_filtre = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]
        for operateur in operateurs_selectionnes:
            st.write(f"### Tirage pour {operateur}:")
            df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
            lignes_tirees = df_operateur.sample(n=min(3, len(df_operateur)))
            if not lignes_tirees.empty:
                for _, ligne in lignes_tirees.iterrows():
                    col_info, col_photo = st.columns([3, 1])

                    with col_info:
                        st.markdown(f"""
                        **Date**: {ligne['Date et Heure début d\'intervention']}
                        **Opérateur**: {ligne['Prénom et nom']}
                        **Équipement**: {ligne['Équipement']}
                        **Localisation**: {ligne['Localisation']}
                        **Problème**: {ligne['Technique'] if pd.notna(ligne['Technique']) else ligne['Opérationnel']}
                        """)

                    with col_photo:
                        if pd.notna(ligne['Photo']):
                            st.image(ligne['Photo'], width=200)
                        else:
                            st.write("Pas de photo disponible")
            else:
                st.write("Pas de données disponibles pour cet opérateur dans la période sélectionnée.")
