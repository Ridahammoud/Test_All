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

# Fonction pour générer un PDF
def generate_pdf(df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    c.drawString(30, height - 40, "Tableau des répétitions des opérateurs")

    y_position = height - 60
    for i, row in df.iterrows():
        text = f"{row['Prénom et nom']} : {row['Repetitions']}"
        c.drawString(30, y_position, text)
        y_position -= 20

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

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

        # Graphique principal
        fig = px.bar(repetitions_graph, x=periode_selectionnee, y='Repetitions', color=col_prenom_nom, barmode='group')
        st.plotly_chart(fig)

        # Tableau avec légende
        st.markdown("""
        #### Légende :
        - **Top 3** : Jaune
        - **Supérieur à la moyenne** : Vert
        - **Inférieur à la moyenne** : Rose
        - **Très inférieur à la moyenne 5** : Rouge
        """)
        
        styled_df = style_moyennes(repetitions_graph)
        st.dataframe(styled_df, use_container_width=True)
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
