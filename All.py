import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from io import BytesIO
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import random

# Fonction de chargement des données
@st.cache_data
def charger_donnees(fichier):
    return pd.read_excel(fichier)

# Fonction pour convertir un dataframe en fichier XLSX
def convert_df_to_xlsx(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Fonction pour générer un PDF
def generate_pdf(df, filename="tableau.pdf"):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter
    c.drawString(30, height - 40, "Tableau des répétitions des opérateurs")
    
    y_position = height - 60
    for i, row in df.iterrows():
        text = f"{row['Prénom et nom']} : {row['Repetitions']}"
        c.drawString(30, y_position, text)
        y_position -= 20
    
    c.save()

# Calcul des moyennes et ajout des styles
def style_moyennes(df, top_n=5, bottom_n=5):
    # Appliquer un style pour les top n et bottom n
    df_top = df.nlargest(top_n, 'Repetitions')
    df_bottom = df.nsmallest(bottom_n, 'Repetitions')

    def apply_styles(row):
        if row.name in df_top.index:
            return ['background-color: lightblue'] * len(row)  # Blue color for top
        elif row.name in df_bottom.index:
            return ['background-color: lightcoral; color: white'] * len(row)  # Red color for bottom
        else:
            return [''] * len(row)  # No style for others

    styled_df = df.style.apply(apply_styles, axis=1)
    return styled_df

# Fonction pour le tirage au sort
def tirage_au_sort(df, debut_periode, fin_periode):
    df_filtre = df[(df['Date'] >= debut_periode) & (df['Date'] <= fin_periode)]
    return df_filtre.sample(n=2)

# Configuration de la page Streamlit
st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")

st.title("📊 Analyse des interventions des opérateurs")

# Chargement du fichier Excel
fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    # Configuration des colonnes et de la disposition du layout
    col1, col2 = st.columns([1, 2])
    
    with col1:
        # Sélection de la colonne pour 'Prénom et nom' et de la colonne de date
        col_prenom_nom = df_principal.columns[4]  # Sélection automatique de la première colonne
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        
        operateurs = df_principal[col_prenom_nom].unique().tolist()  # .tolist() garantit que c'est une liste
        operateurs.append("Total")  # Ajout de l'option "Total"
        operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année", "Total"]
        periode_selectionnee = st.selectbox("Choisissez une période", periodes)
        
        # Tentative de conversion des dates avec gestion des erreurs
        df_principal[col_date] = pd.to_datetime(df_principal[col_date], errors='coerce')
        
        # Gérer les dates invalides et définir les bornes min et max des dates valides
        date_min = df_principal[col_date].min()
        date_max = df_principal[col_date].max()

        # Si les dates sont invalides, on avertit l'utilisateur
        if pd.isna(date_min) or pd.isna(date_max):
            st.warning("Certaines dates dans le fichier sont invalides. Elles ont été ignorées.")
            date_min = date_max = None
        
        debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
        fin_periode = st.date_input("Fin de la période", min_value=debut_periode, max_value=date_max, value=date_max)
    
    # Quand le bouton "Analyser" est cliqué
    if st.button("Analyser"):
        # Filtrage des données pour garder seulement les dates valides
        valid_dates_df = df_principal.dropna(subset=[col_date])
        if len(df_principal) != len(valid_dates_df):
            st.warning(f"{len(df_principal) - len(valid_dates_df)} lignes avec des dates invalides ont été ignorées.")
        
        # Ajout des périodes (Jour, Semaine, Mois, Trimestre, Année)
        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        # Filtrer les données pour la période sélectionnée
        df_graph = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]

        # Si "Total" est sélectionné, on inclut tous les opérateurs
        if "Total" in operateurs_selectionnes:
            operateurs_selectionnes = df_principal[col_prenom_nom].unique().tolist()  # Utilise tous les opérateurs disponibles
        
        # Choisir les colonnes pour grouper les données
        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)
        
        # Calcul des répétitions pour le graphique
        repetitions_graph = df_graph[df_graph[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')

        # Limiter le nombre de données affichées dans le graphique si trop volumineux
        if len(repetitions_graph) > 100:
            repetitions_graph = repetitions_graph.sample(n=100)
        
        # Calcul des moyennes par opérateur et par période
        moyennes_par_periode = repetitions_graph.groupby([periode_selectionnee, col_prenom_nom])['Repetitions'].mean().reset_index()

        # Graphique principal (à gauche)
        with col1:
            fig1 = go.Figure()

            # Courbe des répétitions par période et opérateur
            for operateur in operateurs_selectionnes:
                df_operateur = repetitions_graph[repetitions_graph[col_prenom_nom] == operateur]
                fig1.add_trace(go.Scatter(x=df_operateur[periode_selectionnee], 
                                          y=df_operateur['Repetitions'], 
                                          mode='lines+markers', 
                                          name=operateur))

            fig1.update_layout(title=f"Nombre de répétitions par période (du {debut_periode} au {fin_periode})",
                               xaxis_title=periode_selectionnee,
                               yaxis_title="Répetitions",
                               template="plotly_dark")
            st.plotly_chart(fig1)

        # Graphique des moyennes par opérateur (à droite)
        with col2:
            fig2 = go.Figure()

            # Courbe des moyennes par période et opérateur
            for operateur in operateurs_selectionnes:
                df_operateur_moyenne = moyennes_par_periode[moyennes_par_periode[col_prenom_nom] == operateur]
                fig2.add_trace(go.Scatter(x=df_operateur_moyenne[periode_selectionnee], 
                                          y=df_operateur_moyenne['Repetitions'], 
                                          mode='lines+markers', 
                                          name=operateur))

            moyenne_totale = moyennes_par_periode['Repetitions'].mean()
            fig2.add_trace(go.Scatter(x=moyennes_par_periode[periode_selectionnee], 
                                      y=[moyenne_totale] * len(moyennes_par_periode), 
                                      mode='lines', 
                                      name="Moyenne Totale", 
                                      line=dict(dash='dash', color='red')))
            
            fig2.update_layout(title=f"Moyennes par opérateur (du {debut_periode} au {fin_periode})",
                               xaxis_title=periode_selectionnee,
                               yaxis_title="Moyenne des répétitions",
                               template="plotly_dark")
            st.plotly_chart(fig2)

        # Affichage du tableau des répétitions
        st.write("### Tableau des répétitions par opérateur et période")
        st.dataframe(repetitions_graph)

        # Tirage au sort de deux lignes du fichier
        st.write("### Tirage au sort")
        tirage_resultat = tirage_au_sort(df_principal, debut_periode, fin_periode)
        st.write(tirage_resultat)
