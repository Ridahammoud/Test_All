import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from io import BytesIO
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import plotly.express as px

# Constantes
COLUMN_PRENOM_NOM = 'Prénom et nom'
COLUMN_DATE = 'Date'
PERIODES = ["Jour", "Semaine", "Mois", "Trimestre", "Année"]

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

def style_moyennes(df, top_n=3, bottom_n=5):
    moyenne_totale = df['Repetitions'].mean()
    df_top = df.nlargest(top_n, 'Repetitions')
    df_bottom = df.nsmallest(bottom_n, 'Repetitions')

    def apply_styles(row):
        if row.name in df_top.index:
            return ['background-color: gold; color: black'] * len(row)
        elif row.name in df_bottom.index:
            return ['background-color: lightcoral; color: white'] * len(row)
        elif row['Repetitions'] > moyenne_totale:
            return ['background-color: lightgreen'] * len(row)
        else:
            return ['background-color: lightpink'] * len(row)

    return df.style.apply(apply_styles, axis=1)

# Fonction pour générer un PDF
def generate_pdf(df, filename="tableau.pdf"):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter
    c.drawString(30, height - 40, "Tableau des répétitions des opérateurs")
    y_position = height - 60
    for i, row in df.iterrows():
        text = f"{row[COLUMN_PRENOM_NOM]} : {row['Repetitions']}"
        c.drawString(30, y_position, text)
        y_position -= 20
    c.save()

# Configuration de la page Streamlit
st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")
st.title("📊 Analyse des interventions des opérateurs")

fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    col1, col2 = st.columns([2, 3])

    with col1:
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        operateurs = df_principal[COLUMN_PRENOM_NOM].unique().tolist()
        operateurs.append("Total")
        operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        if "Total" in operateurs_selectionnes:
            operateurs_selectionnes = df_principal[COLUMN_PRENOM_NOM].unique().tolist()
        
        periode_selectionnee = st.selectbox("Choisissez une période", PERIODES)
        
        df_principal[col_date] = pd.to_datetime(df_principal[col_date], errors='coerce')
        date_min = df_principal[col_date].min()
        date_max = df_principal[col_date].max()
        
        if pd.isna(date_min) or pd.isna(date_max):
            st.warning("Certaines dates dans le fichier sont invalides. Elles ont été ignorées.")
            date_min = date_max = None
        
        debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
        fin_periode = st.date_input("Fin de la période", min_value=debut_periode, max_value=date_max, value=date_max)

        # Nouveaux filtres avancés
        date_filter = st.selectbox("Filtre de date", ["Tout", "Derniers 7 jours", "Dernier mois", "Personnalisé"])
        if date_filter == "Personnalisé":
            debut_periode = st.date_input("Début de la période personnalisée", min_value=date_min, max_value=date_max, value=date_min)
            fin_periode = st.date_input("Fin de la période personnalisée", min_value=debut_periode, max_value=date_max, value=date_max)
        
        min_repetitions, max_repetitions = st.slider("Nombre de répétitions", 
                                                     min_value=int(df_principal['Repetitions'].min()), 
                                                     max_value=int(df_principal['Repetitions'].max()), 
                                                     value=(int(df_principal['Repetitions'].min()), int(df_principal['Repetitions'].max())))
        
        categories = st.multiselect("Catégories", df_principal['Categorie'].unique())
        
        search_term = st.text_input("Rechercher")

    if st.button("Analyser"):
        df_principal = df_principal.dropna(subset=[col_date])
        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        # Application des filtres avancés
        df_filtre = df_principal[
            (df_principal[col_date].dt.date >= debut_periode) &
            (df_principal[col_date].dt.date <= fin_periode) &
            (df_principal['Repetitions'] >= min_repetitions) &
            (df_principal['Repetitions'] <= max_repetitions) &
            (df_principal['Categorie'].isin(categories) if categories else True) &
            (df_principal.apply(lambda row: search_term.lower() in str(row).lower(), axis=1) if search_term else True)
        ]

        groupby_cols = [COLUMN_PRENOM_NOM]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)
        
        repetitions_graph = df_filtre[df_filtre[COLUMN_PRENOM_NOM].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        repetitions_tableau = df_filtre[df_filtre[COLUMN_PRENOM_NOM].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')

        with col2:
            fig = go.Figure()
            for operateur in operateurs_selectionnes:
                df_operateur = repetitions_graph[repetitions_graph[COLUMN_PRENOM_NOM] == operateur]
                fig.add_trace(go.Bar(x=df_operateur[periode_selectionnee], y=df_operateur['Repetitions'], name=operateur,
                                     text=df_operateur['Repetitions'], textposition='inside',
                                     hovertemplate='%{y}'))
            
            fig.update_layout(title=f"Nombre de rapports d'intervention (du {debut_periode} au {fin_periode})",
                              xaxis_title=periode_selectionnee, yaxis_title="Répétitions", template="plotly_dark")
            st.plotly_chart(fig)

            moyennes_par_periode = repetitions_graph.groupby([periode_selectionnee, COLUMN_PRENOM_NOM])['Repetitions'].mean().reset_index()
            moyennes_par_operateur = moyennes_par_periode.groupby([COLUMN_PRENOM_NOM])['Repetitions'].mean().reset_index()

            col3, col4 = st.columns([2, 3])

            with col3:
                fig1 = go.Figure()
                colors = px.colors.qualitative.Set1
                for i, operateur in enumerate(operateurs_selectionnes):
                    df_operateur_moyenne = moyennes_par_periode[moyennes_par_periode[COLUMN_PRENOM_NOM] == operateur]
                    fig1.add_trace(go.Scatter(x=df_operateur_moyenne[periode_selectionnee], y=df_operateur_moyenne['Repetitions'],
                                              mode='lines+markers', name=operateur, line=dict(color=colors[i % len(colors)]),
                                              text=df_operateur_moyenne['Repetitions'], textposition='top center'))

                # Ajout de la ligne de moyenne globale
                moyenne_globale = moyennes_par_periode['Repetitions'].mean()
                fig1.add_trace(go.Scatter(
                    x=[moyennes_par_periode[periode_selectionnee].min(), moyennes_par_periode[periode_selectionnee].max()],
                    y=[moyenne_globale, moyenne_globale],
                    mode='lines',
                    name='Moyenne globale',
                    line=dict(color='red', dash='dash'),
                    hoverinfo='y'
                ))

                fig1.update_layout(title=f"Moyenne des répétitions par opérateur ({periode_selectionnee})",
                                   xaxis_title=periode_selectionnee, yaxis_title="Moyenne des répétitions", template="plotly_dark")
                st.plotly_chart(fig1)

            with col4:
                st.write("### Tableau des Moyennes par période et par opérateur")
                styled_df = style_moyennes(moyennes_par_operateur)
                st.dataframe(styled_df, use_container_width=True)

                st.write("### Tableau des rapports d'intervention par période et par opérateur")
                st.dataframe(repetitions_tableau, use_container_width=True)

        st.subheader("Télécharger le tableau des rapports d'interventions")
        xlsx_data = convert_df_to_xlsx(repetitions_tableau)
        st.download_button(label="Télécharger en XLSX", data=xlsx_data, file_name="NombredesRapports.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.subheader("Télécharger le tableau des rapports d'interventions en PDF")
        pdf_filename = "repetitions.pdf"
        generate_pdf(repetitions_tableau, pdf_filename)
        with open(pdf_filename, "rb") as f:
            st.download_button(label="Télécharger en PDF", data=f, file_name=pdf_filename, mime="application/pdf")

        if st.checkbox("Afficher toutes les données"):
            st.dataframe(df_principal)
