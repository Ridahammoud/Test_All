import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from io import BytesIO
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

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

# Appliquer des styles pour les top n et bottom n
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

# Fonction pour le tirage au sort
def tirage_au_sort(df, debut_periode, fin_periode):
    df_filtre = df[(df['Date'] >= debut_periode) & (df['Date'] <= fin_periode)]
    return df_filtre.sample(n=2)

# Configuration de la page Streamlit
st.set_page_config(page_title="Analyse des Interventions", page_icon="📊", layout="wide")
st.title("📊 Analyse des interventions des opérateurs")

fichier_principal = st.file_uploader("Choisissez le fichier principal (donnee_Aesma.xlsx)", type="xlsx")

if fichier_principal is not None:
    df_principal = charger_donnees(fichier_principal)
    
    col1, col2 = st.columns([2, 2])  # Modifier les tailles des colonnes pour donner plus d'espace aux graphiques
    
    with col1:
        col_prenom_nom = df_principal.columns[4]
        col_date = st.selectbox("Choisissez la colonne de date", df_principal.columns)
        
        operateurs = df_principal[col_prenom_nom].unique().tolist()
        operateurs.append("Total")  # Ajout de l'option "Total"
        operateurs_selectionnes = st.multiselect("Choisissez un ou plusieurs opérateurs", operateurs)
        
        # Si "Total" est sélectionné, on inclut tous les opérateurs
        if "Total" in operateurs_selectionnes:
            operateurs_selectionnes = df_principal[col_prenom_nom].unique().tolist()  # Utilise tous les opérateurs disponibles
        
        periodes = ["Jour", "Semaine", "Mois", "Trimestre", "Année"]
        periode_selectionnee = st.selectbox("Choisissez une période", periodes)
        
        df_principal[col_date] = pd.to_datetime(df_principal[col_date], errors='coerce')
        
        date_min = df_principal[col_date].min()
        date_max = df_principal[col_date].max()

        if pd.isna(date_min) or pd.isna(date_max):
            st.warning("Certaines dates dans le fichier sont invalides. Elles ont été ignorées.")
            date_min = date_max = None
        
        debut_periode = st.date_input("Début de la période", min_value=date_min, max_value=date_max, value=date_min)
        fin_periode = st.date_input("Fin de la période", min_value=debut_periode, max_value=date_max, value=date_max)
    
    if st.button("Analyser"):
        df_principal = df_principal.dropna(subset=[col_date])

        df_principal['Jour'] = df_principal[col_date].dt.date
        df_principal['Semaine'] = df_principal[col_date].dt.to_period('W').astype(str)
        df_principal['Mois'] = df_principal[col_date].dt.to_period('M').astype(str)
        df_principal['Trimestre'] = df_principal[col_date].dt.to_period('Q').astype(str)
        df_principal['Année'] = df_principal[col_date].dt.year

        df_graph = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]

        groupby_cols = [col_prenom_nom]
        if periode_selectionnee != "Total":
            groupby_cols.append(periode_selectionnee)
        
        repetitions_graph = df_graph[df_graph[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        repetitions_tableau = df_principal[df_principal[col_prenom_nom].isin(operateurs_selectionnes)].groupby(groupby_cols).size().reset_index(name='Repetitions')
        
        with col2:
            # Graphique principal (barres)
            fig = go.Figure()

            # Ajout des données du graphique de répétitions
            for operateur in operateurs_selectionnes:
                df_operateur = repetitions_graph[repetitions_graph[col_prenom_nom] == operateur]
                fig.add_trace(go.Bar(x=df_operateur[periode_selectionnee],
                                     y=df_operateur['Repetitions'],
                                     name=operateur,
                                     text=df_operateur['Repetitions'],
                                     textposition='auto'))  # Afficher les valeurs sur les barres
            
            fig.update_layout(title=f"Nombre de rapports d'intervention (du {debut_periode} au {fin_periode})",
                              xaxis_title=periode_selectionnee,
                              yaxis_title="Répetitions",
                              template="plotly_dark")
            st.plotly_chart(fig)

        # Calcul des moyennes par opérateur et par période
        moyennes_par_periode = repetitions_graph.groupby([periode_selectionnee, col_prenom_nom])['Repetitions'].mean().reset_index()

        with col1:
            # Graphique des moyennes par opérateur
            fig1 = go.Figure()

            for operateur in operateurs_selectionnes:
                df_operateur_moyenne = moyennes_par_periode[moyennes_par_periode[col_prenom_nom] == operateur]
                fig1.add_trace(go.Scatter(x=df_operateur_moyenne[periode_selectionnee], 
                                          y=df_operateur_moyenne['Repetitions'], 
                                          mode='lines+markers', 
                                          name=operateur,
                                          text=df_operateur_moyenne['Repetitions'],
                                          textposition='top center'))  # Afficher les valeurs des moyennes
            
            moyenne_totale = moyennes_par_periode['Repetitions'].mean()
            fig1.add_trace(go.Scatter(x=moyennes_par_periode[periode_selectionnee], 
                                      y=[moyenne_totale] * len(moyennes_par_periode), 
                                      mode='lines', 
                                      name="Moyenne Totale", 
                                      line=dict(dash='dash', color='red')))

            fig1.update_layout(title=f"Moyennes par opérateur (du {debut_periode} au {fin_periode})",
                               xaxis_title=periode_selectionnee,
                               yaxis_title="Moyenne des répétitions",
                               template="plotly_dark")
            st.plotly_chart(fig1)

            # Affichage du tableau des moyennes par opérateur
            st.write("### Tableau des Moyennes par période et par opérateur")
            styled_df = style_moyennes(moyennes_par_periode)
            st.dataframe(styled_df, use_container_width=True)

        # Affichage du tableau des répétitions
        st.write("### Tableau des répétitions par opérateur et période")
        st.dataframe(repetitions_graph)

        # Tirage au sort de deux lignes du fichier
        st.subheader("Tirage au sort de deux lignes par opérateur")
        df_filtre = df_principal[(df_principal[col_date].dt.date >= debut_periode) & (df_principal[col_date].dt.date <= fin_periode)]
        for operateur in operateurs_selectionnes:
            st.write(f"Tirage pour {operateur}:")
            df_operateur = df_filtre[df_filtre[col_prenom_nom] == operateur]
            lignes_tirees = df_operateur.sample(n=min(2, len(df_operateur)))
            if not lignes_tirees.empty:
                lignes_tirees['Photo'] = lignes_tirees['Photo'].apply(lambda x: f'<img src="{x}" width="100"/>')
                lignes_tirees['Photo 2'] = lignes_tirees['Photo 2'].apply(lambda x: f'<img src="{x}" width="100"/>')
                st.markdown(lignes_tirees.to_html(escape=False), unsafe_allow_html=True)
            else:
                st.write("Aucune ligne disponible pour ce tirage.")
            
            st.write("---")

        st.subheader("Télécharger le tableau des rapports d'interventions")
        xlsx_data = convert_df_to_xlsx(repetitions_tableau)
        st.download_button(label="Télécharger en XLSX", data=xlsx_data, file_name="NombredesRapports.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.subheader("Télécharger le tableau des rapports d'interventions en PDF")
        pdf_filename = "repetitions.pdf"
        generate_pdf(repetitions_tableau, pdf_filename)
        with open(pdf_filename, "rb") as f:
            st.download_button(label="Télécharger en PDF", data=f, file_name=pdf_filename, mime="application/pdf")

    if st.checkbox("Afficher toutes les données"):
        st.dataframe(df_principal)
