import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np

# Stockage des données dans l'état de session
if 'original_data' not in st.session_state:
    st.session_state.original_data = None
if 'header_rows' not in st.session_state:
    st.session_state.header_rows = None

def find_header_row(df, required_terms):
    for index, row in df.iterrows():
        if all(term in row.astype(str).values for term in required_terms):
            return index
    return None

def process_excel(file):
    try:
        df = pd.read_excel(file, header=None)
        required_terms = ['Code', 'Nom', 'Prénom']
        header_row_index = find_header_row(df, required_terms)

        if header_row_index is None:
            st.error("La ligne contenant 'Code', 'Nom' et 'Prénom' n'a pas été trouvée.")
            return None, None

        # Séparation entête et données
        rows_before_header = df.iloc[:header_row_index]
        df_data = df.iloc[header_row_index:]

        # Traitement dataframe principal
        df_data.columns = df_data.iloc[0]
        df_data = df_data[1:].reset_index(drop=True)
        df_data.rename(columns={'Code': 'A:Code'}, inplace=True)
        
        if df_data.empty:
            st.error("Fichier Excel vide après traitement.")
            return None, None

        # Conversion des codes étudiants en numérique
        df_data['A:Code'] = pd.to_numeric(df_data['A:Code'], errors='coerce')
        invalid_codes = df_data[df_data['A:Code'].isna()]
        
        if not invalid_codes.empty:
            st.warning(f"{len(invalid_codes)} codes étudiants invalides détectés")

        # Stockage dans l'état de session
        st.session_state.original_data = df_data
        st.session_state.header_rows = rows_before_header

        return df_data, rows_before_header

    except Exception as e:
        st.error(f"Erreur lors du traitement Excel : {str(e)}")
        return None, None

def process_csv(file):
    try:
        df = pd.read_csv(file, delimiter=';', encoding='latin1')
        required_columns = ['A:Code', 'Nom', 'Prénom', 'Note']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"Colonnes manquantes : {', '.join(missing_columns)}")
            return None, None

        # Conversion et validation des données
        df['A:Code'] = pd.to_numeric(df['A:Code'], errors='coerce')
        df['Note'] = pd.to_numeric(df['Note'], errors='coerce')
        
        # Détection d'anomalies
        anomalies = []
        
        # 1. Notes manquantes
        if df['Note'].isna().any():
            anomalies.append(f"{df['Note'].isna().sum()} notes manquantes")
        
        # 2. Notes hors intervalle
        if (df['Note'] < 0).any() or (df['Note'] > 20).any():
            anomalies.append("Notes hors intervalle [0-20]")
        
        # 3. Codes étudiants manquants
        if df['A:Code'].isna().any():
            anomalies.append(f"{df['A:Code'].isna().sum()} codes étudiants invalides")

        # 4. Correspondance avec l'Excel original
        if st.session_state.original_data is not None:
            original_codes = set(st.session_state.original_data['A:Code'])
            csv_codes = set(df['A:Code'].dropna())
            missing_codes = original_codes - csv_codes
            
            if missing_codes:
                anomalies.append(f"{len(missing_codes)} étudiants de l'Excel non présents dans le CSV")

        return df, anomalies

    except Exception as e:
        st.error(f"Erreur lors du traitement CSV : {str(e)}")
        return None, None

def generate_final_excel(cleaned_csv):
    try:
        # Fusion des données
        final_df = st.session_state.original_data.merge(
            cleaned_csv[['A:Code', 'Note']],
            on='A:Code',
            how='left'
        )

        # Reconstitution du fichier original avec les notes
        header_df = st.session_state.header_rows
        final_file = BytesIO()
        
        with pd.ExcelWriter(final_file, engine='openpyxl') as writer:
            if not header_df.empty:
                header_df.to_excel(writer, index=False, header=False)
            final_df.to_excel(writer, index=False, startrow=len(header_df))
        
        final_file.seek(0)
        return final_file

    except Exception as e:
        st.error(f"Erreur lors de la génération du fichier final : {str(e)}")
        return None

# Interface utilisateur
st.title("📚 Système de gestion des notes AMC")

tab1, tab2 = st.tabs(["📤 Préparation Excel", "📥 Traitement CSV"])

with tab1:
    st.header("1. Préparation du fichier étudiant")
    uploaded_excel = st.file_uploader("Importer le fichier Excel original", type="xlsx")
    
    if uploaded_excel:
        processed_data, _ = process_excel(uploaded_excel)
        
        if processed_data is not None:
            st.success("Fichier Excel validé !")
            st.dataframe(processed_data.head())
            
            csv_buffer = BytesIO()
            processed_data.to_csv(csv_buffer, index=False, encoding='latin1')
            st.download_button(
                label="📥 Télécharger le CSV pour AMC",
                data=csv_buffer.getvalue(),
                file_name="liste_etudiants_amc.csv",
                mime="text/csv"
            )

with tab2:
    st.header("2. Intégration des notes AMC")
    
    if st.session_state.original_data is None:
        st.warning("Veuillez d'abord traiter le fichier Excel dans l'onglet 1")
    else:
        uploaded_csv = st.file_uploader("Importer le CSV de notes AMC", type="csv")
        
        if uploaded_csv:
            cleaned_csv, anomalies = process_csv(uploaded_csv)
            
            if cleaned_csv is not None:
                # Affichage des anomalies
                if anomalies:
                    st.error("🚨 Anomalies détectées :")
                    for anomaly in anomalies:
                        st.write(f"- {anomaly}")
                else:
                    st.success("✅ Aucune anomalie critique détectée")
                
                # Affichage des statistiques
                st.subheader("Statistiques des notes")
                col1, col2, col3 = st.columns(3)
                col1.metric("Moyenne générale", f"{cleaned_csv['Note'].mean():.2f}/20")
                col2.metric("Note maximale", f"{cleaned_csv['Note'].max():.2f}/20")
                col3.metric("Note minimale", f"{cleaned_csv['Note'].min():.2f}/20")
                
                # Génération du fichier final
                final_excel = generate_final_excel(cleaned_csv)
                
                if final_excel:
                    st.download_button(
                        label="📥 Télécharger le fichier final",
                        data=final_excel,
                        file_name="notes_finales.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Affichage prévisualisation
                    st.subheader("Aperçu des données fusionnées")
                    preview_df = pd.read_excel(final_excel)
                    st.dataframe(preview_df.head())