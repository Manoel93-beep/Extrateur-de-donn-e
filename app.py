import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="Extraction 1xBet vers Excel", layout="wide")
st.title("⚽ Extracteur de Matchs 1xBet vers Excel")
st.write("Glissez votre fichier PDF exporté de 1xBet pour générer automatiquement votre tableau Excel.")

# Zone de dépôt du fichier
uploaded_file = st.file_uploader("Choisissez le fichier PDF 1xBet", type="pdf")

def extraire_donnees_pdf(pdf_file):
    matchs = []
    current_competition = "Compétition Inconnue"
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
                
            lines = text.split("\n")
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                
                # Détection de la compétition
                if "Championnat" in line or "Coupe" in line or "Ligue" in line or "Copa" in line:
                    current_competition = line
                    i += 1
                    continue
                
                # Recherche d'un bloc de match (Date/Heure suivie des équipes et des cotes)
                # Format type: 27/05 01:00 ou 27/05/01:00
                match_date = re.search(r'(\d{2}/\d{2})[\s/:]+(\d{2}:\d{2})', line)
                
                if match_date and (i + 2) < len(lines):
                    date_str = f"{match_date.group(1)} {match_date.group(2)}"
                    
                    # Les lignes suivantes contiennent généralement les équipes et les cotes
                    equipe_dom = lines[i+1].strip()
                    equipe_ext = lines[i+2].strip()
                    
                    # Extraction des cotes (on cherche des nombres décimaux isolés dans les lignes proches)
                    all_numbers = []
                    for offset in range(1, 5):
                        if i + offset < len(lines):
                            found = re.findall(r'\b\d+\.\d+\b', lines[i+offset])
                            all_numbers.extend(found)
                    
                    # On prend les 3 premières cotes distinctes trouvées (1, X, 2)
                    cotes = " / ".join(all_numbers[:3]) if len(all_numbers) >= 3 else "N/A"
                    
                    matchs.append({
                        "🏠 Équipe domicile": equipe_dom,
                        "✈️ Équipe extérieur": equipe_ext,
                        "🏆 Compétition": current_competition,
                        "🕒 Coup d'envoi": date_str,
                        "💰 Cotes (1 / X / 2)": cotes
                    })
                    i += 3
                    continue
                i += 1
    return matchs

if uploaded_file is not None:
    with st.spinner("Extraction des données en cours..."):
        donnees = extraire_donnees_pdf(uploaded_file)
        
        if donnees:
            df = pd.DataFrame(donnees)
            st.success(f"{len(df)} matchs extraits avec succès !")
            
            # Affichage de l'aperçu
            st.dataframe(df)
            
            # Conversion en Excel en mémoire
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Matchs 1xBet')
            
            # Bouton de téléchargement
            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=buffer.getvalue(),
                file_name="extraction_1xbet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Aucun match n'a pu être extrait. Vérifiez le format du PDF.")