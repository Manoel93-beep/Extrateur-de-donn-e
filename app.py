import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 1. Configuration de la page
st.set_page_config(page_title="Extracteur de Données 1xBet", page_icon="⚽", layout="wide")

# 2. Injection CSS (Correction du TypeError)
st.markdown("""
    <style>
    .main-title {
        font-size: 28px;
        font-weight: bold;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">⚽ Extracteur de Données 1xBet vers Excel</div>', unsafe_allow_html=True)
st.write("Glissez votre fichier PDF exporté de 1xBet pour générer automatiquement votre tableau Excel.")

# 3. Fonction de parsing algorithmique robuste
def extraire_matchs_depuis_pdf(pdf_file):
    matchs = []
    competition_actuelle = "Compétition Inconnue"
    
    # Mots-clés pour identifier une ligne de compétition
    mots_cles_comp = ["championnat", "ligue", "coupe", "copa", "liga", "euro", "w-league", "u19", "division", "matches amicaux", "6x6"]
    # Éléments parasites à exclure pour éviter les décalages de colonnes
    lignes_a_ignorer = ["téléphone", "inscription", "connexion", "bonus", "combiné du jour", "cote globale", "informations", "qui sommes", "paris", "jeux", "en direct", "avant-match", "télécharger"]

    lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    lines.append(line.strip())

    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Ignorer les lignes vides ou les éléments d'interface utilisateur
        if not line or any(ignore in line.lower() for ignore in lignes_a_ignorer):
            i += 1
            continue
            
        # Détection et nettoyage d'un en-tête de compétition
        if any(keyword in line.lower() for keyword in mots_cles_comp) and not re.search(r'\d{2}/\d{2}', line):
            # Supprime les résidus de colonnes collés en fin de ligne (ex: "1 X 2 1X 12 2X +7")
            clean_comp = re.sub(r'\s+1\s+X\s+2.*$', '', line, flags=re.IGNORECASE)
            competition_actuelle = clean_comp.strip()
            i += 1
            continue
            
        # Détection de la ligne temporelle (Coup d'envoi)
        # Gère les formats : 27/05/01:00, 27/05 10:45, ou les variantes d'extraction
        date_match = re.search(r'(\d{2}/\d{2}(?:/\d{2})?[\s\:]?\d{2}\:\d{2})|(\d{2}/\d{2})', line)
        
        if date_match:
            coup_denvoi = date_match.group(0)
            
            # Récupération des équipes (situées sur les lignes immédiatement précédentes)
            equipe1 = ""
            equipe2 = ""
            candidates = []
            
            j = i - 1
            while j >= 0 and len(candidates) < 2:
                prev_line = lines[j]
                # On valide que la ligne précédente n'est ni une compétition ni une ligne système parasite
                if prev_line and not any(k in prev_line.lower() for k in mots_cles_comp) and not any(ig in prev_line.lower() for ig in lignes_a_ignorer):
                    candidates.append(prev_line)
                j -= 1
                
            if len(candidates) >= 2:
                equipe2 = candidates[0]
                equipe1 = candidates[1]
            elif len(candidates) == 1:
                equipe1 = candidates[0]
                equipe2 = "À déterminer"
                
            # Extraction dynamique des cotes (1 / X / 2) sur les lignes suivantes
            cotes_trouvees = []
            k = i + 1
            while k < len(lines) and len(cotes_trouvees) < 3:
                next_line = lines[k]
                # Capture des valeurs décimales des cotes
                decimals = re.findall(r'\b\d+[\.,]\d+\b', next_line)
                if decimals:
                    cotes_trouvees.extend(decimals)
                elif any(keyword in next_line.lower() for keyword in mots_cles_comp):
                    # Si on bascule sur un en-tête de compétition, on stoppe la recherche de cotes
                    break
                k += 1
                
            cotes_str = " / ".join(cotes_trouvees[:3]) if cotes_trouvees else "-"
            
            # Validation finale pour éviter d'insérer des en-têtes comme noms d'équipes
            if equipe1 and equipe2 and not any(x in equipe1.lower() for x in ["1 x 2", "championnat", "ligue"]):
                matchs.append({
                    "🏠 Équipe domicile": equipe1,
                    "✈️ Équipe extérieur": equipe2,
                    "🏆 Compétition": competition_actuelle,
                    "🕒 Coup d'envoi": coup_denvoi,
                    "💰 Cotes": cotes_str
                })
        i += 1
        
    return pd.DataFrame(matchs)

# 4. Interface utilisateur Streamlit
uploaded_file = st.file_uploader("Choisissez le fichier PDF 1xBet", type="pdf")

if uploaded_file is not None:
    with st.spinner("Analyse et extraction des lignes de paris..."):
        df_matchs = extraire_matchs_depuis_pdf(uploaded_file)
        
    if not df_matchs.empty:
        st.success(f"🎉 {len(df_matchs)} matchs extraits avec succès !")
        
        # Affichage du DataFrame aligné avec vos en-têtes cibles
        st.dataframe(df_matchs, use_container_width=True)
        
        # Génération du fichier Excel téléchargeable
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_matchs.to_excel(writer, index=False, sheet_name='Lignes 1xBet')
        
        st.download_button(
            label="📥 Télécharger le tableau Excel",
            data=buffer.getvalue(),
            file_name="recap_matchs_1xbet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Aucune donnée n'a pu être extraite. Vérifiez que le PDF provient bien d'un export de ligne standard.")
