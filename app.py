import streamlit as st
import os

# --- IMPORTATION DES MODULES SUR MESURE ---
from outil_facturation import traiter_facturation
from outil_stock import traiter_stock
from outil_dpd import traiter_suivi_colis
from outil_robinson import traiter_robinson

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Portail Web-to-Print", page_icon="🚀", layout="centered")

# --- CSS PERSONNALISÉ ---
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .main-title { font-size: 2.5rem; font-weight: 700; color: #222222; margin-bottom: 0px; }
        .subtitle { font-size: 1.1rem; color: #BA2025; margin-bottom: 30px; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# BARRE LATÉRALE (MENU)
# ==========================================
try:
    st.sidebar.image("logo.png", use_container_width=True)
except:
    st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3714/3714856.png", width=60)

st.sidebar.markdown("### Votre boîte à outils")
st.sidebar.markdown("Sélectionnez le traitement que vous souhaitez effectuer :")

outil = st.sidebar.radio("", [
    "📊 Facturation Mensuelle", 
    "📦 Nettoyage des Stocks",
    "🚚 Extraction Suivi Colis (DPD)",
    "💰 Récap Robinson"
])

st.sidebar.markdown("---")
st.sidebar.caption("🔧 Outils internes - MicroLynx")


# ==========================================
# CORPS DE L'APPLICATION
# ==========================================
st.markdown('<p class="main-title">Portail Web-to-Print</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Automatisation et traitement de vos exports Excel</p>', unsafe_allow_html=True)

if outil == "📊 Facturation Mensuelle":
    st.header("Facturation Mensuelle")
    st.info("💡 **Instructions :** Glissez l'export brut des commandes (ex: `2.xlsx`) pour générer le tableau classé, fusionné et valorisé.")
    
    fichier = st.file_uploader("📂 Importez le fichier Excel de facturation", type=["xlsx"])
    if fichier:
        try:
            with st.spinner("⏳ Traitement en cours..."):
                fichier_final, nb_cmd, nb_manuelles = traiter_facturation(fichier)
            
            st.success("✅ Tableau généré avec succès !")
            col1, col2 = st.columns(2)
            col1.metric("🛒 Commandes uniques", nb_cmd)
            col2.metric("⚠️ À tarifer manuellement", nb_manuelles)
            
            st.download_button(
                label="📥 Télécharger le tableau final", 
                data=fichier_final, 
                file_name="Facturation_Finale.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" 
            )
        except Exception as e:
            st.error(f"Oups ! Ce fichier ne ressemble pas à un export de Facturation. Avez-vous glissé le bon ? 🧐\n\n*(Détail technique : {e})*")

elif outil == "📦 Nettoyage des Stocks":
    st.header("Nettoyage du Rapport de Stock")
    st.info("💡 **Instructions :** Glissez l'export brut des stocks (ex: `Report.xlsx`) pour générer la version épurée et formatée.")
    
    fichier = st.file_uploader("📂 Importez le rapport de stock", type=["xlsx"])
    if fichier:
        try:
            with st.spinner("⏳ Traitement en cours..."):
                fichier_final, nb_articles = traiter_stock(fichier)
            
            st.success("✅ Stock nettoyé avec succès !")
            st.metric("📦 Articles en stock", nb_articles)
            
            st.download_button(
                label="📥 Télécharger l'état des stocks", 
                data=fichier_final, 
                file_name="Stock_Final.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        except Exception as e:
            st.error(f"Oups ! Ce fichier ne ressemble pas à un export de Stock. Avez-vous glissé le bon ? 🧐\n\n*(Détail technique : {e})*")

elif outil == "🚚 Extraction Suivi Colis (DPD)":
    st.header("Extraction des Suivis Colis (DPD)")
    st.info("💡 **Instructions :** Glissez votre tableau de suivi DPD brut puis saisissez le numéro de dossier. L'outil génèrera un tableau ciblé avec des liens de suivi cliquables.")
    
    fichier = st.file_uploader("📂 Importez le fichier de suivi", type=["xlsx"])
    
    if fichier:
        numero_dossier = st.text_input("🎯 Saisissez le numéro de dossier (Votre référence 1) :", placeholder="ex: 2012217")
        
        if st.button("Rechercher le dossier", type="primary"):
            if numero_dossier:
                try:
                    with st.spinner(f"⏳ Recherche du dossier {numero_dossier} en cours..."):
                        fichier_traite, nb_colis = traiter_suivi_colis(fichier, numero_dossier)
                        
                        st.session_state['fichier_dpd_pret'] = fichier_traite
                        st.session_state['dossier_dpd_actuel'] = numero_dossier
                        st.session_state['nb_colis_trouves'] = nb_colis
                except Exception as e:
                    st.error(f"Oups ! Un petit souci technique : {e} 🧐")
            else:
                st.warning("⚠️ Veuillez saisir un numéro de dossier avant de cliquer sur Rechercher.")
                
    if 'fichier_dpd_pret' in st.session_state and 'dossier_dpd_actuel' in st.session_state:
        st.success(f"✅ Dossier {st.session_state['dossier_dpd_actuel']} extrait avec succès !")
        st.metric("🚚 Colis trouvés", st.session_state.get('nb_colis_trouves', 0))
        
        st.download_button(
            label=f"📥 Télécharger le suivi ({st.session_state['dossier_dpd_actuel']})", 
            data=st.session_state['fichier_dpd_pret'], 
            file_name=f"Suivi_Dossier_{st.session_state['dossier_dpd_actuel']}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

elif outil == "💰 Récap Robinson":
    st.header("Récapitulatif Robinson")
    st.info("💡 **Instructions :** Glissez l'export texte brut de votre ERP (.txt) pour générer le tableau valorisé.")
    
    fichier = st.file_uploader("📂 Importez le fichier ERP (TXT)", type=["txt"])
    if fichier:
        try:
            with st.spinner("⏳ Nettoyage et calculs en cours..."):
                fichier_final, nb_lignes = traiter_robinson(fichier)
            
            st.success("✅ Tableau Robinson généré avec succès !")
            st.metric("🧾 Factures traitées", nb_lignes)
            
            st.download_button(
                label="📥 Télécharger le Récapitulatif", 
                data=fichier_final, 
                file_name="Recap_Robinson.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        except Exception as e:
            st.error(f"Oups ! Ce fichier texte ne semble pas valide. Avez-vous glissé le bon export ? 🧐\n\n*(Détail technique : {e})*")
