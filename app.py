import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Comparateur Inventaire vs Réception", layout="wide")
st.title("🔍 Comparaison Inventaire ↔ Réception")

st.markdown(
    """
    **Téléversez un fichier Excel** contenant deux onglets nommés `Inventaire` et `Reception`.
    Chaque onglet doit comporter au moins les colonnes `Code article`, `Libelle` et la quantité (`Qte inventaire` pour Inventaire, `Qte recue (UVC)` pour Réception).
    """
)

uploaded_file = st.file_uploader("📂 Fichier Excel", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        inventaire_df = xls.parse("Inventaire")
        reception_df = xls.parse("Reception")
    except Exception as e:
        st.error(f"Erreur de lecture du fichier Excel : {e}")
        st.stop()

    # Nettoyer noms colonnes
    inventaire_df.columns = inventaire_df.columns.str.strip()
    reception_df.columns = reception_df.columns.str.strip()

    # Fonction nettoyage libellé
    def nettoyer_libelle(libelle):
        return re.sub(r'^(?:[A-Za-z]\d+\s*)+', '', str(libelle)).strip()

    for df in (inventaire_df, reception_df):
        if 'Libelle' in df.columns:
            df['Libelle_nettoye'] = df['Libelle'].apply(nettoyer_libelle)
        else:
            df['Libelle_nettoye'] = ''

    # Préparation
    inv = inventaire_df.rename(columns={'Qte inventaire':'Qty_Inv'})
    inv = inv[['Code article', 'Libelle_nettoye', 'Qty_Inv']]
    rec = reception_df.rename(columns={'Qte recue (UVC)':'Qty_Rec'})
    rec = rec[['Code article', 'Libelle_nettoye', 'Qty_Rec']]

    # Fusion
    merged = pd.merge(inv, rec, on='Code article', how='outer', suffixes=('_Inv', '_Rec'), indicator=True)

    # Séparations
    both = merged[merged['_merge']=='both'].copy()
    only_inv = merged[merged['_merge']=='left_only'].copy()
    only_rec = merged[merged['_merge']=='right_only'].copy()

    # Affichage dans onglets
    tab1, tab2, tab3 = st.tabs(["Articles communs", "Uniquement Inventaire", "Uniquement Réception"] )

    with tab1:
        st.subheader("Articles présents dans les deux onglets")
        st.dataframe(both.reset_index(drop=True))
    with tab2:
        st.subheader("Articles présents seulement en Inventaire")
        st.dataframe(only_inv.reset_index(drop=True))
    with tab3:
        st.subheader("Articles présents seulement en Réception")
        st.dataframe(only_rec.reset_index(drop=True))

    # Export Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        both.to_excel(writer, sheet_name='Commun', index=False)
        only_inv.to_excel(writer, sheet_name='Inventaire_only', index=False)
        only_rec.to_excel(writer, sheet_name='Reception_only', index=False)
        writer.save()
        processed_data = output.getvalue()

    st.download_button(
        label="⬇️ Télécharger le rapport Excel",
        data=processed_data,
        file_name="Comparaison_Inventaire_Reception.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Veuillez uploader un fichier Excel pour commencer la comparaison.")