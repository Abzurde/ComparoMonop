import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Comparateur Inventaire vs Réception", layout="wide")
st.title("🔍 Comparaison Inventaire ↔ Réception")

st.markdown(
    """
    **Téléversez un fichier Excel** contenant deux onglets nommés `Inventaire` et `Reception`.
    Chaque onglet doit comporter au moins les colonnes `Code article`, `Libelle` et la quantité
    (`Qte inventaire` pour Inventaire, `Qte recue (UVC)` pour Réception).
    """
)

uploaded_file = st.file_uploader("📂 Fichier Excel", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
    try:
        inventaire_df = pd.read_excel(uploaded_file, sheet_name="Inventaire", engine='openpyxl')
        reception_df = pd.read_excel(uploaded_file, sheet_name="Reception", engine='openpyxl')
    except Exception as e:
        st.error(f"Erreur de lecture du fichier Excel : {e}\n\n" +
                 "Vérifiez que vos onglets sont nommés exactement `Inventaire` et `Reception`, " +
                 "et que le package `openpyxl` est installé.")
        st.stop()

    # Nettoyage des colonnes
    inventaire_df.columns = inventaire_df.columns.str.strip()
    reception_df.columns = reception_df.columns.str.strip()

    # Nettoyage des libellés
    def nettoyer_libelle(libelle):
        return re.sub(r'^(?:[A-Za-z]\d+\s*)+', '', str(libelle)).strip()

    inventaire_df['Libelle_nettoye'] = inventaire_df.get('Libelle', '').apply(nettoyer_libelle)
    reception_df['Libelle_nettoye'] = reception_df.get('Libelle', '').apply(nettoyer_libelle)

    # Préparation des DataFrames
    df_inv = inventaire_df.rename(columns={'Qte inventaire': 'Qty_Inv'})[['Code article', 'Libelle_nettoye', 'Qty_Inv']]
    df_rec = reception_df.rename(columns={'Qte recue (UVC)': 'Qty_Rec'})[['Code article', 'Libelle_nettoye', 'Qty_Rec']]

    # Fusion avec renommage de l'indicateur
    merged = pd.merge(
        df_inv,
        df_rec,
        on='Code article',
        how='outer',
        suffixes=('_Inv', '_Rec'),
        indicator='Appartenance'
    )
    # Recode les valeurs pour plus de lisibilité
    merged['Appartenance'] = merged['Appartenance'].map({
        'both': 'Commun',
        'left_only': 'Seulement Inventaire',
        'right_only': 'Seulement Réception'
    })

    # Séparation selon Appartenance
    df_both = merged[merged['Appartenance'] == 'Commun']
    df_only_inv = merged[merged['Appartenance'] == 'Seulement Inventaire']
    df_only_rec = merged[merged['Appartenance'] == 'Seulement Réception']

    # Affichage
    tab1, tab2, tab3 = st.tabs(["Articles communs", "Uniquement Inventaire", "Uniquement Réception"])
    with tab1:
        st.dataframe(df_both.reset_index(drop=True))
    with tab2:
        st.dataframe(df_only_inv.reset_index(drop=True))
    with tab3:
        st.dataframe(df_only_rec.reset_index(drop=True))

    # Export Excel avec 3 feuilles
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_both.to_excel(writer, sheet_name='Articles_communs', index=False)
        df_only_inv.to_excel(writer, sheet_name='Inventaire_uniquement', index=False)
        df_only_rec.to_excel(writer, sheet_name='Reception_uniquement', index=False)
    buffer.seek(0)
    excel_data = buffer.read()

    st.download_button(
        label="⬇️ Télécharger le rapport Excel (3 feuilles)",
        data=excel_data,
        file_name="Comparaison_Inventaire_Reception.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Veuillez uploader un fichier Excel pour commencer la comparaison.")
