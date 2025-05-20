import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Si vous ajoutez le logo dans votre repo Streamlit, placez-le à la racine et nommez-le 'monoprix_logo.png'
LOGO_PATH = 'monoprix_logo.png'

st.set_page_config(page_title="Comparateur Inventaire vs Réception", layout="wide")
# Affiche le logo local
st.image(LOGO_PATH, width=150)
st.title("Comparaison Inventaire ↔ Réception")

st.markdown(
    """
    **Téléversez un fichier Excel** contenant deux onglets nommés `Inventaire` et `Reception`.
    Chaque onglet doit comporter au moins les colonnes `Code article`, `Libelle` et la quantité
    (`Qte inventaire` pour Inventaire, `Qte recue (UVC)` pour Réception).

    Utilisez la **barre de recherche** ci-dessous pour filtrer les articles par expression régulière.
    Les différences de quantités seront mises en évidence dans le tableau et dans le fichier Excel exporté.
    """
)

uploaded_file = st.file_uploader("📂 Fichier Excel", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
    try:
        inv_df = pd.read_excel(uploaded_file, sheet_name="Inventaire", engine='openpyxl')
        rec_df = pd.read_excel(uploaded_file, sheet_name="Reception", engine='openpyxl')
    except Exception as e:
        st.error(f"Erreur de lecture du fichier Excel : {e}\nVérifiez le nom des onglets et que `openpyxl` est installé.")
        st.stop()

    # Nettoyage et préparation
    for df in (inv_df, rec_df):
        df.columns = df.columns.str.strip()
    inv_df['Libelle_nettoye'] = inv_df.get('Libelle', '').apply(lambda x: re.sub(r'^(?:[A-Za-z]\d+\s*)+', '', str(x)).strip())
    rec_df['Libelle_nettoye'] = rec_df.get('Libelle', '').apply(lambda x: re.sub(r'^(?:[A-Za-z]\d+\s*)+', '', str(x)).strip())

    df_inv = inv_df.rename(columns={'Qte inventaire': 'Qty_Inv'})[['Code article', 'Libelle_nettoye', 'Qty_Inv']]
    df_rec = rec_df.rename(columns={'Qte recue (UVC)': 'Qty_Rec'})[['Code article', 'Libelle_nettoye', 'Qty_Rec']]

    merged = pd.merge(
        df_inv, df_rec,
        on='Code article', how='outer', suffixes=('_Inv', '_Rec'),
        indicator='Appartenance'
    )
    merged['Appartenance'] = merged['Appartenance'].map({
        'both': 'Commun',
        'left_only': 'Seulement Inventaire',
        'right_only': 'Seulement Réception'
    })
    merged['Diff'] = merged['Qty_Inv'].fillna(0) - merged['Qty_Rec'].fillna(0)

    # Recherche regex
    regex = st.text_input("Recherche (regex) sur Code article ou Libellé", "")
    if regex:
        try:
            mask = (
                merged['Code article'].astype(str).str.contains(regex, regex=True, na=False) |
                merged['Libelle_nettoye'].astype(str).str.contains(regex, regex=True, na=False)
            )
            merged = merged[mask]
        except re.error:
            st.warning("Expression régulière invalide.")

    # Séparations
    df_both = merged[merged['Appartenance'] == 'Commun']
    df_only_inv = merged[merged['Appartenance'] == 'Seulement Inventaire']
    df_only_rec = merged[merged['Appartenance'] == 'Seulement Réception']

    # Style en front
    def highlight_diff(row):
        return ['background-color: #fff2ac' if row['Diff'] != 0 else '' for _ in row]

    tab1, tab2, tab3 = st.tabs(["Articles communs", "Uniquement Inventaire", "Uniquement Réception"])
    with tab1:
        st.dataframe(df_both.reset_index(drop=True).style.apply(highlight_diff, axis=1))
    with tab2:
        st.dataframe(df_only_inv.reset_index(drop=True).style.apply(highlight_diff, axis=1))
    with tab3:
        st.dataframe(df_only_rec.reset_index(drop=True).style.apply(highlight_diff, axis=1))

    # Export Excel avec mise en évidence via XlsxWriter
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        for sheet_name, df_sheet in [
            ('Articles_communs', df_both),
            ('Inventaire_uniquement', df_only_inv),
            ('Reception_uniquement', df_only_rec)
        ]:
            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]
            # Format pour les cellules en différence
            diff_format = workbook.add_format({'bg_color': '#FFF2AC'})
            # Trouver l'index de la colonne 'Diff'
            header = df_sheet.columns.tolist()
            if 'Diff' in header:
                col_idx = header.index('Diff')
                # Appliquer format conditionnel : sur toute la colonne Diff
                first_row = 1
                last_row  = len(df_sheet)
                worksheet.conditional_format(first_row, col_idx, last_row, col_idx,
                                             {'type': 'cell', 'criteria': '!=', 'value': 0, 'format': diff_format})
    buffer.seek(0)
    st.download_button(
        label="⬇️ Télécharger le rapport Excel (3 feuilles)",
        data=buffer.read(),
        file_name="Comparaison_Inventaire_Reception.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Veuillez uploader un fichier Excel pour commencer la comparaison.")
