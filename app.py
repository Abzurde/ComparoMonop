import streamlit as st
import pandas as pd
import re
from io import BytesIO
from xlsxwriter.utility import xl_col_to_name

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

    # Préparation des données
    for df in (inv_df, rec_df): df.columns = df.columns.str.strip()
    inv_df['Libelle_nettoye'] = inv_df.get('Libelle', '').apply(lambda x: re.sub(r'^(?:[A-Za-z]\d+\s*)+', '', str(x)).strip())
    rec_df['Libelle_nettoye'] = rec_df.get('Libelle', '').apply(lambda x: re.sub(r'^(?:[A-Za-z]\d+\s*)+', '', str(x)).strip())
    df_inv = inv_df.rename(columns={'Qte inventaire': 'Qty_Inv'})[['Code article', 'Libelle_nettoye', 'Qty_Inv']]
    df_rec = rec_df.rename(columns={'Qte recue (UVC)': 'Qty_Rec'})[['Code article', 'Libelle_nettoye', 'Qty_Rec']]
    merged = pd.merge(df_inv, df_rec, on='Code article', how='outer', suffixes=('_Inv', '_Rec'), indicator='Appartenance')
    merged['Appartenance'] = merged['Appartenance'].map({'both':'Commun','left_only':'Seulement Inventaire','right_only':'Seulement Réception'})
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

    # Séparation
    df_both = merged[merged['Appartenance']=='Commun']
    df_only_inv = merged[merged['Appartenance']=='Seulement Inventaire']
    df_only_rec = merged[merged['Appartenance']=='Seulement Réception']

    # Style front
    def highlight_diff(r): return ['background-color: #fff2ac' if r['Diff']!=0 else '' for _ in r]
    t1,t2,t3 = st.tabs(["Articles communs","Uniquement Inventaire","Uniquement Réception"])
    with t1: st.dataframe(df_both.reset_index(drop=True).style.apply(highlight_diff,axis=1))
    with t2: st.dataframe(df_only_inv.reset_index(drop=True).style.apply(highlight_diff,axis=1))
    with t3: st.dataframe(df_only_rec.reset_index(drop=True).style.apply(highlight_diff,axis=1))

    # Export Excel avec mise en valeur de toute la ligne si Diff != 0
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        for name,df_sheet in [('Articles_communs',df_both),('Inventaire_uniquement',df_only_inv),('Reception_uniquement',df_only_rec)]:
            df_sheet.to_excel(writer, sheet_name=name, index=False)
            wb = writer.book
            ws = writer.sheets[name]
            fmt = wb.add_format({'bg_color':'#FFF2AC'})
            # trouve index de Diff
            idx = df_sheet.columns.get_loc('Diff')
            # transforme en lettre excel
            diff_col = xl_col_to_name(idx)
            # range : de ligne 2 à len+1
            first = 2
            last = len(df_sheet)+1
            # formule: =$C2<>0 appliquée sur toute la ligne
            ws.conditional_format(f"A{first}:{xl_col_to_name(len(df_sheet.columns)-1)}{last}",{
                'type':'formula',
                'criteria':f"=${diff_col}{first}<>0",
                'format':fmt
            })
    buf.seek(0)
    st.download_button("⬇️ Télécharger le rapport Excel (3 feuilles)",buf.read(),"Comparaison.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Veuillez uploader un fichier Excel pour commencer la comparaison.")
