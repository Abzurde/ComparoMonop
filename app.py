import streamlit as st
import pandas as pd
import re
from io import BytesIO
from xlsxwriter.utility import xl_col_to_name

# Logo local (ajoutez 'monoprix_logo.png' à la racine du repo)
LOGO_PATH = 'monoprix_logo.png'

st.set_page_config(page_title="Comparateur Inventaire vs Réception", layout="wide")
st.image(LOGO_PATH, width=150)
st.title("Comparaison Inventaire ↔ Réception")

st.markdown(
    """
    Téléversez un fichier Excel avec deux onglets `Inventaire` et `Reception`.
    Recherchez et filtrez par regex, et visualisez les écarts.
    """
)

uploaded_file = st.file_uploader("📂 Fichier Excel", type=["xlsx"])
if not uploaded_file:
    st.info("Veuillez uploader un fichier Excel pour commencer la comparaison.")
    st.stop()

# Lecture
try:
    inv_df = pd.read_excel(uploaded_file, sheet_name="Inventaire", engine='openpyxl')
    rec_df = pd.read_excel(uploaded_file, sheet_name="Reception", engine='openpyxl')
except Exception as e:
    st.error(f"Erreur de lecture du fichier Excel : {e}")
    st.stop()

# Préparation
def clean_df(df, qty_col):
    df = df.copy()
    df.columns = df.columns.str.strip()
    # Nettoyage libellé
    df['Libelle_nettoye'] = df.get('Libelle', '').apply(
        lambda x: re.sub(r'^(?:[A-Za-z]\d+\s*)+', '', str(x)).strip()
    )
    df = df[['Code article', 'Libelle_nettoye', qty_col]].rename(columns={qty_col: qty_col.split()[0]})
    # Convertit en int
    df[qty_col.split()[0]] = pd.to_numeric(df[qty_col.split()[0]], errors='coerce').fillna(0).round().astype(int)
    return df

df_inv = clean_df(inv_df, 'Qte inventaire')
df_inv.rename(columns={'Libelle_nettoye':'Libelle_Inv', 'Qte':'Qty_Inv'}, inplace=True)

df_rec = clean_df(rec_df, 'Qte recue (UVC)')
df_rec.rename(columns={'Libelle_nettoye':'Libelle_Rec', 'Qte':'Qty_Rec'}, inplace=True)

# Fusion
merged = pd.merge(
    df_inv, df_rec,
    on='Code article', how='outer',
    suffixes=('_Inv','_Rec'),
    indicator='Appartenance'
)
merged['Appartenance'] = merged['Appartenance'].map({
    'both':'Commun',
    'left_only':'Seulement Inventaire',
    'right_only':'Seulement Réception'
})
# Calcule et cast Diff
df_merged = merged.copy()
df_merged['Qty_Inv'] = df_merged['Qty_Inv'].fillna(0).astype(int)
df_merged['Qty_Rec'] = df_merged['Qty_Rec'].fillna(0).astype(int)
df_merged['Diff'] = df_merged['Qty_Inv'] - df_merged['Qty_Rec']

# Filtre regex: couvre code et libellés Inv/Rec
regex = st.text_input("Recherche (regex) sur Code ou Libellé", "")
if regex:
    try:
        mask = (
            df_merged['Code article'].astype(str).str.contains(regex, regex=True, na=False) |
            df_merged['Libelle_Inv'].astype(str).str.contains(regex, regex=True, na=False) |
            df_merged['Libelle_Rec'].astype(str).str.contains(regex, regex=True, na=False)
        )
        df_merged = df_merged[mask]
    except re.error:
        st.warning("Expression régulière invalide.")

# Séparation
df_both = df_merged[df_merged['Appartenance']=='Commun']
df_only_inv = df_merged[df_merged['Appartenance']=='Seulement Inventaire']
df_only_rec = df_merged[df_merged['Appartenance']=='Seulement Réception']

# Mise en évidence front
def highlight(r): return ['background-color: #fff2ac' if r['Diff']!=0 else '' for _ in r]

tabs = st.tabs(["Commun","Seulement Inv.","Seulement Rec."])
for tab, df_sheet in zip(tabs, [df_both, df_only_inv, df_only_rec]):
    with tab:
        st.dataframe(df_sheet.reset_index(drop=True).style.apply(highlight, axis=1))

# Export Excel
buf = BytesIO()
with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
    for name, df_sheet in [
        ('Articles_communs', df_both),
        ('Inventaire_uniquement', df_only_inv),
        ('Reception_uniquement', df_only_rec)
    ]:
        df_sheet.to_excel(writer, sheet_name=name, index=False)
        wb = writer.book
        ws = writer.sheets[name]
        fmt = wb.add_format({'bg_color':'#FFF2AC'})
        idx = df_sheet.columns.get_loc('Diff')
        col_letter = xl_col_to_name(idx)
        max_col = xl_col_to_name(len(df_sheet.columns)-1)
        max_row = len(df_sheet) + 1
        ws.conditional_format(
            f"A2:{max_col}{max_row}",
            {'type':'formula', 'criteria':f"=${col_letter}2<>0", 'format':fmt}
        )
buf.seek(0)
st.download_button(
    "⬇️ Télécharger le rapport Excel (3 feuilles)",
    buf.read(),
    "Comparaison.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
