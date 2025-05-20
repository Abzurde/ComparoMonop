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
for df in (inv_df, rec_df):
    df.columns = df.columns.str.strip()
    df['Libelle_nettoye'] = df.get('Libelle', '').apply(
        lambda x: re.sub(r'^(?:[A-Za-z]\d+\s*)+', '', str(x)).strip()
    )

df_inv = inv_df.rename(columns={'Qte inventaire':'Qty_Inv'})[['Code article','Libelle_nettoye','Qty_Inv']]
df_rec = rec_df.rename(columns={'Qte recue (UVC)':'Qty_Rec'})[['Code article','Libelle_nettoye','Qty_Rec']]

# Convertit quantités en entier
for df_q, col in [(df_inv,'Qty_Inv'), (df_rec,'Qty_Rec')]:
    df_q[col] = pd.to_numeric(df_q[col], errors='coerce').fillna(0).astype(int)

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
merged['Diff'] = merged['Qty_Inv'] - merged['Qty_Rec']

# Filtre regex
regex = st.text_input("Recherche (regex) sur Code ou Libellé", "")
if regex:
    try:
        lib = merged['Libelle_nettoye'] if 'Libelle_nettoye' in merged else pd.Series(['']*len(merged))
        mask = (
            merged['Code article'].astype(str).str.contains(regex, regex=True, na=False) |
            lib.astype(str).str.contains(regex, regex=True, na=False)
        )
        merged = merged[mask]
    except re.error:
        st.warning("Expression régulière invalide.")

# Séparation
df_both = merged[merged['Appartenance']=='Commun']
df_only_inv = merged[merged['Appartenance']=='Seulement Inventaire']
df_only_rec = merged[merged['Appartenance']=='Seulement Réception']

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
        # Applique la règle à chaque ligne
        ws.conditional_format(
            f"A2:{xl_col_to_name(len(df_sheet.columns)-1)}{len(df_sheet)+1}",
            {'type':'formula', 'criteria':f"=${col_letter}2<>0", 'format':fmt}
        )
buf.seek(0)
st.download_button(
    "⬇️ Télécharger le rapport Excel (3 feuilles)",
    buf.read(),
    "Comparaison.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
