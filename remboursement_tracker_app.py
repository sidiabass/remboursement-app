try:
    import streamlit as st
except ModuleNotFoundError:
    raise ImportError("Le module 'streamlit' est requis pour exécuter cette application. Veuillez l'installer avec 'pip install streamlit'.")

import pandas as pd
import datetime

def load_data():
    try:
        df_agents = pd.read_excel("agents_data.xlsx")
    except:
        df_agents = pd.DataFrame(columns=["Matricule", "Nom Agent", "Affaire", "Bureau de Poste", "Montant Total"])
    try:
        df_remb = pd.read_excel("remboursements.xlsx")
    except:
        df_remb = pd.DataFrame(columns=["Date", "Matricule", "Montant Versé"])
    return df_agents, df_remb

def save_data(df_agents, df_remb):
    df_agents.to_excel("agents_data.xlsx", index=False)
    df_remb.to_excel("remboursements.xlsx", index=False)

st.set_page_config(page_title="Suivi des Remboursements", layout="wide")
st.title("Logiciel de Suivi des Remboursements par Matricule")

df_agents, df_remb = load_data()
tabs = st.tabs(["Ajouter un Agent", "Saisir un Remboursement", "Synthèse"])

with tabs[0]:
    st.subheader("Ajout d'un nouvel agent")
    with st.form("ajouter_agent"):
        matricule = st.text_input("Matricule (ex: 609797J)")
        nom = st.text_input("Nom de l'agent")
        affaire = st.text_input("Affaire concernée")
        bureau = st.text_input("Bureau de poste")
        montant = st.number_input("Montant total du préjudice", min_value=0.0, step=1000.0)
        submit = st.form_submit_button("Ajouter")

        if submit and matricule:
            new_agent = pd.DataFrame([[matricule.upper(), nom, affaire, bureau, montant]], columns=df_agents.columns)
            df_agents = pd.concat([df_agents, new_agent], ignore_index=True)
            save_data(df_agents, df_remb)
            st.success("Agent ajouté avec succès")

with tabs[1]:
    st.subheader("Saisie de remboursement")
    with st.form("remboursement"):
        matricule_r = st.selectbox("Sélectionner un matricule", df_agents["Matricule"].unique())
        montant_verse = st.number_input("Montant versé", min_value=0.0, step=1000.0)
        date_r = st.date_input("Date du remboursement", datetime.date.today())
        submit_r = st.form_submit_button("Enregistrer")

        if submit_r:
            new_remb = pd.DataFrame([[date_r, matricule_r, montant_verse]], columns=df_remb.columns)
            df_remb = pd.concat([df_remb, new_remb], ignore_index=True)
            save_data(df_agents, df_remb)
            st.success("Remboursement enregistré")

with tabs[2]:
    st.subheader("Tableau de synthèse")
    df_summary = df_agents.copy()
    rembourse_par_matricule = df_remb.groupby("Matricule")["Montant Versé"].sum().to_dict()
    df_summary["Total Remboursé"] = df_summary["Matricule"].map(rembourse_par_matricule).fillna(0.0)
    df_summary["Reste à Rembourser"] = df_summary["Montant Total"] - df_summary["Total Remboursé"]
    st.dataframe(df_summary)
    st.download_button("Télécharger la synthèse", data=df_summary.to_excel(index=False), file_name="synthese_remboursements.xlsx")
