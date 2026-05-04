import streamlit as st
import pandas as pd
import tempfile
import os

# 👉 IMPORT DEINE LOGIK
from generator import run

st.set_page_config(page_title="Probefahrten FW Lauerz", layout="wide")

st.title("🚒 Probefahrten FW Lauerz")

# -------------------------
# Eingaben
# -------------------------

year = st.text_input("Jahr", value="2026")

excel = st.file_uploader("Mannschaftsliste (Excel)", type=["xlsx"])
names = []

if excel:
    df = pd.read_excel(excel)

    names = [
        f"{row['Grad']} {row['Vorname']} {row['Name']}"
        for _, row in df.iterrows()
    ]
template = st.file_uploader("Vorlage (Word)", type=["docx"])

# -------------------------
# Fahrzeuge
# -------------------------

st.subheader("Fahrzeuge - Monate ohne Probefahrten")

months = ["Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"]
disable = []

cols = st.columns(4)

for i, m in enumerate(months):
    if cols[i % 4].checkbox(m):
        disable.append(i)

# -------------------------
# Poseidon
# -------------------------

st.subheader("Poseidon - Monate ohne Probefahrten")

pose_disable = [False]*12
pose_text = [""]*12

for i, m in enumerate(months):
    col1, col2 = st.columns([1,2])

    pose_disable[i] = col1.checkbox(m, key=f"pose_{i}")

    pose_text[i] = col2.text_input(
        "Text",
        key=f"text_{i}",
        label_visibility="collapsed"
    )

# -------------------------
# Grossreinigung
# -------------------------

st.subheader("Poseidon - Grossreinigung")

gross_m = st.selectbox("Monat", months)
gross_d = st.text_input("Datum")

st.markdown("**Personen (optional)**")

gross_persons = []

for i in range(3):
    person = st.selectbox(
        f"Person {i+1}",
        [""] + names,
        key=f"gross_{i}"
    )
    gross_persons.append(person)

# -------------------------
# GENERIEREN
# -------------------------

if st.button("🚀 Generieren"):

    gross_persons = [p for p in gross_persons if p]
    gross_persons = list(dict.fromkeys(gross_persons))

    if not excel or not template:
        st.error("Bitte Excel und Vorlage hochladen")
    else:

        with tempfile.TemporaryDirectory() as tmp:

            excel_path = os.path.join(tmp, "input.xlsx")
            template_path = os.path.join(tmp, "template.docx")

            with open(excel_path, "wb") as f:
                f.write(excel.getbuffer())

            with open(template_path, "wb") as f:
                f.write(template.getbuffer())

            docx = run(
                excel_path,
                template_path,
                year,
                disable,
                {
                    "disable": pose_disable,
                    "text": pose_text,
                    "gross_m": months.index(gross_m),
                    "gross_d": gross_d,
                    "gross_persons": gross_persons
                }
            )

            with open(docx, "rb") as f:
                st.download_button(
                    "📄 DOCX herunterladen",
                    f,
                    file_name="Probefahrten.docx"
                )
