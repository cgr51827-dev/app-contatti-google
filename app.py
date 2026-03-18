import io
import re
from typing import List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel to Google Contacts CSV", page_icon="📇")

st.title("📇 Excel → Google Contacts CSV")

OUTPUT_COLUMNS = ["Name", "Phone 1 - Value"]
PHONE_COLUMN_INDEXES = [6, 7, 8, 13]  # G H I N


def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalize_phone(value):
    text = clean_text(value)
    if not text:
        return ""

    if text.startswith("'"):
        text = text[1:]

    text = re.sub(r"[()\-./\s]", "", text)

    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]

    return text


def format_name(row):
    return f"{clean_text(row['COD. ESTERNO'])} - {clean_text(row['DEBITORE'])}{clean_text(row['LOTTO'])}"


def read_excel_file(file_bytes):
    excel = io.BytesIO(file_bytes)

    # prima prova xlsx
    try:
        dati = pd.read_excel(excel, sheet_name="Dati", dtype=str, engine="openpyxl")
        excel.seek(0)
        recapiti = pd.read_excel(excel, sheet_name="Recapiti", dtype=str, engine="openpyxl")
    except:
        excel.seek(0)
        dati = pd.read_excel(excel, sheet_name="Dati", dtype=str, engine="xlrd")
        excel.seek(0)
        recapiti = pd.read_excel(excel, sheet_name="Recapiti", dtype=str, engine="xlrd")

    return dati, recapiti


def build_contacts(file_bytes):
    dati, recapiti = read_excel_file(file_bytes)

    recapiti["PHONE1"] = recapiti.iloc[:, 6]
    recapiti["PHONE2"] = recapiti.iloc[:, 7]
    recapiti["PHONE3"] = recapiti.iloc[:, 8]
    recapiti["PHONE4"] = recapiti.iloc[:, 13]

    dati["KEY"] = dati["CODICE"].astype(str)
    recapiti["KEY"] = recapiti["PRATICA"].astype(str)

    merged = recapiti.merge(dati, on="KEY")

    rows = []

    for _, row in merged.iterrows():
        base_name = format_name(row)

        phones = [
            normalize_phone(row["PHONE1"]),
            normalize_phone(row["PHONE2"]),
            normalize_phone(row["PHONE3"]),
            normalize_phone(row["PHONE4"]),
        ]

        phones = [p for p in phones if p]
        phones = list(dict.fromkeys(phones))

        for i, phone in enumerate(phones, 1):
            rows.append({
                "Name": f"{base_name} n.{i}",
                "Phone 1 - Value": phone
            })

    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    return df.drop_duplicates()


def to_csv(df):
    return df.to_csv(index=False).encode("utf-8-sig")


uploaded_file = st.file_uploader("Carica Excel (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        df = build_contacts(uploaded_file.getvalue())

        st.success(f"Creati {len(df)} contatti")
        st.dataframe(df)

        st.download_button("Scarica CSV", to_csv(df), "contatti.csv")

    except Exception as e:
        st.error(str(e))
