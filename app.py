import io
import re
from typing import List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel to Contatti", page_icon="📇")

st.title("📇 Excel → Excel Contatti")

OUTPUT_COLUMNS = ["NOME", "TELEFONO"]


def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalize_key(value):
    return clean_text(value)


def normalize_phone(value):
    text = clean_text(value)
    if not text:
        return ""

    if text.startswith("'"):
        text = text[1:].strip()

    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]

    sci_match = re.fullmatch(r"[-+]?\d+(?:\.\d+)?[eE][-+]?\d+", text)
    if sci_match:
        try:
            text = format(float(text), ".0f")
        except Exception:
            pass

    text = text.replace("\u00a0", " ")
    text = re.sub(r"[()\-./]", "", text)
    text = re.sub(r"\s+", "", text)

    return text


def format_name(cod_esterno, debitore, lotto):
    cod_esterno = clean_text(cod_esterno)
    debitore = clean_text(debitore)
    lotto = clean_text(lotto)

    left = cod_esterno
    right = f"{debitore}{lotto}"

    if left and right:
        return f"{left} - {right}"
    return left or right


def read_excel_sheets(file_bytes):
    excel = io.BytesIO(file_bytes)

    try:
        dati = pd.read_excel(excel, sheet_name="Dati", dtype=str, engine="openpyxl")
        excel.seek(0)
        recapiti = pd.read_excel(excel, sheet_name="Recapiti", dtype=str, engine="openpyxl")
    except Exception:
        excel.seek(0)
        dati = pd.read_excel(excel, sheet_name="Dati", dtype=str, engine="xlrd")
        excel.seek(0)
        recapiti = pd.read_excel(excel, sheet_name="Recapiti", dtype=str, engine="xlrd")

    return dati, recapiti


def build_contacts(file_bytes):
    dati, recapiti = read_excel_sheets(file_bytes)

    expected_dati = ["CODICE", "COD. ESTERNO", "DEBITORE", "LOTTO"]
    expected_recapiti = ["PRATICA"]

    missing_dati = [c for c in expected_dati if c not in dati.columns]
    missing_recapiti = [c for c in expected_recapiti if c not in recapiti.columns]

    if missing_dati:
        raise ValueError(f"Nel foglio Dati mancano: {', '.join(missing_dati)}")
    if missing_recapiti:
        raise ValueError(f"Nel foglio Recapiti mancano: {', '.join(missing_recapiti)}")

    if recapiti.shape[1] < 14:
        raise ValueError("Il foglio Recapiti non ha abbastanza colonne: servono almeno fino alla colonna N.")

    recapiti = recapiti.copy()
    dati = dati.copy()

    recapiti["TEL_G"] = recapiti.iloc[:, 6]
    recapiti["TEL_H"] = recapiti.iloc[:, 7]
    recapiti["TEL_I"] = recapiti.iloc[:, 8]
    recapiti["TEL_N"] = recapiti.iloc[:, 13]

    dati["KEY"] = dati["CODICE"].map(normalize_key)
    recapiti["KEY"] = recapiti["PRATICA"].map(normalize_key)

    merged = recapiti.merge(
        dati[["KEY", "COD. ESTERNO", "DEBITORE", "LOTTO"]],
        on="KEY",
        how="inner",
    )

    if merged.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    rows: List[dict] = []

    for _, row in merged.iterrows():
        base_name = format_name(row["COD. ESTERNO"], row["DEBITORE"], row["LOTTO"])
        if not base_name:
            continue

        phones = [
            normalize_phone(row["TEL_G"]),
            normalize_phone(row["TEL_H"]),
            normalize_phone(row["TEL_I"]),
            normalize_phone(row["TEL_N"]),
        ]

        phones = [p for p in phones if p]
        phones = list(dict.fromkeys(phones))

        for i, phone in enumerate(phones, start=1):
            rows.append(
                {
                    "NOME": f"{base_name} n.{i}",
                    "TELEFONO": phone,
                }
            )

    output = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)

    if output.empty:
        return output

    return output.drop_duplicates(subset=OUTPUT_COLUMNS, keep="first").reset_index(drop=True)


def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Contatti")
        ws = writer.sheets["Contatti"]
        ws.column_dimensions["A"].width = 50
        ws.column_dimensions["B"].width = 22
        for cell in ws["B"][1:]:
            cell.number_format = "@"
    output.seek(0)
    return output.getvalue()


uploaded_file = st.file_uploader("Carica file Excel (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        output_df = build_contacts(uploaded_file.getvalue())

        st.success(f"File creato con {len(output_df)} righe")
        st.dataframe(output_df, use_container_width=True)

        st.download_button(
            label="Scarica Excel",
            data=to_excel_bytes(output_df),
            file_name="contatti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Errore: {e}")
else:
    st.info("Carica un file Excel.")

