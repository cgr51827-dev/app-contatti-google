import io
import re
from typing import List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel to Google Contacts CSV", page_icon="📇", layout="centered")

st.title("📇 Excel → Google Contacts CSV")
st.write(
    "Carica un file Excel con i fogli **Dati** e **Recapiti** per generare un CSV compatibile con Google Contacts."
)

OUTPUT_COLUMNS = ["Name", "Phone 1 - Value"]

# Colonne Excel fisiche del foglio Recapiti:
# G=6, H=7, I=8, N=13 in indice zero-based
PHONE_COLUMN_INDEXES = [6, 7, 8, 13]


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def normalize_key(value) -> str:
    return clean_text(value)


def format_name(cod_esterno, debitore, lotto) -> str:
    cod_esterno = clean_text(cod_esterno)
    debitore = clean_text(debitore)
    lotto = clean_text(lotto)

    left = cod_esterno
    right = f"{debitore}{lotto}"

    if left and right:
        return f"{left} - {right}"
    return left or right


def normalize_phone(value) -> str:
    text = clean_text(value)
    if not text:
        return ""

    if text.startswith("'"):
        text = text[1:].strip()

    sci_match = re.fullmatch(r"[-+]?\d+(?:\.\d+)?[eE][-+]?\d+", text)
    if sci_match:
        try:
            text = format(float(text), ".0f")
        except Exception:
            pass
    elif re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]

    text = text.replace("\u00a0", " ")
    text = re.sub(r"[()\-./]", "", text)
    text = re.sub(r"\s+", "", text)

    return text


def read_excel_sheets(excel_bytes: bytes):
    excel_file = io.BytesIO(excel_bytes)

    try:
        dati = pd.read_excel(excel_file, sheet_name="Dati", dtype=str, engine="openpyxl")
        excel_file.seek(0)
        recapiti = pd.read_excel(excel_file, sheet_name="Recapiti", dtype=str, engine="openpyxl")
    except Exception:
        excel_file.seek(0)
        dati = pd.read_excel(excel_file, sheet_name="Dati", dtype=str, engine="xlrd")
        excel_file.seek(0)
        recapiti = pd.read_excel(excel_file, sheet_name="Recapiti", dtype=str, engine="xlrd")

    return dati, recapiti


def build_contacts(excel_bytes: bytes) -> pd.DataFrame:
    dati, recapiti = read_excel_sheets(excel_bytes)

    expected_dati = ["CODICE", "COD. ESTERNO", "DEBITORE", "LOTTO"]
    expected_recapiti = ["PRATICA"]

    missing_dati = [c for c in expected_dati if c not in dati.columns]
    missing_recapiti = [c for c in expected_recapiti if c not in recapiti.columns]

    if missing_dati:
        raise ValueError(f"Nel foglio 'Dati' mancano: {', '.join(missing_dati)}")
    if missing_recapiti:
        raise ValueError(f"Nel foglio 'Recapiti' manca: {', '.join(missing_recapiti)}")

    max_idx = recapiti.shape[1] - 1
    invalid_indexes = [i for i in PHONE_COLUMN_INDEXES if i > max_idx]
    if invalid_indexes:
        raise ValueError(
            "Il foglio 'Recapiti' non ha abbastanza colonne per leggere G, H, I, N."
        )

    recapiti = recapiti.copy()
    dati = dati.copy()

    recapiti["PHONE_G"] = recapiti.iloc[:, 6]
    recapiti["PHONE_H"] = recapiti.iloc[:, 7]
    recapiti["PHONE_I"] = recapiti.iloc[:, 8]
    recapiti["PHONE_N"] = recapiti.iloc[:, 13]

    phone_columns = ["PHONE_G", "PHONE_H", "PHONE_I", "PHONE_N"]

    dati["CODICE_KEY"] = dati["CODICE"].map(normalize_key)
    recapiti["PRATICA_KEY"] = recapiti["PRATICA"].map(normalize_key)

    merged = recapiti.merge(
        dati[["CODICE_KEY", "COD. ESTERNO", "DEBITORE", "LOTTO"]],
        left_on="PRATICA_KEY",
        right_on="CODICE_KEY",
        how="inner",
    )

    if merged.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    merged["BASE_NAME"] = merged.apply(
        lambda row: format_name(row["COD. ESTERNO"], row["DEBITORE"], row["LOTTO"]),
        axis=1,
    )

    rows: List[dict] = []

    for _, row in merged.iterrows():
        base_name = clean_text(row["BASE_NAME"])
        if not base_name:
            continue

        phones = []
        for col in phone_columns:
            phone = normalize_phone(row.get(col, ""))
            if phone:
                phones.append(phone)

        phones = list(dict.fromkeys(phones))

        for idx, phone in enumerate(phones, start=1):
            rows.append(
                {
                    "Name": f"{base_name} n.{idx}",
                    "Phone 1 - Value": phone,
                }
            )

    output = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)

    if output.empty:
        return output

    return output.drop_duplicates(subset=OUTPUT_COLUMNS, keep="first").reset_index(drop=True)


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


uploaded_file = st.file_uploader("Carica file Excel (.xls, .xlsx)", type=["xls", "xlsx"])

with st.expander("Regole applicate", expanded=False):
    st.markdown(
        """
- Join tra **Dati.CODICE** e **Recapiti.PRATICA**
- Nome contatto: **COD. ESTERNO - DEBITORELOTTO**
- Una riga per ogni numero trovato nelle colonne Excel **G, H, I, N**
- Suffisso automatico **n.1, n.2, ...**
- Output CSV con colonne **Name** e **Phone 1 - Value**
- Rimozione duplicati
- Lettura dei numeri come testo per non perdere gli zeri iniziali
        """
    )

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.getvalue()
        output_df = build_contacts(file_bytes)

        st.success(f"CSV generato con {len(output_df)} righe.")
        st.dataframe(output_df, use_container_width=True)

        st.download_button(
            label="Scarica CSV",
            data=to_csv_bytes(output_df),
            file_name="google_contacts.csv",
            mime="text/csv",
        )
    except Exception as e:
        st.error(f"Errore: {e}")
else:
    st.info("Carica un file Excel per generare il CSV.")
    except Exception as e:
        st.error(f"Errore: {e}")
else:
    st.info("Carica un file Excel per generare il CSV.")
