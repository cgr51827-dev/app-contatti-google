import io
import re
import zipfile
from typing import List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel to Google Contacts CSV", page_icon="📇", layout="centered")

st.title("📇 Excel → Google Contacts CSV")
st.write(
    "Carica un file Excel con i fogli **Dati** e **Recapiti** per generare un CSV compatibile con Google Contacts."
)

RECAPITI_COLUMNS = ["G", "H", "I", "N"]
OUTPUT_COLUMNS = ["Name", "Phone 1 - Value"]


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
    """
    Mantiene gli zeri iniziali e ripulisce valori tipici Excel.
    """
    text = clean_text(value)
    if not text:
        return ""

    # Rimuove apostrofo iniziale usato spesso in Excel per forzare testo
    if text.startswith("'"):
        text = text[1:].strip()

    # Se Excel ha convertito in float intero tipo 0039... -> 3.9E+11 o 12345.0
    # tentiamo una normalizzazione conservativa.
    sci_match = re.fullmatch(r"[-+]?\d+(?:\.\d+)?[eE][-+]?\d+", text)
    if sci_match:
        try:
            as_int = format(float(text), ".0f")
            text = as_int
        except Exception:
            pass
    elif re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]

    # Teniamo solo spazi, + e cifre; rimuoviamo separatori comuni
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[()\-./]", "", text)
    text = re.sub(r"\s+", "", text)

    return text


def build_contacts(excel_bytes: bytes) -> pd.DataFrame:
    dati = pd.read_excel(excel_bytes, sheet_name="Dati", dtype=str)
    recapiti = pd.read_excel(excel_bytes, sheet_name="Recapiti", dtype=str)

    expected_dati = ["CODICE", "COD. ESTERNO", "DEBITORE", "LOTTO"]
    expected_recapiti = ["PRATICA"]

    missing_dati = [c for c in expected_dati if c not in dati.columns]
    missing_recapiti = [c for c in expected_recapiti if c not in recapiti.columns]
    missing_phone_cols = [c for c in RECAPITI_COLUMNS if c not in recapiti.columns]

    if missing_dati:
        raise ValueError(f"Nel foglio 'Dati' mancano le colonne: {', '.join(missing_dati)}")
    if missing_recapiti:
        raise ValueError(f"Nel foglio 'Recapiti' mancano le colonne: {', '.join(missing_recapiti)}")
    if missing_phone_cols:
        raise ValueError(
            "Nel foglio 'Recapiti' mancano le colonne numeri richieste: " + ", ".join(missing_phone_cols)
        )

    dati = dati.copy()
    recapiti = recapiti.copy()

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
        lambda row: format_name(row["COD. ESTERNO"], row["DEBITORE"], row["LOTTO"]), axis=1
    )

    rows: List[dict] = []

    for _, row in merged.iterrows():
        base_name = clean_text(row["BASE_NAME"])
        if not base_name:
            continue

        phones = []
        for col in RECAPITI_COLUMNS:
            phone = normalize_phone(row.get(col, ""))
            if phone:
                phones.append(phone)

        # no duplicati all'interno della singola pratica/nome
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

    # no duplicati globali esatti
    output = output.drop_duplicates(subset=OUTPUT_COLUMNS, keep="first").reset_index(drop=True)
    return output


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


uploaded_file = st.file_uploader("Carica file Excel (.xls, .xlsx)", type=["xls", "xlsx"])

with st.expander("Regole applicate", expanded=False):
    st.markdown(
        """
- Join tra **Dati.CODICE** e **Recapiti.PRATICA**
- Nome contatto: **COD. ESTERNO - DEBITORELOTTO**
- Una riga per ogni numero trovato nelle colonne **G, H, I, N**
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

        csv_bytes = to_csv_bytes(output_df)
        st.download_button(
            label="Scarica CSV",
            data=csv_bytes,
            file_name="google_contacts.csv",
            mime="text/csv",
        )
    except Exception as e:
        st.error(f"Errore durante l'elaborazione: {e}")
else:
    st.info("Carica un file per generare il CSV.")
