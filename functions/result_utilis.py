import os
import pandas as pd

# Chemin vers Ram2022.xlsx dans Database
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
RAM_XLSX_PATH = os.path.join(BASE_DIR, "Database", "Ram2022.xlsx")


def load_summary(excel_file: str) -> pd.DataFrame:
    """
    Charge la feuille 'Résumé' et retourne un DataFrame avec Station et Z_CC49.
    """
    df = pd.read_excel(
        excel_file,
        sheet_name="Résumé",
        usecols=["Station", "Z_CC49"]
    )
    df["Station"] = df["Station"].astype(str).str.strip().str.upper()
    return df


def load_station_data(excel_file: str, sheet_name: str) -> pd.DataFrame:
    """
    Charge la feuille d'une station et formate les colonnes Date / Heure et Résultat.
    """
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    if "Date / Heure" in df.columns:
        df["Date / Heure"] = pd.to_datetime(
            df["Date / Heure"],
            format="%d/%m/%Y %Hh%Mm%S",
            dayfirst=True,
            errors="coerce"
        )
    if "Résultat" in df.columns:
        df["Résultat"] = pd.to_numeric(df["Résultat"], errors="coerce")
    return df


def load_ram_info() -> dict:
    """
    Lit Ram2022.xlsx et retourne un mapping :
      { code_station: { 'Z_CC49', 'SITE', 'PHMA (m NGF)', ... } }
    """
    if not os.path.exists(RAM_XLSX_PATH):
        return {}
    df = pd.read_excel(RAM_XLSX_PATH)
    mapping = {}
    for _, row in df.iterrows():
        code = str(row.get("Nom CD50", "")).strip().upper()
        if not code:
            continue
        mapping[code] = {
            "Z_CC49": row.get("Z_CC49"),
            "SITE":   row.get("SITE"),
            "PHMA (m NGF)": row.get("PHMA (m NGF)"),
            "PMVE (m NGF)": row.get("PMVE (m NGF)"),
            "PMME (m NGF)": row.get("PMME (m NGF)"),
            "NM (m NGF)":  row.get("NM (m NGF)")
        }
    return mapping



