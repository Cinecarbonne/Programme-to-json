#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Normalize Ciné Carbonne Excel (Feuil1) into a clean table (v3: direct column mapping).

Entrée :  input/source.xlsx  (Feuil1)
Sortie :  work/normalised.xlsx
"""

from pathlib import Path
import re
from datetime import datetime, time
from dateutil import parser as dtparser
import pandas as pd

# --- chemins fixes ---
INPUT_PATH  = Path("input/source.xlsx")
OUTPUT_PATH = Path("work/normalized.xlsx")
SHEET_NAME  = "Feuil1"

# --- colonnes du fichier source ---
COL_A, COL_B, COL_C = 0, 1, 2
COL_TITRE = 4      # E
COL_VERSION = 5    # F
COL_CM = 6         # G
COL_PRIX = 7       # H
COL_CATEG = 8      # I
COL_TARIF = 9      # J
COL_COMMENT = 10   # K

WEEKDAYS_FR = {"lundi","mardi","mercredi","jeudi","vendredi","samedi","dimanche"}


def is_weekday_label(x):
    if x is None or (isinstance(x,float) and pd.isna(x)):
        return False
    s = str(x).strip().lower()
    return any(s.startswith(w) for w in WEEKDAYS_FR)


def parse_date_cell(x):
    if pd.isna(x): 
        return None
    if isinstance(x, (pd.Timestamp, datetime)): 
        return pd.to_datetime(x).date()
    s = str(x).strip()
    for dayfirst in (True, False):
        try:
            d = dtparser.parse(s, dayfirst=dayfirst, fuzzy=True)
            return d.date()
        except Exception:
            pass
    return None


def parse_time_cell(x):
    if pd.isna(x): 
        return None
    if isinstance(x, (pd.Timestamp, datetime)):
        t = pd.to_datetime(x).time()
        return t.replace(second=0, microsecond=0)
    if isinstance(x, time):
        return x.replace(second=0, microsecond=0)
    s = str(x).strip()
    m = re.search(r"(\d{1,2})\s*[h:]\s*(\d{1,2})?$", s)
    if m:
        hh = int(m.group(1))
        mm = int(m.group(2)) if m.group(2) is not None else 0
        if 0 <= hh < 24 and 0 <= mm < 60:
            return time(hh, mm)
    try:
        t = dtparser.parse(s).time()
        return t.replace(second=0, microsecond=0)
    except Exception:
        return None


def norm_str(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): 
        return None
    s = str(x).strip()
    return s or None


def normalize_version(v):
    v = (v or "").strip().upper()
    if v == "VOSTFR": 
        return "VOstFR"
    if v in ("VO", "VF"): 
        return v
    return "VF"


def main():
    if not INPUT_PATH.exists():
        raise SystemExit(f"❌ Fichier introuvable : {INPUT_PATH}")

    raw = pd.read_excel(INPUT_PATH, sheet_name=SHEET_NAME, header=None, dtype=object)
    records = []
    current_date = None

    for _, row in raw.iterrows():
        # mise à jour du jour courant
        a, b = row.get(COL_A), row.get(COL_B)
        if is_weekday_label(a) and parse_date_cell(b):
            current_date = parse_date_cell(b)

        # détection séance
        t = parse_time_cell(row.get(COL_C))
        titre = norm_str(row.get(COL_TITRE))
        if current_date and t and titre:
            version = normalize_version(norm_str(row.get(COL_VERSION)))
            cm = norm_str(row.get(COL_CM))
            prix = norm_str(row.get(COL_PRIX))
            categorie = norm_str(row.get(COL_CATEG))
            tarif = norm_str(row.get(COL_TARIF))
            commentaire = norm_str(row.get(COL_COMMENT))

            records.append({
                "Date": current_date.strftime("%Y-%m-%d"),
                "Heure": f"{t.hour:02d}:{t.minute:02d}",
                "Titre": titre,
                "Version": version,
                "CM": cm,
                "Prix": prix,
                "Categorie": categorie,
                "Tarif": tarif,
                "Commentaire": commentaire,
            })

    df = pd.DataFrame(records, columns=[
        "Date","Heure","Titre","Version","CM","Prix","Categorie","Tarif","Commentaire"
    ])

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(OUTPUT_PATH, index=False)
    print(f"✅ Écrit : {OUTPUT_PATH} ({len(df)} lignes)")


if __name__ == "__main__":
    main()
