#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
excel_to_json.py — Convertit work/enriched.xlsx vers public/data/programme_struct_enrichi.json
- JSON trié chronologiquement (datetime_local sinon date+heure)
- Pas de purge par défaut (pas de suppression des séances passées)
- Clé robuste: datetime_local OU date|heure|titre + version + tmdb_id
- Champs exportés: compatibles avec ton HTML (prix, tarif, commentaire, backdrops, etc.)
- FIX: la colonne "prix" est acceptée quelle que soit sa casse ("Prix", "PRIX") et
       les alias "recompense(s)", "récompense(s)" sont aussi supportés.
"""

import json
from pathlib import Path
from datetime import datetime
import argparse
import pandas as pd

# Emplacements
IN_XLSX  = Path("work/enriched.xlsx")
OUT_JSON = Path("public/data/programme.json")

# Champs exportés (garde l’ordre)
FIELDS_TO_KEEP = [
    "datetime_local","date","heure",
    "titre","titre_original","realisateur","acteurs_principaux",
    "genres","duree_min","annee","pays","version",
    "tarif","prix","categorie","commentaire",
    "synopsis","affiche_url","backdrop_url","backdrops",
    "trailer_url","tmdb_id","imdb_id",
    "allocine_url"]

PRIX_ALIASES = ["prix", "recompense", "recompenses", "récompense", "récompenses"]

def safe_str(x):
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

import re

ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
ISO_DT   = re.compile(r"^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(:\d{2})?$")
EU_DATE  = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")

def parse_dt(obj: dict):
    """
    Parse déterministe pour éviter les inversions jour/mois.
    Priorité: datetime_local (ISO), puis date+heure.
    Retourne un pd.Timestamp ou None.
    """
    dl = (obj.get("datetime_local") or "").strip()
    if dl:
        # ISO datetime "YYYY-MM-DD HH:MM" ou "YYYY-MM-DDTHH:MM(:SS)?"
        if ISO_DT.match(dl):
            # Essais explicites avant fallback
            dt = pd.to_datetime(dl.replace("T", " "), format="%Y-%m-%d %H:%M", errors="coerce")
            if pd.isna(dt):
                dt = pd.to_datetime(dl, dayfirst=False, errors="coerce")
            return None if pd.isna(dt) else dt
        # Fallback (historique)
        dt = pd.to_datetime(dl, dayfirst=True, errors="coerce")
        return None if pd.isna(dt) else dt

    d = (obj.get("date") or obj.get("Date") or "").strip()
    h = (obj.get("heure") or obj.get("Heure") or "").strip() or "00:00"

    if not d:
        return None

    # ISO date "YYYY-MM-DD" => parse strict
    if ISO_DATE.match(d):
        dt = pd.to_datetime(f"{d} {h}", format="%Y-%m-%d %H:%M", errors="coerce")
        return None if pd.isna(dt) else dt

    # Européen "DD/MM/YYYY" => dayfirst=True
    if EU_DATE.match(d):
        dt = pd.to_datetime(f"{d} {h}", dayfirst=True, errors="coerce")
        return None if pd.isna(dt) else dt

    # Dernier recours (tolérant mais moins déterministe)
    dt = pd.to_datetime(f"{d} {h}", dayfirst=True, errors="coerce")
    return None if pd.isna(dt) else dt
def base_key(obj: dict) -> str:
    """
    Clé de base: dl|{datetime_local} sinon dht|{date}|{heure}|{titre-lc}
    """
    dl = (obj.get("datetime_local") or "").strip()
    if dl:
        return f"dl|{dl}"
    date_s = (obj.get("date") or "").strip()
    heure_s = (obj.get("heure") or "").strip()
    titre_s = ((obj.get("titre") or obj.get("title") or obj.get("Titre") or "")).strip().lower()
    return f"dht|{date_s}|{heure_s}|{titre_s}"

def make_key(obj: dict) -> str:
    """
    Clé robuste: clé de base + version + tmdb_id pour éviter collisions VO/VF & co.
    """
    k = base_key(obj)
    ver = (obj.get("version") or "").strip().lower()
    tid = (obj.get("tmdb_id") or "").strip()
    if ver:
        k += f"|v:{ver}"
    if tid:
        k += f"|id:{tid}"
    return k

def load_existing() -> list:
    if OUT_JSON.exists():
        try:
            with open(OUT_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            pass
    return []

def row_to_obj(r: pd.Series) -> dict:
    """
    Convertit une ligne pandas -> dict pour JSON, avec gestion case-insensible et alias pour 'prix'.
    """
    # dictionnaire "clé minuscule -> valeur"
    raw = {}
    try:
        raw = {str(k).strip().lower(): r[k] for k in r.index}
    except Exception:
        for k in r.index:
            try:
                raw[str(k).strip().lower()] = r[k]
            except Exception:
                pass

    obj = {}
    for k in FIELDS_TO_KEEP:
        if k == "backdrops":
            val = r.get(k, "")
            if val == "" and "backdrops" not in r and "backdrops" in raw:
                val = raw.get("backdrops", "")
            s = safe_str(val)
            if s and s.lstrip().startswith("["):
                try:
                    obj[k] = json.loads(s)
                except Exception:
                    obj[k] = []
            else:
                obj[k] = []
            continue

        if k == "prix":
            v = r.get("prix", None)
            if v is None:
                for alias in PRIX_ALIASES:
                    if alias in raw:
                        v = raw[alias]
                        break
            obj[k] = safe_str(v)
            continue

        v = r.get(k, None)
        if v is None:
            v = raw.get(k.lower(), "")
        obj[k] = safe_str(v)

    return obj

def drop_past(items: list, mode: str) -> list:
    """
    Supprime les séances passées si demandé.
    mode = 'date' (date<aujourd'hui) ou 'datetime' (dt<maintenant).
    """
    now = pd.Timestamp.now()
    today = now.normalize().date()
    kept = []
    for obj in items:
        dt = parse_dt(obj)
        if dt is None:
            kept.append(obj)
            continue
        if mode == "datetime":
            if dt >= now:
                kept.append(obj)
        else:  # 'date'
            if dt.date() >= today:
                kept.append(obj)
    return kept

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--merge-existing", action="store_true",
                    help="Fusionner avec le JSON existant (sinon on repart uniquement de l'Excel).")
    ap.add_argument("--prune-mode", choices=["none","date","datetime"], default="none",
                    help="Suppression des séances passées: none (défaut), date (< aujourd'hui), datetime (< maintenant).")
    args = ap.parse_args()

    # Base = Excel enrichi
    if not IN_XLSX.exists():
        raise SystemExit(f"[ERREUR] {IN_XLSX} introuvable.")
    df = pd.read_excel(IN_XLSX, sheet_name=0, dtype=str).fillna("")

    merged: dict[str, dict] = {}

    # Optionnel: partir du JSON existant
    if args.merge_existing:
        for x in load_existing():
            merged[ make_key(x) ] = x

    # Ajoute/remplace avec l'Excel
    for _, r in df.iterrows():
        obj = row_to_obj(r)
        merged[ make_key(obj) ] = obj

    items = list(merged.values())

    # Purge éventuelle (par défaut: none)
    if args.prune_mode != "none":
        items = drop_past(items, mode=args.prune_mode)

    # Tri chronologique
    def sort_key(o):
        dt = parse_dt(o)
        return dt.to_pydatetime() if dt is not None else datetime(9999,1,1)
    items.sort(key=sort_key)

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)

    print(f"[done] écrit: {OUT_JSON}  ({len(items)} séances)")
    print(f"[info] options: merge_existing={args.merge_existing} prune_mode={args.prune_mode}")

if __name__ == "__main__":
    main()
