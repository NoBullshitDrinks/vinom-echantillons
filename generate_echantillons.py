#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Met à jour la page de saisie des échantillons à partir de TarifVinom_Master.xlsx.
Calé sur la structure réelle du master (feuille "Tarif") :
  Code_Vinistoria | Statut | Region | Sous_region | Couleur | Appellation |
  Cuvee | Domaine | Label | Millesime | Contenance_cl | Tarif_HT | Stock | ...

Règles appliquées :
- Les régions et leur ORDRE viennent du master (colonne Region) = même classement que le tarif.
- Si "Cuvee" est vide, le nom affiché = "Domaine" (cas des châteaux bordelais).
- N° fournisseur = 3 premiers chiffres du Code_Vinistoria ; BDR*/BAL* -> 267 (Bouche du Roi).
  Le nom du fournisseur vient de Liste_des_fournisseurs_au_13_04_2026.xlsx (repli : le domaine).
- Contenance != 75 cl affichée en gras (magnums) et reportée dans le libellé envoyé à Yasmina.

Usage :
    pip install pandas openpyxl
    python generate_echantillons.py
"""
import json, re, sys
from pathlib import Path
import pandas as pd

MASTER      = Path("TarifVinom_Master.xlsx")
SHEET       = "Tarif"
FOURNISSEURS= Path("Liste_des_fournisseurs_au_13_04_2026.xlsx")   # facultatif
AGENTS_XLSX = Path("Liste_agents_Vinom.xlsx")                     # facultatif
HTML        = Path("echantillons.html")   # <-- mettre "index.html" si la page a été renommée

BDR_BAL_NF = 267   # Domaine la Bouche du Roi / Bal du Roi

def nf_of(code):
    c = str(code).strip().upper()
    if c.startswith("BDR") or c.startswith("BAL"): return BDR_BAL_NF
    m = re.match(r"(\d{3})", c)
    return int(m.group(1)) if m else 0

def load_fournisseurs():
    if not FOURNISSEURS.exists(): return {}
    f = pd.read_excel(FOURNISSEURS, dtype=str)
    f["NF"] = pd.to_numeric(f["N° Fourn."], errors="coerce")
    f = f.dropna(subset=["NF"]).drop_duplicates("NF")
    return {int(n): (nom or "") for n, nom in zip(f["NF"], f["Nom 1"])}

def build_references():
    if not MASTER.exists(): sys.exit(f"Introuvable : {MASTER.resolve()}")
    df = pd.read_excel(MASTER, sheet_name=SHEET, dtype=str).fillna("")
    fmap = load_fournisseurs()

    regions = []
    for x in df["Region"]:
        x = str(x).strip()
        if x and x not in regions: regions.append(x)

    refs, seen, anomalies = [], set(), []
    for i, r in df.iterrows():
        code  = str(r["Code_Vinistoria"]).strip()
        cuvee = str(r["Cuvee"]).strip()
        dom   = str(r["Domaine"]).strip()
        if not cuvee and not dom: continue
        key = code if code else f"NOCODE{i}"
        if key in seen: continue
        seen.add(key)

        nf   = nf_of(code)
        four = fmap.get(nf, "") or dom
        if nf == 0 or nf not in fmap: anomalies.append((code, cuvee or dom, dom))

        mill = str(r["Millesime"]).strip()
        base = cuvee if cuvee else dom
        des  = (base + (" " + mill if mill and mill.upper() != "NM" else "")).strip()
        cl   = str(r["Contenance_cl"]).strip().replace(".0", "")
        fmt  = "" if cl in ("75", "") else cl + " cl"

        refs.append({"key": key, "code": code, "des": des, "dom": dom,
                     "app": str(r["Appellation"]).strip(), "coul": str(r["Couleur"]).strip(),
                     "reg": str(r["Region"]).strip() or "Autres",
                     "nf": int(nf), "four": four, "fmt": fmt})

    order = {reg: i for i, reg in enumerate(regions)}
    refs.sort(key=lambda x: (order.get(x["reg"], 99), (x["dom"] or "zzz").upper(), x["des"]))
    return refs, regions, anomalies

def build_agents():
    if not AGENTS_XLSX.exists(): return None
    a = pd.read_excel(AGENTS_XLSX, dtype=str).fillna("")
    a.columns = [str(c).strip() for c in a.columns]
    out = []
    for _, x in a.iterrows():
        code = str(x.get("Code", "")).strip()
        nom  = str(x.get("Nom complet", "")).strip()
        actif = str(x.get("Actif ?", "")).strip().lower()
        if re.fullmatch(r"[A-Z]{3}", code) and code not in ("FRA", "OMB") and actif == "oui":
            out.append({"code": code, "nom": nom})
    return sorted(out, key=lambda x: x["nom"])

def replace_block(html, var, data):
    return re.sub(r"(const %s = )(\[.*?\]);" % var,
                  lambda m: m.group(1) + json.dumps(data, ensure_ascii=False) + ";",
                  html, count=1, flags=re.S)

def main():
    if not HTML.exists(): sys.exit(f"Introuvable : {HTML.resolve()} (renommer HTML= en haut du script ?)")
    html = HTML.read_text(encoding="utf-8")
    refs, regions, anomalies = build_references()
    html = replace_block(html, "REFERENCES", refs)
    html = replace_block(html, "REGION_ORDER", regions)
    ags = build_agents()
    if ags is not None: html = replace_block(html, "AGENTS", ags)
    HTML.write_text(html, encoding="utf-8")

    print(f"OK — {len(refs)} références, {len(regions)} régions"
          + (f", {len(ags)} agents" if ags is not None else "") + f" écrits dans {HTML}")
    sans_code = [r["des"] for r in refs if not r["code"]]
    if sans_code:
        print(f"\n/!\\ {len(sans_code)} référence(s) SANS code article dans le master "
              f"(Yasmina devra les identifier à la main) :")
        for d in sans_code: print("   -", d)
    if anomalies:
        print(f"\n/!\\ {len(anomalies)} référence(s) sans fournisseur identifié "
              f"(repli sur le nom du domaine) :")
        for c, d, dom in anomalies: print(f"   - {c or '(sans code)'} | {d} | {dom}")

if __name__ == "__main__":
    main()
