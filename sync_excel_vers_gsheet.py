"""
Synchronise les données SOURCE du fichier Excel vers le Google Sheet « Action_2026-c_New ».
Lit les VALEURS en cache des formules (openpyxl/pandas) — donc PAS de #REF! (contrairement
à une conversion Google du .xlsx).

  - onglet "Portefeuille BNC" : colonnes A–J  (source BNC + Pré Aff + MAJ Aff)
  - onglet "Prospects"        : colonnes A–E  (source + Pré Aff + MAJ Aff)

Les colonnes calculées par l'app (Prix $, Pré G %, Pré YF, MAJ YF, …) ne sont jamais touchées.

À LANCER EN LOCAL, avec le Python pythoncore. Le dossier Google Drive local est détecté
automatiquement selon la langue de Windows (My Drive / Mon Drive / Mon disque) via chemins_bnc.py.

Pré-requis (préparation unique) :
  1. Installer les libs :  python -m pip install gspread pandas openpyxl
  2. Placer le JSON du compte de service à l'emplacement CHEMIN_CRED ci-dessous.
  3. Partager le Google Sheet « Action_2026-c_New » avec l'e-mail du compte de service
     (champ client_email du JSON), en droit Éditeur.
"""

import os
import traceback
from datetime import datetime, date, time
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import gspread

# ======================== CONFIGURATION (à adapter) ========================
from chemins_bnc import dossier_actions   # Windows FR (« Mon Drive ») ou EN (« My Drive »)
CHEMIN_XLSX = os.path.join(dossier_actions(), "Action_2026-c_New.xlsx")
CHEMIN_CRED = r"C:\Users\jfilt\bnc_secrets\compte_service.json"   # JSON du compte de service (local, hors Drive)
CHEMIN_LOG = os.path.join(dossier_actions(), "bnc_sync_log.txt")             # journal des exécutions
NOM_GOOGLE_SHEET = "Action_2026-c_New"

# Onglet -> config : nb de colonnes A.. à pousser, colonne Symbole, et plage des
# colonnes CALCULÉES (écrites par l'app) à effacer si le symbole d'une ligne change.
# Pré Aff / MAJ Aff ne sont PLUS poussés par la synchro aller (ils viennent de
# Surperformance_LesAffaires via sync_affaires_vers_gsheet.py). On ne pousse donc que
# A–H (Portefeuille) / A–C (Prospects). La zone "calc" effacée si le symbole change
# couvre tout ce qui dépend du symbole au-delà de la source : I–Q / D–I (Pré Aff,
# MAJ Aff + colonnes calculées par l'app).
FEUILLES = {
    "Portefeuille BNC": {"nb_cols": 8, "col_symbole": "C", "calc": ("I", "Q")},  # A–H ; Sym=C ; clear I–Q
    "Prospects":        {"nb_cols": 3, "col_symbole": "A", "calc": ("D", "I")},  # A–C ; Sym=A ; clear D–I
}
# ===========================================================================


def journal(message):
    """Affiche et enregistre une ligne horodatée dans le journal."""
    horodatage = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S")
    ligne = f"[{horodatage}] {message}"
    # La console Windows (cp1252) peut planter sur certains caractères : on protège l'affichage.
    try:
        print(ligne)
    except Exception:
        try:
            print(ligne.encode("ascii", "replace").decode("ascii"))
        except Exception:
            pass
    try:
        with open(CHEMIN_LOG, "a", encoding="utf-8") as f:
            f.write(ligne + "\n")
    except Exception:
        pass


def serialiser(valeur):
    """Convertit une cellule pandas en une valeur acceptée par l'API Google Sheets."""
    if pd.isna(valeur):
        return ""
    if isinstance(valeur, time):                       # heure seule (ex. MAJ Aff)
        return valeur.strftime("%H:%M:%S")
    if isinstance(valeur, (pd.Timestamp, datetime, date)):
        return valeur.strftime("%Y-%m-%d")
    if isinstance(valeur, np.integer):
        return int(valeur)
    if isinstance(valeur, np.floating):
        return float(valeur)
    if isinstance(valeur, (str, int, float, bool)):
        return valeur
    return str(valeur)                                  # repli sûr pour tout type inattendu


def lettre_colonne(n):
    """1 -> A, 2 -> B, ... 10 -> J."""
    resultat = ""
    while n > 0:
        n, reste = divmod(n - 1, 26)
        resultat = chr(65 + reste) + resultat
    return resultat


def lettre_vers_index(lettre):
    """A->0, B->1, ... C->2."""
    idx = 0
    for c in lettre:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx - 1


def norm_sym(valeur):
    """Normalise un symbole pour comparaison ('', '0', 'nan' -> vide)."""
    s = str(valeur).strip()
    return "" if s in ("", "0", "nan", "None", "NaN") else s


def main():
    if not os.path.exists(CHEMIN_XLSX):
        journal(f"ERREUR : fichier Excel introuvable : {CHEMIN_XLSX}")
        return
    if not os.path.exists(CHEMIN_CRED):
        journal(f"ERREUR : JSON du compte de service introuvable : {CHEMIN_CRED}")
        return

    gc = gspread.service_account(filename=CHEMIN_CRED)
    classeur = gc.open(NOM_GOOGLE_SHEET)
    journal(f"Google Sheet ouvert : {NOM_GOOGLE_SHEET}")

    for nom_feuille, cfg in FEUILLES.items():
        nb_cols = cfg["nb_cols"]
        idx_sym = lettre_vers_index(cfg["col_symbole"])
        debut_calc, fin_calc = cfg["calc"]

        # 1) Lire les nb_cols premières colonnes du xlsx (en-tête INCLUS = header=None).
        #    pandas lit les VALEURS en cache des formules (pas les formules).
        df = pd.read_excel(
            CHEMIN_XLSX, sheet_name=nom_feuille, engine="openpyxl", header=None
        )
        sous = df.iloc[:, :nb_cols]
        valeurs = [[serialiser(v) for v in ligne] for ligne in sous.itertuples(index=False)]

        try:
            ws = classeur.worksheet(nom_feuille)
        except gspread.exceptions.WorksheetNotFound:
            journal(f"  [ATTENTION] Onglet '{nom_feuille}' absent du Google Sheet - ignore.")
            continue

        # 2) Lire les ANCIENS symboles du Sheet AVANT d'écraser A–J
        actuel = ws.get_all_values()

        # 3) Écrire A.. (source) à partir de A1
        derniere_lig = len(valeurs)
        plage = f"A1:{lettre_colonne(nb_cols)}{derniere_lig}"
        ws.update(range_name=plage, values=valeurs, value_input_option="USER_ENTERED")

        # 4) Si le symbole d'une ligne a CHANGÉ, ses colonnes calculées (K–Q / F–I) sont
        #    périmées (elles appartenaient à l'ancien symbole) -> on les efface.
        a_effacer = []
        for r in range(2, derniere_lig + 1):
            nouv = norm_sym(df.iloc[r - 1, idx_sym]) if idx_sym < df.shape[1] else ""
            anc = norm_sym(actuel[r - 1][idx_sym]) if (r - 1 < len(actuel) and idx_sym < len(actuel[r - 1])) else ""
            if nouv != anc:
                a_effacer.append(f"{debut_calc}{r}:{fin_calc}{r}")
        if a_effacer:
            ws.batch_clear(a_effacer)

        journal(f"  [OK] {nom_feuille} : {derniere_lig} lignes x {nb_cols} col. ; "
                f"{len(a_effacer)} ligne(s) calculee(s) effacee(s) (symbole change).")

    journal("Synchronisation terminee avec succes.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        journal(f"ÉCHEC : {type(e).__name__} - {e}")
        journal(traceback.format_exc())
