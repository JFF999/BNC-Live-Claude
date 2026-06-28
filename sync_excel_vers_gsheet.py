"""
Synchronise les données SOURCE du fichier Excel vers le Google Sheet « Action_2026-c_New ».
Lit les VALEURS en cache des formules (openpyxl/pandas) — donc PAS de #REF! (contrairement
à une conversion Google du .xlsx).

  - onglet "Portefeuille BNC" : colonnes A–J  (source BNC + Pré Aff + MAJ Aff)
  - onglet "Prospects"        : colonnes A–E  (source + Pré Aff + MAJ Aff)

Les colonnes calculées par l'app (Prix $, Pré G %, Pré YF, MAJ YF, …) ne sont jamais touchées.

À LANCER EN LOCAL, avec le Python qui voit le lecteur G: (pythoncore).

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
CHEMIN_XLSX = r"G:\My Drive\Actions\Action_2026-c_New.xlsx"
CHEMIN_CRED = r"C:\Users\jfilt\bnc_secrets\compte_service.json"   # JSON du compte de service (local, hors Drive)
CHEMIN_LOG = r"G:\My Drive\Actions\bnc_sync_log.txt"             # journal des exécutions
NOM_GOOGLE_SHEET = "Action_2026-c_New"

# Onglet -> nombre de colonnes (depuis A) à pousser depuis le xlsx.
FEUILLES = {
    "Portefeuille BNC": 10,  # A–J
    "Prospects": 5,          # A–E
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

    for nom_feuille, nb_cols in FEUILLES.items():
        # 1) Lire les nb_cols premières colonnes du xlsx (en-tête INCLUS = header=None).
        #    pandas lit les VALEURS en cache des formules (pas les formules).
        df = pd.read_excel(
            CHEMIN_XLSX, sheet_name=nom_feuille, engine="openpyxl", header=None
        )
        sous = df.iloc[:, :nb_cols]
        valeurs = [[serialiser(v) for v in ligne] for ligne in sous.itertuples(index=False)]

        # 2) Écrire dans l'onglet correspondant du Google Sheet, à partir de A1
        try:
            ws = classeur.worksheet(nom_feuille)
        except gspread.exceptions.WorksheetNotFound:
            journal(f"  [ATTENTION] Onglet '{nom_feuille}' absent du Google Sheet - ignore.")
            continue

        derniere_lig = len(valeurs)
        plage = f"A1:{lettre_colonne(nb_cols)}{derniere_lig}"
        ws.update(range_name=plage, values=valeurs, value_input_option="USER_ENTERED")
        journal(f"  [OK] {nom_feuille} : {derniere_lig} lignes x {nb_cols} colonnes ({plage}).")

    journal("Synchronisation terminee avec succes.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        journal(f"ÉCHEC : {type(e).__name__} - {e}")
        journal(traceback.format_exc())
