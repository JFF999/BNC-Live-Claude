"""
Synchronise les colonnes A–H de Action_2026-c_New.xlsx vers le Google Sheet
« Action_2026-c_New » (mêmes onglets, mêmes colonnes A–H).

À LANCER EN LOCAL, avec le Python qui voit le lecteur G: (pythoncore) :
    & "C:\\Users\\jfilt\\AppData\\Local\\Python\\pythoncore-3.14-64\\python.exe" sync_excel_vers_gsheet.py

Pré-requis (préparation unique) :
  1. Installer les libs :  python -m pip install gspread pandas openpyxl
  2. Avoir le fichier JSON du compte de service Google (voir CHEMIN_CRED ci-dessous).
  3. Partager le Google Sheet « Action_2026-c_New » avec l'e-mail du compte de
     service (client_email du JSON), en droit **Éditeur**.
"""

from datetime import datetime, date
import numpy as np
import pandas as pd
import gspread

# ======================== CONFIGURATION (à adapter) ========================
CHEMIN_XLSX = r"G:\My Drive\Actions\Action_2026-c_New.xlsx"
CHEMIN_CRED = r"G:\My Drive\Actions\compte_service.json"   # <-- ton JSON de compte de service
NOM_GOOGLE_SHEET = "Action_2026-c_New"

# Onglets à synchroniser et nombre de colonnes (A–H = 8) à pousser depuis le xlsx.
# Adapte si un onglet a besoin d'un autre nombre de colonnes.
FEUILLES = {
    "Portefeuille BNC": 8,   # A–H
    "Prospects": 8,          # A–H
}
# ===========================================================================


def serialiser(valeur):
    """Convertit une cellule pandas en une valeur acceptée par l'API Google Sheets."""
    if pd.isna(valeur):
        return ""
    if isinstance(valeur, (pd.Timestamp, datetime, date)):
        return valeur.strftime("%Y-%m-%d")
    if isinstance(valeur, np.integer):
        return int(valeur)
    if isinstance(valeur, np.floating):
        return float(valeur)
    return valeur


def lettre_colonne(n):
    """1 -> A, 2 -> B, ... 8 -> H."""
    resultat = ""
    while n > 0:
        n, reste = divmod(n - 1, 26)
        resultat = chr(65 + reste) + resultat
    return resultat


def main():
    print(f"Connexion au compte de service…")
    gc = gspread.service_account(filename=CHEMIN_CRED)
    classeur = gc.open(NOM_GOOGLE_SHEET)
    print(f"Google Sheet ouvert : « {NOM_GOOGLE_SHEET} »")

    for nom_feuille, nb_cols in FEUILLES.items():
        # 1) Lire les nb_cols premières colonnes du xlsx (en-tête INCLUS = header=None)
        df = pd.read_excel(
            CHEMIN_XLSX, sheet_name=nom_feuille, engine="openpyxl", header=None
        )
        sous = df.iloc[:, :nb_cols]
        valeurs = [[serialiser(v) for v in ligne] for ligne in sous.itertuples(index=False)]

        # 2) Écrire dans l'onglet correspondant du Google Sheet, à partir de A1
        try:
            ws = classeur.worksheet(nom_feuille)
        except gspread.exceptions.WorksheetNotFound:
            print(f"  ⚠ Onglet « {nom_feuille} » absent du Google Sheet — ignoré.")
            continue

        derniere_col = lettre_colonne(nb_cols)
        derniere_lig = len(valeurs)
        plage = f"A1:{derniere_col}{derniere_lig}"

        ws.update(range_name=plage, values=valeurs, value_input_option="USER_ENTERED")
        print(f"  ✓ {nom_feuille} : {derniere_lig} lignes × {nb_cols} colonnes ({plage}).")

    print("Synchronisation terminée.")


if __name__ == "__main__":
    main()
