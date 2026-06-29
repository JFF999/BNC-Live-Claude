"""
Met à jour les données « Les Affaires » dans le Google Sheet Action_2026-c_New,
à partir de la feuille Surperformance_LesAffaires.

Source (Surperformance_LesAffaires, onglet Feuil1) :
    col A = Date (-> MAJ Aff)   |   col C = Symbole   |   col D = Cours cible (-> Pré Aff)

Destination (Action_2026-c_New), onglets « Portefeuille BNC » et « Prospects » :
    écrit « Pré Aff » et « MAJ Aff », appariés par SYMBOLE (exact, sinon base sans suffixe
    .TO/.V/.NE/.CN — ex. Surperformance « BCE » -> portefeuille « BCE.TO »).

À LANCER EN LOCAL avec le Python pythoncore (voir sync_excel_vers_gsheet.py pour les pré-requis).
"""

import os
import traceback
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread

# ======================== CONFIGURATION ========================
CHEMIN_CRED = r"C:\Users\jfilt\bnc_secrets\compte_service.json"
CHEMIN_LOG = r"G:\My Drive\Actions\bnc_sync_log.txt"
NOM_SOURCE = "Surperformance_LesAffaires"
NOM_DEST = "Action_2026-c_New"

SRC_COL_DATE = 0      # A
SRC_COL_SYMBOLE = 2   # C
SRC_COL_CIBLE = 3     # D
FEUILLES_DEST = ["Portefeuille BNC", "Prospects"]
SUFFIXES_CAD = ('.TO', '.V', '.NE', '.CN')
# ===============================================================


def journal(message):
    horodatage = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S")
    ligne = f"[{horodatage}] {message}"
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


def parse_nombre(valeur):
    """'320,00 $US' / '39,00 $ ' -> 320.0 / 39.0 (sans regex, compatible Arrow)."""
    s = str(valeur)
    for x in ('$', 'US', 'CA', chr(0xa0), chr(0x202f), ' '):
        s = s.replace(x, '')
    s = s.replace(',', '.')
    try:
        return round(float(s), 4)
    except (ValueError, TypeError):
        return None


def base_symbole(s):
    for suf in SUFFIXES_CAD:
        if s.endswith(suf):
            return s[:-len(suf)]
    return s


def trouver(entetes, nom):
    cible = ' '.join(nom.split())
    for i, h in enumerate(entetes):
        if ' '.join(str(h).split()) == cible:
            return i
    return None


def main():
    if not os.path.exists(CHEMIN_CRED):
        journal(f"ERREUR : JSON compte de service introuvable : {CHEMIN_CRED}")
        return

    gc = gspread.service_account(filename=CHEMIN_CRED)

    # 1) Lire la source : map SYMBOLE -> (date, cible)
    src = gc.open(NOM_SOURCE).sheet1
    lignes = src.get_all_values()
    affaires = {}
    for row in lignes[1:]:
        if len(row) <= SRC_COL_CIBLE:
            continue
        sym = str(row[SRC_COL_SYMBOLE]).strip().upper()
        if not sym:
            continue
        cible = parse_nombre(row[SRC_COL_CIBLE])
        date = str(row[SRC_COL_DATE]).strip()
        if cible is not None:
            affaires[sym] = (date, cible)
    journal(f"{len(affaires)} objectifs lus depuis « {NOM_SOURCE} ».")

    # 2) Écrire dans chaque onglet de la destination
    dest = gc.open(NOM_DEST)
    for nom_feuille in FEUILLES_DEST:
        try:
            ws = dest.worksheet(nom_feuille)
        except gspread.exceptions.WorksheetNotFound:
            journal(f"  [ATTENTION] Onglet '{nom_feuille}' absent - ignore.")
            continue

        vals = ws.get_all_values()
        if not vals:
            continue
        entetes = vals[0]
        i_sym = trouver(entetes, 'Symbole')
        i_pa = trouver(entetes, 'Pré Aff')
        i_maj = trouver(entetes, 'MAJ Aff')
        if i_sym is None or i_pa is None:
            journal(f"  [ATTENTION] '{nom_feuille}' : colonnes Symbole/Pré Aff introuvables - ignore.")
            continue

        updates = []
        n_maj = 0
        n_vides = 0
        for r, row in enumerate(vals[1:], start=2):
            if len(row) <= i_sym:
                continue
            sym = str(row[i_sym]).strip().upper()
            if not sym or sym == '0':
                continue
            entree = affaires.get(sym) or affaires.get(base_symbole(sym))
            pa_actuel = str(row[i_pa]).strip() if len(row) > i_pa else ""
            maj_actuel = str(row[i_maj]).strip() if (i_maj is not None and len(row) > i_maj) else ""
            if entree:
                date, cible = entree
                updates.append({'range': gspread.utils.rowcol_to_a1(r, i_pa + 1), 'values': [[cible]]})
                if i_maj is not None and date:
                    updates.append({'range': gspread.utils.rowcol_to_a1(r, i_maj + 1), 'values': [[date]]})
                n_maj += 1
            else:
                # Hors Surperformance (source unique) : Pré Aff / MAJ Aff doivent être VIDES.
                # On n'écrit que si la cellule n'est pas déjà vide (économie d'appels).
                a_vide = False
                if pa_actuel != "":
                    updates.append({'range': gspread.utils.rowcol_to_a1(r, i_pa + 1), 'values': [[""]]})
                    a_vide = True
                if i_maj is not None and maj_actuel != "":
                    updates.append({'range': gspread.utils.rowcol_to_a1(r, i_maj + 1), 'values': [[""]]})
                    a_vide = True
                if a_vide:
                    n_vides += 1

        if updates:
            ws.batch_update(updates, value_input_option='USER_ENTERED')
        journal(f"  [OK] {nom_feuille} : {n_maj} mis a jour, {n_vides} vide(s) (hors Surperformance).")

    journal("Mise a jour Les Affaires terminee avec succes.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        journal(f"ECHEC : {type(e).__name__} - {e}")
        journal(traceback.format_exc())
