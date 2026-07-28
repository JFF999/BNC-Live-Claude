"""
Met à jour les données « Les Affaires » dans le Google Sheet Action_2026-c_New,
à partir de l'onglet « LesAffaires » DU MÊME classeur (avant : fichier séparé
Surperformance_LesAffaires).

Source (Action_2026-c_New, onglet LesAffaires) :
    colonnes repérées par NOM d'en-tête : « Date » (-> MAJ Aff), « Symbole »,
    « Cours cible » (-> Pré Aff). Disposition courante : A=Date, B=Symbole,
    C=Description, D=Cours cible (le repérage par nom rend l'ordre indifférent).

Destination (Action_2026-c_New), onglet « Prospects » :
    écrit « Pré Aff » et « MAJ Aff », appariés par clé de symbole unifiée
    (suffixe .TO/.V/.NE/.CN retiré + point de classe -> tiret ; ex. « BBD.B »,
    « BBD-B » et « BBD-B.TO » se rejoignent ; la cible la plus RÉCENTE gagne).

NB : la même synchro est disponible via le bouton 📰 de l'app (streamlit_V8_app.py),
qui importe les helpers de ce module.

À LANCER EN LOCAL avec le Python pythoncore. Pré-requis (une fois) :
    python -m pip install gspread
    + le JSON du compte de service dans C:\\Users\\jfilt\\bnc_secrets\\compte_service.json
"""

import os
import traceback
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import gspread

JOURS_RETENTION_SOURCE = 365   # purge des entrées LesAffaires plus vieilles que ça


def purger_source(ws, jours=JOURS_RETENTION_SOURCE):
    """Retire de l'onglet source les entrées plus vieilles que `jours` (elles seraient
    de toute façon périmées pour le calcul). Les lignes à date illisible sont gardées.
    Écrit d'abord, réduit ensuite (jamais de clear avant écriture). Renvoie le nb purgé."""
    vals = ws.get_all_values()
    if len(vals) < 2:
        return 0
    limite = (datetime.now() - timedelta(days=jours)).strftime("%Y-%m-%d")
    garde = []
    for r in vals[1:]:
        d = str(r[0]).strip()[:10] if r else ""
        if not d or not d[:4].isdigit() or d >= limite:
            garde.append(r)
    n_sup = (len(vals) - 1) - len(garde)
    if n_sup > 0:
        lignes = [vals[0]] + garde
        ws.update(lignes, value_input_option="USER_ENTERED")
        ws.resize(rows=max(len(lignes), 2))
    return n_sup

# ======================== CONFIGURATION ========================
CHEMIN_CRED = r"C:\Users\jfilt\bnc_secrets\compte_service.json"
from chemins_bnc import dossier_actions   # Windows FR (« Mon Drive ») ou EN (« My Drive »)
CHEMIN_LOG = os.path.join(dossier_actions(), "bnc_sync_log.txt")
NOM_DEST = "Action_2026-c_New"
ONGLET_SOURCE = "LesAffaires"    # onglet source, dans le MÊME classeur

# Défauts de repli si l'en-tête correspondante est absente (sinon : repérage par nom).
SRC_COL_DATE = 0      # A  (« Date »)
SRC_COL_SYMBOLE = 1   # B  (« Symbole ») — était en C avant l'ajout de « Description »
SRC_COL_CIBLE = 3     # D  (« Cours cible »)
FEUILLES_DEST = ["Prospects"]
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


def cle_symbole(s):
    """Clé de correspondance UNIFIÉE : suffixe canadien retiré + point de classe
    d'action normalisé en tiret. Ainsi « BBD.B » (notation Les Affaires),
    « BBD-B » et « BBD-B.TO » (notation Yahoo) donnent tous « BBD-B », et la
    map garde la plus récente des trois."""
    return base_symbole(str(s).strip().upper()).replace('.', '-')


def trouver(entetes, nom):
    cible = ' '.join(nom.split())
    for i, h in enumerate(entetes):
        if ' '.join(str(h).split()) == cible:
            return i
    return None


# --- Complétion des lignes ajoutées à la main dans Prospects (A+B seulement) ---
# Bourses Google Finance par suffixe Yahoo ; les classes d'action reprennent le
# point côté GF (BBD-B.TO -> TSE:BBD.B). US : « SYM:NASDAQ » — la formule K du
# Sheet retire « :NASDAQ » et GOOGLEFINANCE résout la bourse toute seule.
SUFFIXE_GF = {'.TO': 'TSE', '.V': 'CVE', '.NE': 'NEO', '.CN': 'CNSX'}


def symbole_gf(sym):
    s = str(sym).strip().upper()
    for suf, prefixe in SUFFIXE_GF.items():
        if s.endswith(suf):
            return f"{prefixe}:{s[:-len(suf)].replace('-', '.')}"
    return f"{s.replace('-', '.')}:NASDAQ"


def completer_lignes_prospects(ws, vals=None):
    """Complète les colonnes dérivées des lignes ajoutées à la MAIN dans Prospects
    (l'utilisateur ne remplit que Symbole + Description) : C Lien Yahoo,
    J Lien Google Finance, L Prix GF et M Devise (formules GOOGLEFINANCE).
    K (Symbole GF) se remplit seul via l'ARRAYFORMULA du Sheet ; F-I (Prix $,
    Pré G %, Pré YF, MAJ YF) sont écrites par l'app au prochain passage Yahoo.

    Une ligne n'est « nouvelle » que si Lien Yahoo ET Lien Google Finance sont
    tous deux vides : on ne touche jamais aux lignes existantes (ex. les CDR
    dont le Prix GF est vide EXPRÈS, Google Finance ne les couvrant pas).
    Renvoie le nombre de lignes complétées."""
    if vals is None:
        vals = ws.get_all_values()
    if not vals:
        return 0
    ent = vals[0]
    i_sym = trouver(ent, 'Symbole')
    i_desc = trouver(ent, 'Description')
    i_ly = trouver(ent, 'Lien Yahoo')
    i_lgf = trouver(ent, 'Lien Google Finance')
    i_sgf = trouver(ent, 'Symbole GF')
    i_pgf = trouver(ent, 'Prix GF')
    i_dev = trouver(ent, 'Devise')
    if i_sym is None or i_ly is None or i_lgf is None:
        return 0

    def cellule_vide(row, i):
        return i is not None and (len(row) <= i or not str(row[i]).strip())

    updates = []
    n_completees = 0
    for r, row in enumerate(vals[1:], start=2):
        sym = str(row[i_sym]).strip().upper() if len(row) > i_sym else ''
        if not sym or sym == '0':
            continue
        # Nouvelle ligne = les DEUX liens sont vides (sinon on ne touche à rien).
        if not (cellule_vide(row, i_ly) and cellule_vide(row, i_lgf)):
            continue

        desc = (str(row[i_desc]).strip() if (i_desc is not None and len(row) > i_desc) else '') or sym
        texte = f"{desc} ({sym}) Stock Price, News, Quote & History - Yahoo Finance".replace('"', '""')
        updates.append({'range': gspread.utils.rowcol_to_a1(r, i_ly + 1),
                        'values': [[f'=HYPERLINK("https://ca.finance.yahoo.com/quote/{sym}";"{texte}")']]})
        updates.append({'range': gspread.utils.rowcol_to_a1(r, i_lgf + 1),
                        'values': [[f"https://www.google.com/finance/beta/quote/{symbole_gf(sym)}"]]})
        if i_sgf is not None:
            cell_k = gspread.utils.rowcol_to_a1(r, i_sgf + 1)
            if cellule_vide(row, i_pgf):
                updates.append({'range': gspread.utils.rowcol_to_a1(r, i_pgf + 1),
                                'values': [[f"=GOOGLEFINANCE({cell_k})"]]})
            if cellule_vide(row, i_dev):
                updates.append({'range': gspread.utils.rowcol_to_a1(r, i_dev + 1),
                                'values': [[f'=GOOGLEFINANCE({cell_k};"currency")']]})
        n_completees += 1

    if updates:
        ws.batch_update(updates, value_input_option='USER_ENTERED')
    return n_completees


def date_key(s):
    """Clé de tri d'une date (chaîne) : datetime, ou datetime.min si illisible."""
    s = str(s).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return datetime.min


def main():
    if not os.path.exists(CHEMIN_CRED):
        journal(f"ERREUR : JSON compte de service introuvable : {CHEMIN_CRED}")
        return

    gc = gspread.service_account(filename=CHEMIN_CRED)
    dest_classeur = gc.open(NOM_DEST)

    # 1) Lire la source : map SYMBOLE -> (date, cible), onglet du même classeur
    src = dest_classeur.worksheet(ONGLET_SOURCE)
    n_purge = purger_source(src)
    if n_purge:
        journal(f"{n_purge} entrée(s) de plus de {JOURS_RETENTION_SOURCE} j purgée(s) de « {ONGLET_SOURCE} ».")
    lignes = src.get_all_values()
    # Colonnes repérées par NOM d'en-tête (robuste au déplacement des colonnes de
    # « LesAffaires ») ; repli sur les défauts A/B/D si une en-tête est absente.
    ent_src = lignes[0] if lignes else []
    i_date = trouver(ent_src, 'Date')
    i_sym = trouver(ent_src, 'Symbole')
    i_cible = trouver(ent_src, 'Cours cible')
    if i_date is None:  i_date = SRC_COL_DATE
    if i_sym is None:   i_sym = SRC_COL_SYMBOLE
    if i_cible is None: i_cible = SRC_COL_CIBLE
    affaires = {}
    for row in lignes[1:]:
        if len(row) <= max(i_sym, i_cible, i_date):
            continue
        sym = str(row[i_sym]).strip().upper()
        if not sym:
            continue
        cible = parse_nombre(row[i_cible])
        date = str(row[i_date]).strip()
        if cible is None:
            continue
        # Cle UNIFIEE (exact/base/notation de classe confondus) : un meme titre peut
        # apparaitre sous plusieurs notations et plusieurs dates -> on garde la PLUS RECENTE.
        cle = cle_symbole(sym)
        ancien = affaires.get(cle)
        if ancien is None or date_key(date) >= date_key(ancien[0]):
            affaires[cle] = (date, cible)
    journal(f"{len(affaires)} objectifs lus depuis l'onglet « {ONGLET_SOURCE} ».")

    # 2) Écrire dans chaque onglet de la destination
    dest = dest_classeur
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
            # Correspondance par cle unifiee : exact, base et notations de classe
            # (BBD.B / BBD-B.TO) se rejoignent, la map contient deja la plus recente.
            entree = affaires.get(cle_symbole(sym))
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

        # Lignes ajoutees a la main (A+B seulement) : completer C/J/L/M.
        n_compl = completer_lignes_prospects(ws, vals)
        if n_compl:
            journal(f"  [OK] {nom_feuille} : {n_compl} nouvelle(s) ligne(s) completee(s) (liens + GOOGLEFINANCE).")

    journal("Mise a jour Les Affaires terminee avec succes.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        journal(f"ECHEC : {type(e).__name__} - {e}")
        journal(traceback.format_exc())
