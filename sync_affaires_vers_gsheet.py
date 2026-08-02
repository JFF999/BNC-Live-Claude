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


def notation_connue_prospects(vals_prospects):
    """Map clé unifiée -> notation Yahoo telle qu'écrite dans Prospects
    (source de vérité : AAPL -> AAPL, TECK-B.TO -> TECK-B.TO, ODV.V -> ODV.V).
    Une clé portée par PLUSIEURS notations (titre US + son CDR canadien, ex.
    NVDA et NVDA.TO) vaut None : « connu mais ambigu », le symbole source ne
    sera PAS touché (l'appariement Pré Aff sert les deux de toute façon)."""
    notation = {}
    if not vals_prospects:
        return notation
    i_sym = trouver(vals_prospects[0], 'Symbole')
    if i_sym is None:
        return notation
    for row in vals_prospects[1:]:
        s = str(row[i_sym]).strip().upper() if len(row) > i_sym else ''
        if s and s != '0':
            cle = cle_symbole(s)
            if cle in notation and notation[cle] != s:
                notation[cle] = None          # ambigu : ne pas y toucher
            elif cle not in notation:
                notation[cle] = s
    return notation


def normaliser_symboles_source(ws, vals=None, notation_connue=None):
    """Normalise la colonne Symbole de l'onglet LesAffaires : les entrées collées
    depuis LesAffaires.com arrivent SANS suffixe de bourse (AC, TECK.B).
    1) Si Prospects connaît le titre (clé unifiée), on adopte SA notation —
       AAPL reste AAPL, TECK.B -> TECK-B.TO, ODV -> ODV.V. Indispensable : le
       marqueur « $US » des cibles est inconsistant, plein de titres US n'en ont pas.
    2) Sinon (symbole inconnu), règle du cours cible : cible non vide sans « US »
       -> notation Yahoo « .TO » (point de classe -> tiret : QBR.B -> QBR-B.TO).
    Écrit les corrections dans la feuille ET dans `vals` (pour la suite du run).
    Renvoie le nombre de symboles corrigés."""
    if vals is None:
        vals = ws.get_all_values()
    if not vals:
        return 0
    notation_connue = notation_connue or {}
    ent = vals[0]
    i_sym = trouver(ent, 'Symbole')
    i_cible = trouver(ent, 'Cours cible')
    if i_sym is None:  i_sym = SRC_COL_SYMBOLE
    if i_cible is None: i_cible = SRC_COL_CIBLE

    updates = []
    for r, row in enumerate(vals[1:], start=2):
        sym = str(row[i_sym]).strip().upper() if len(row) > i_sym else ''
        if not sym:
            continue
        cle = cle_symbole(sym)
        if cle in notation_connue:
            connu = notation_connue[cle]
            if connu is None:
                continue                      # ambigu (US + CDR) : ne pas toucher
            nouveau = connu                   # notation officielle de Prospects
        else:
            cible = str(row[i_cible]).strip() if len(row) > i_cible else ''
            if not cible or 'US' in cible.upper():
                continue                      # cible vide (devise inconnue) ou titre US
            if sym.endswith(SUFFIXES_CAD):
                continue                      # déjà en notation Yahoo canadienne
            nouveau = sym.replace('.', '-') + '.TO'
        if nouveau == sym == str(row[i_sym]).strip():
            continue                          # déjà conforme
        updates.append({'range': gspread.utils.rowcol_to_a1(r, i_sym + 1),
                        'values': [[nouveau]]})
        vals[r - 1][i_sym] = nouveau          # la suite du run lit le symbole corrigé
    if updates:
        ws.batch_update(updates, value_input_option='USER_ENTERED')
    return len(updates)


def ajouter_titres_surperformance(ws_prospects, vals_prospects, lignes_source, seuil=15.0):
    """Ajoute à Prospects une ligne (Symbole + Description seulement) pour chaque
    titre de l'onglet LesAffaires qui respecte LES TROIS critères :
      1) absent de Prospects (clé de symbole unifiée) ;
      2) Note = « Surperformance » (strict : « Surperformance sectorielle » exclu) ;
      3) Potentiel >= 15 % (texte FR « 34,70% » toléré).
    N'écrit que les colonnes A/B : les liens et formules GOOGLEFINANCE viennent de
    completer_lignes_prospects, Pré Aff/MAJ Aff de la synchro, F-I de l'app.
    Met aussi à jour `vals_prospects` en mémoire (les étapes suivantes du même run
    voient les nouvelles lignes). Renvoie la liste des symboles ajoutés."""
    if not vals_prospects or not lignes_source:
        return []
    ent_p = vals_prospects[0]
    i_sym_p = trouver(ent_p, 'Symbole')
    i_desc_p = trouver(ent_p, 'Description')
    if i_sym_p is None:
        return []
    cles_presentes = set()
    for row in vals_prospects[1:]:
        s = str(row[i_sym_p]).strip().upper() if len(row) > i_sym_p else ''
        if s and s != '0':
            cles_presentes.add(cle_symbole(s))

    ent_s = lignes_source[0]
    i_sym = trouver(ent_s, 'Symbole')
    i_desc = trouver(ent_s, 'Description')
    i_pot = trouver(ent_s, 'Potentiel')
    i_note = trouver(ent_s, 'Note')
    i_date = trouver(ent_s, 'Date')
    if i_sym is None or i_pot is None or i_note is None:
        return []

    a_ajouter = {}                     # cle -> (date, symbole, description) ; la plus RÉCENTE gagne
    for row in lignes_source[1:]:
        sym = str(row[i_sym]).strip().upper() if len(row) > i_sym else ''
        if not sym:
            continue
        note = str(row[i_note]).strip().lower() if len(row) > i_note else ''
        if note != 'surperformance':
            continue
        pot = parse_nombre(str(row[i_pot]).replace('%', '')) if len(row) > i_pot else None
        if pot is None or pot < seuil:
            continue
        cle = cle_symbole(sym)
        if cle in cles_presentes:
            continue
        desc = str(row[i_desc]).strip() if (i_desc is not None and len(row) > i_desc) else ''
        date = str(row[i_date]).strip() if (i_date is not None and len(row) > i_date) else ''
        ancien = a_ajouter.get(cle)
        if ancien is None or date_key(date) >= date_key(ancien[0]):
            a_ajouter[cle] = (date, sym, desc or sym)

    if not a_ajouter:
        return []
    # Écrire UNIQUEMENT A:B (ne jamais poser de '' sous l'ARRAYFORMULA de K).
    premiere = len(vals_prospects) + 1
    nouvelles_ab = [[sym, desc] for _, sym, desc in a_ajouter.values()]
    ws_prospects.update(nouvelles_ab,
                        f"A{premiere}:B{premiere + len(nouvelles_ab) - 1}",
                        value_input_option='USER_ENTERED')
    for _, sym, desc in a_ajouter.values():
        ligne = [''] * len(ent_p)
        ligne[i_sym_p] = sym
        if i_desc_p is not None:
            ligne[i_desc_p] = desc
        vals_prospects.append(ligne)
    return [sym for _, sym, _d in a_ajouter.values()]


ONGLET_TOP50 = "Aff_Top50Quebec"   # palmarès Les Affaires « Top 50 Québec »


def ajouter_titres_top50(ws_prospects, vals_prospects, lignes_top50, seuil=25.0):
    """Ajoute à Prospects une ligne (Symbole + Description seulement) pour chaque
    titre de l'onglet Aff_Top50Quebec (palmarès Les Affaires) qui respecte LES
    DEUX critères :
      1) absent de Prospects (clé de symbole unifiée) ;
      2) Pré G % >= 25 % (texte FR « 34,44% » toléré).
    Les symboles du palmarès sont déjà en notation Yahoo (ex. DRX.TO) ; la
    description vient de « Nom de l'entreprise ». Même contrat que
    ajouter_titres_surperformance : n'écrit que les colonnes A/B (liens et
    formules via completer_lignes_prospects, Pré Aff via la synchro, F-I via
    l'app), et met à jour `vals_prospects` en mémoire. Renvoie la liste des
    symboles ajoutés."""
    if not vals_prospects or not lignes_top50:
        return []
    ent_p = vals_prospects[0]
    i_sym_p = trouver(ent_p, 'Symbole')
    i_desc_p = trouver(ent_p, 'Description')
    if i_sym_p is None:
        return []
    cles_presentes = set()
    for row in vals_prospects[1:]:
        s = str(row[i_sym_p]).strip().upper() if len(row) > i_sym_p else ''
        if s and s != '0':
            cles_presentes.add(cle_symbole(s))

    ent_s = lignes_top50[0]
    i_sym = trouver(ent_s, 'Symbole')
    i_desc = trouver(ent_s, "Nom de l'entreprise")
    i_preg = trouver(ent_s, 'Pré G %')
    if i_sym is None or i_preg is None:
        return []

    a_ajouter = {}                     # cle -> (symbole, description)
    for row in lignes_top50[1:]:
        sym = str(row[i_sym]).strip().upper() if len(row) > i_sym else ''
        if not sym:
            continue
        preg = parse_nombre(str(row[i_preg]).replace('%', '')) if len(row) > i_preg else None
        if preg is None or preg < seuil:
            continue
        cle = cle_symbole(sym)
        if cle in cles_presentes or cle in a_ajouter:
            continue
        desc = str(row[i_desc]).strip() if (i_desc is not None and len(row) > i_desc) else ''
        a_ajouter[cle] = (sym, desc or sym)

    if not a_ajouter:
        return []
    # Écrire UNIQUEMENT A:B (ne jamais poser de '' sous l'ARRAYFORMULA de K).
    premiere = len(vals_prospects) + 1
    nouvelles_ab = [[sym, desc] for sym, desc in a_ajouter.values()]
    ws_prospects.update(nouvelles_ab,
                        f"A{premiere}:B{premiere + len(nouvelles_ab) - 1}",
                        value_input_option='USER_ENTERED')
    for sym, desc in a_ajouter.values():
        ligne = [''] * len(ent_p)
        ligne[i_sym_p] = sym
        if i_desc_p is not None:
            ligne[i_desc_p] = desc
        vals_prospects.append(ligne)
    return [sym for sym, _d in a_ajouter.values()]


# Jaune doux — même teinte que le surlignage « possédé » de l'app (or à 40 % sur blanc).
COULEUR_POSSEDE = {'red': 1.0, 'green': 0.937, 'blue': 0.6}


def surligner_possedes(spreadsheet, ws_prospects, vals_prospects, vals_portefeuille):
    """Fond JAUNE sur chaque ligne de Prospects dont le titre est aussi dans
    Portefeuille BNC (clé unifiée ; lignes No.=0 du portefeuille ignorées),
    fond par défaut sur les autres — un titre vendu perd donc son jaune.
    Les plages contiguës sont regroupées (peu de requêtes). Renvoie le
    nombre de lignes surlignées."""
    if not vals_prospects or len(vals_prospects) < 2:
        return 0
    ent_p = vals_prospects[0]
    i_sym = trouver(ent_p, 'Symbole')
    if i_sym is None:
        return 0

    cles_port = set()
    if vals_portefeuille:
        ent_b = vals_portefeuille[0]
        i_sym_b = trouver(ent_b, 'Symbole')
        i_no = trouver(ent_b, 'No.')
        if i_sym_b is not None:
            for row in vals_portefeuille[1:]:
                s = str(row[i_sym_b]).strip().upper() if len(row) > i_sym_b else ''
                if not s or s == '0':
                    continue
                if i_no is not None and len(row) > i_no and str(row[i_no]).strip() == '0':
                    continue                  # ligne inactive du portefeuille
                cles_port.add(cle_symbole(s))

    n_cols = len(ent_p)
    etats = []                                # True = à surligner
    for row in vals_prospects[1:]:
        s = str(row[i_sym]).strip().upper() if len(row) > i_sym else ''
        etats.append(bool(s) and s != '0' and cle_symbole(s) in cles_port)

    requests = []
    debut = 0
    for i in range(1, len(etats) + 1):
        if i == len(etats) or etats[i] != etats[debut]:
            if etats[debut]:
                cell = {'userEnteredFormat': {'backgroundColor': COULEUR_POSSEDE}}
            else:
                cell = {'userEnteredFormat': {}}   # remet le fond par défaut
            requests.append({'repeatCell': {
                'range': {'sheetId': ws_prospects.id,
                          'startRowIndex': debut + 1, 'endRowIndex': i + 1,
                          'startColumnIndex': 0, 'endColumnIndex': n_cols},
                'cell': cell,
                'fields': 'userEnteredFormat.backgroundColor'}})
            debut = i
    if requests:
        spreadsheet.batch_update({'requests': requests})
    return sum(etats)


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
    # Symboles collés sans suffixe de bourse : normaliser en notation Yahoo AVANT
    # de construire la map (notation Prospects prioritaire, sinon règle du $US).
    try:
        ws_pros = dest_classeur.worksheet('Prospects')
        vals_pros = ws_pros.get_all_values()
    except Exception:
        ws_pros, vals_pros = None, []
    n_norm = normaliser_symboles_source(src, lignes, notation_connue_prospects(vals_pros))
    if n_norm:
        journal(f"{n_norm} symbole(s) normalisé(s) dans « {ONGLET_SOURCE} ».")

    # Nouveaux titres « Surperformance » (Potentiel >= 15 %) absents de Prospects :
    # créer leur ligne (A+B) — la suite du run remplit Pré Aff, liens et formules.
    if ws_pros is not None:
        try:
            ajoutes = ajouter_titres_surperformance(ws_pros, vals_pros, lignes)
            if ajoutes:
                journal(f"{len(ajoutes)} titre(s) Surperformance ajouté(s) à Prospects : {', '.join(ajoutes)}")
        except Exception as e:
            journal(f"  [ATTENTION] ajout Surperformance : {type(e).__name__} - {e}")
        # Palmarès Top 50 Québec : titres à Pré G % >= 25 % absents de Prospects.
        try:
            src_top50 = dest_classeur.worksheet(ONGLET_TOP50).get_all_values()
            ajoutes50 = ajouter_titres_top50(ws_pros, vals_pros, src_top50)
            if ajoutes50:
                journal(f"{len(ajoutes50)} titre(s) Top 50 Québec ajouté(s) à Prospects : {', '.join(ajoutes50)}")
        except Exception as e:
            journal(f"  [ATTENTION] ajout Top 50 Québec : {type(e).__name__} - {e}")
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

        # Surligner en jaune les titres de Prospects presents dans Portefeuille BNC.
        if nom_feuille == 'Prospects':
            try:
                vals_port = dest.worksheet('Portefeuille BNC').get_all_values()
                n_jaune = surligner_possedes(dest, ws, vals, vals_port)
                journal(f"  [OK] {nom_feuille} : {n_jaune} ligne(s) surlignee(s) (titres du portefeuille).")
            except Exception as e:
                journal(f"  [ATTENTION] surlignage : {type(e).__name__} - {e}")

    journal("Mise a jour Les Affaires terminee avec succes.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        journal(f"ECHEC : {type(e).__name__} - {e}")
        journal(traceback.format_exc())
