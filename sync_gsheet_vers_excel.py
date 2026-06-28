"""
Synchronisation INVERSE : Google Sheet « Action_2026-c_New » -> Excel Action_2026-c_New.xlsx.
Copie les colonnes CALCULÉES (écrites par l'app) du Sheet vers l'Excel :
  - Portefeuille BNC : K–Q  (Prix $, Pré G %, Gain $, Gain %, Var %, Pré YF, MAJ YF)
  - Prospects        : F–I  (Prix $, Pré G %, Pré YF, MAJ YF)

MÉTHODE : édition CHIRURGICALE du XML du .xlsx. On ne réécrit que les cellules visées ;
toutes les autres (dont les formules A–J et leur valeur en cache) restent intactes.
=> Contrairement à openpyxl, le cache des formules N'EST PAS effacé, donc la synchro
   ALLER (Excel -> Sheet) du matin continue de fonctionner.

À LANCER EN LOCAL avec le Python pythoncore. Pré-requis : voir sync_excel_vers_gsheet.py.
"""

import os
import re
import zipfile
import traceback
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread

# ======================== CONFIGURATION ========================
CHEMIN_XLSX = r"G:\My Drive\Actions\Action_2026-c_New.xlsx"
CHEMIN_CRED = r"C:\Users\jfilt\bnc_secrets\compte_service.json"
CHEMIN_LOG = r"G:\My Drive\Actions\bnc_sync_log.txt"
NOM_GOOGLE_SHEET = "Action_2026-c_New"

# Feuille -> (première lettre, dernière lettre) de la zone calculée à recopier.
ZONES = {
    "Portefeuille BNC": ("K", "Q"),
    "Prospects": ("F", "I"),
}
COL_DATE = "MAJ YF"   # en-tête de la colonne horodatage -> écrite comme date Excel (série)
EPOCH_EXCEL = datetime(1899, 12, 30)
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


def lettre_vers_index(lettre):
    """A->0, B->1, ... K->10."""
    idx = 0
    for c in lettre:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx - 1


def index_vers_lettre(idx):
    """0->A, 1->B, ..."""
    idx += 1
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def lettres_zone(debut, fin):
    return [chr(c) for c in range(ord(debut), ord(fin) + 1)]


def mapper_feuilles_xml(chemin):
    """Renvoie {nom_feuille: 'xl/worksheets/sheetN.xml'} via workbook.xml + rels."""
    z = zipfile.ZipFile(chemin)
    wb = z.read('xl/workbook.xml').decode('utf-8', 'replace')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'replace')
    z.close()
    rid_vers_cible = {m.group(1): m.group(2)
                      for m in re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    mapping = {}
    for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        cible = rid_vers_cible.get(m.group(2), '')
        if cible:
            mapping[m.group(1)] = 'xl/' + cible.lstrip('/')
    return mapping


def contenu_cellule(valeur, est_date):
    """Construit le corps <v>...</v> d'une cellule, ou None si vide."""
    if valeur is None or valeur == "":
        return None
    if est_date:
        s = str(valeur).strip()
        dt = None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(s, fmt)
                break
            except ValueError:
                pass
        if dt is None:
            return None
        serie = (dt - EPOCH_EXCEL).total_seconds() / 86400.0
        return f"<v>{serie:.8f}</v>"
    try:
        return f"<v>{float(valeur)}</v>"
    except (ValueError, TypeError):
        return None


def remplacer_cellule(bloc, ref, corps):
    """Remplace la cellule <c r="ref" ...> dans le bloc XML d'une ligne, en gardant son style s."""
    pat = re.compile(r'<c r="' + re.escape(ref) + r'"([^>]*?)(?:/>|>.*?</c>)', re.S)
    m = pat.search(bloc)
    if not m:
        return bloc  # cellule absente : on ne crée rien (la grille existe déjà normalement)
    sm = re.search(r'\bs="(\d+)"', m.group(1))
    s_attr = f' s="{sm.group(1)}"' if sm else ''
    cellule = f'<c r="{ref}"{s_attr}/>' if corps is None else f'<c r="{ref}"{s_attr}>{corps}</c>'
    return bloc[:m.start()] + cellule + bloc[m.end():]


def symbole_de_bloc(bloc, col_symbole, rownum):
    """Extrait le symbole (valeur en cache) de la cellule col_symbole de la ligne."""
    cm = re.search(r'<c r="' + col_symbole + rownum + r'"[^>]*?(?:/>|>.*?</c>)', bloc, re.S)
    if not cm:
        return ""
    vm = re.search(r'<v>(.*?)</v>', cm.group(0), re.S)
    return vm.group(1).strip() if vm else ""


def modifier_xml_feuille(xml, colonnes, col_symbole, map_symbole):
    """Pour chaque ligne, lit le SYMBOLE (cache de la cellule col_symbole) et écrit les
    valeurs du Sheet correspondant à CE symbole (correspondance par symbole, pas position).
    colonnes: liste (lettre, est_date). map_symbole: {symbole: {lettre: valeur}}."""
    def traiter_row(m):
        rownum = m.group(1)
        bloc = m.group(0)
        symbole = symbole_de_bloc(bloc, col_symbole, rownum)
        vals = map_symbole.get(symbole)
        if not symbole or vals is None:
            return bloc
        for lettre, est_date in colonnes:
            bloc = remplacer_cellule(bloc, f"{lettre}{rownum}", contenu_cellule(vals.get(lettre), est_date))
        return bloc
    return re.sub(r'<row r="(\d+)"[^>]*>.*?</row>', traiter_row, xml, flags=re.S)


def reecrire_xlsx(chemin, modifs):
    """Réécrit le .xlsx en remplaçant uniquement les XML modifiés (le reste à l'identique)."""
    tmp = chemin + ".tmp"
    with zipfile.ZipFile(chemin, "r") as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = modifs.get(item.filename) or src.read(item.filename)
            if isinstance(data, str):
                data = data.encode("utf-8")
            dst.writestr(item, data)
    os.replace(tmp, chemin)


def main():
    if not os.path.exists(CHEMIN_XLSX):
        journal(f"ERREUR : Excel introuvable : {CHEMIN_XLSX}")
        return
    if not os.path.exists(CHEMIN_CRED):
        journal(f"ERREUR : JSON compte de service introuvable : {CHEMIN_CRED}")
        return

    gc = gspread.service_account(filename=CHEMIN_CRED)
    classeur = gc.open(NOM_GOOGLE_SHEET)
    feuilles_xml = mapper_feuilles_xml(CHEMIN_XLSX)

    modifs = {}
    for nom_feuille, (debut, fin) in ZONES.items():
        if nom_feuille not in feuilles_xml:
            journal(f"  [ATTENTION] Feuille XML introuvable pour '{nom_feuille}' - ignoree.")
            continue

        # Valeurs BRUTES du Sheet (nombres = nombres, texte = texte)
        ws = classeur.worksheet(nom_feuille)
        donnees = ws.get_values(value_render_option=gspread.utils.ValueRenderOption.unformatted)
        if not donnees:
            journal(f"  [ATTENTION] '{nom_feuille}' : Sheet vide - ignore.")
            continue
        entetes = donnees[0]
        if "Symbole" not in entetes:
            journal(f"  [ATTENTION] '{nom_feuille}' : colonne Symbole absente - ignore.")
            continue
        idx_sym = entetes.index("Symbole")
        lettre_sym = index_vers_lettre(idx_sym)

        lettres = lettres_zone(debut, fin)
        colonnes = []
        for L in lettres:
            idx = lettre_vers_index(L)
            est_date = (idx < len(entetes) and str(entetes[idx]).strip() == COL_DATE)
            colonnes.append((L, est_date))

        # symbole -> {lettre: valeur Sheet}  (correspondance par SYMBOLE, pas par position)
        map_symbole = {}
        for r in range(2, len(donnees) + 1):
            ligne = donnees[r - 1]
            sym = str(ligne[idx_sym]).strip() if idx_sym < len(ligne) else ""
            if not sym or sym == "0":
                continue
            map_symbole[sym] = {
                L: (ligne[lettre_vers_index(L)] if lettre_vers_index(L) < len(ligne) else "")
                for L in lettres
            }

        chemin_xml = feuilles_xml[nom_feuille]
        z = zipfile.ZipFile(CHEMIN_XLSX)
        xml = z.read(chemin_xml).decode("utf-8", "replace")
        z.close()
        modifs[chemin_xml] = modifier_xml_feuille(xml, colonnes, lettre_sym, map_symbole)
        journal(f"  [OK] {nom_feuille} : {len(map_symbole)} symboles correspondus ({debut}-{fin}).")

    if modifs:
        reecrire_xlsx(CHEMIN_XLSX, modifs)
        journal("Synchronisation inverse terminee avec succes (cache des formules preserve).")
    else:
        journal("Rien a ecrire.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        journal(f"ECHEC : {type(e).__name__} - {e}")
        journal(traceback.format_exc())
