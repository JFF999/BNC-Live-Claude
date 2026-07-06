"""
Étend le mécanisme VLOOKUP aux colonnes CALCULÉES par l'app (comme Aff_Data l'a fait
pour Pré Aff / MAJ Aff). But : que ces colonnes suivent aussi le symbole de leur ligne
au recalcul Excel, au lieu d'être des valeurs fixes collées à une position.

Deux feuilles cachées, une par onglet visible (colonnes différentes) :
  - Data_Pros (sheet4) : Symbole | Prix $ | Pré G % | Pré YF | MAJ YF        -> Prospects F..I
  - Data_Port (sheet5) : Symbole | Prix $ | Pré G % | Gain $ | Gain % | Var % | Pré YF | MAJ YF
                                                                              -> Portefeuille K..Q
MAJ YF (chaîne "YYYY-MM-DD HH:MM" dans le Sheet) est sérialisée pour l'affichage date.

setup (une fois)  :  python vlookup_calc.py [chemin.xlsx]
   ajoute les 2 feuilles cachées + convertit F..I / K..Q en formules VLOOKUP (chirurgie
   XML, préserve le cache des formules ; retire calcChain). Idempotent.
La synchro inverse importe `refresh_data_calc()` pour reconstruire ces feuilles à chaque run.
"""

import os
import re
import sys
import shutil
import zipfile
from datetime import datetime

import gspread

CHEMIN_XLSX_DEFAUT = r"C:\Users\jfilt\My Drive\Actions\Action_2026-c_New.xlsx"
CHEMIN_CRED = r"C:\Users\jfilt\bnc_secrets\compte_service.json"
NOM_GOOGLE_SHEET = "Action_2026-c_New"
EPOCH_EXCEL = datetime(1899, 12, 30)

# (col_visible, entête Sheet, est_date)
SPECS = [
    {
        "visible_xml": "xl/worksheets/sheet1.xml", "symcol": "A",
        "data_name": "Data_Pros", "data_xml": "xl/worksheets/sheet4.xml",
        "rid": "rId9", "sheetid": 4, "source_tab": "Prospects",
        "cols": [("F", "Prix $", False), ("G", "Pré G %", False),
                 ("H", "Pré YF", False), ("I", "MAJ YF", True)],
    },
    {
        "visible_xml": "xl/worksheets/sheet2.xml", "symcol": "C",
        "data_name": "Data_Port", "data_xml": "xl/worksheets/sheet5.xml",
        "rid": "rId10", "sheetid": 5, "source_tab": "Portefeuille BNC",
        "cols": [("K", "Prix $", False), ("L", "Pré G %", False), ("M", "Gain $", False),
                 ("N", "Gain %", False), ("O", "Var %", False), ("P", "Pré YF", False),
                 ("Q", "MAJ YF", True)],
    },
]


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def col_lettre(idx1):  # 1->A, 2->B, ...
    s = ""
    while idx1 > 0:
        idx1, r = divmod(idx1 - 1, 26)
        s = chr(65 + r) + s
    return s


def vers_serie(s):
    s = str(s).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return (datetime.strptime(s, fmt) - EPOCH_EXCEL).total_seconds() / 86400.0
        except ValueError:
            pass
    return None


# ------------------- lecture des valeurs calculées depuis le Sheet -------------------
def _lire_tab(classeur, spec):
    """Renvoie liste [(symbole, [valeurs par col]), ...] pour un onglet."""
    ws = classeur.worksheet(spec["source_tab"])
    vals = ws.get_values(value_render_option=gspread.utils.ValueRenderOption.unformatted)
    if not vals:
        return []
    ent = [" ".join(str(h).split()) for h in vals[0]]
    try:
        i_sym = ent.index("Symbole")
    except ValueError:
        return []
    idx = []
    for _, header, est_date in spec["cols"]:
        h = " ".join(header.split())
        idx.append((ent.index(h) if h in ent else None, est_date))
    out, vus = [], set()
    for ligne in vals[1:]:
        if i_sym >= len(ligne):
            continue
        sym = str(ligne[i_sym]).strip()
        if not sym or sym == "0" or sym in vus:
            continue
        vus.add(sym)
        valeurs = []
        for i_col, est_date in idx:
            v = ligne[i_col] if (i_col is not None and i_col < len(ligne)) else ""
            if est_date and v not in (None, ""):
                s = vers_serie(v)
                v = s if s is not None else ""
            valeurs.append(v)
        out.append((sym, valeurs))
    return out


def _cellule_valeur(ref, v):
    if v is None or str(v).strip() == "":
        return f'<c r="{ref}" t="inlineStr"><is><t></t></is></c>'
    try:
        return f'<c r="{ref}"><v>{float(v)}</v></c>'
    except (ValueError, TypeError):
        return f'<c r="{ref}" t="inlineStr"><is><t>{esc(v)}</t></is></c>'


def _construire_data_xml(spec, rows):
    ncol = 1 + len(spec["cols"])
    entete = [f'<c r="A1" t="inlineStr"><is><t>Symbole</t></is></c>']
    for j, (_, header, _) in enumerate(spec["cols"], start=2):
        entete.append(f'<c r="{col_lettre(j)}1" t="inlineStr"><is><t>{esc(header)}</t></is></c>')
    lignes = ['<row r="1">' + "".join(entete) + "</row>"]
    r = 2
    for sym, valeurs in rows:
        cells = [f'<c r="A{r}" t="inlineStr"><is><t>{esc(sym)}</t></is></c>']
        for j, v in enumerate(valeurs, start=2):
            cells.append(_cellule_valeur(f"{col_lettre(j)}{r}", v))
        lignes.append(f'<row r="{r}">' + "".join(cells) + "</row>")
        r += 1
    n = max(r - 1, 1)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:{col_lettre(ncol)}{n}"/><sheetData>' + "".join(lignes) +
            "</sheetData></worksheet>")


def refresh_data_calc():
    """{data_xml: contenu_xml} pour reconstruire Data_Pros / Data_Port depuis le Sheet."""
    gc = gspread.service_account(filename=CHEMIN_CRED)
    classeur = gc.open(NOM_GOOGLE_SHEET)
    out = {}
    for spec in SPECS:
        rows = _lire_tab(classeur, spec)
        out[spec["data_xml"]] = _construire_data_xml(spec, rows)
    return out


# ------------------- conversion des cellules visibles en VLOOKUP -------------------
def _convertir_feuille(xml, spec):
    symcol = spec["symcol"]
    dernier = col_lettre(1 + len(spec["cols"]))
    n = [0]

    def poser(bloc, ref, formule):
        pat = re.compile(r'<c r="' + re.escape(ref) + r'"([^>]*?)(?:/>|>.*?</c>)', re.S)
        m = pat.search(bloc)
        if not m:
            return bloc
        sm = re.search(r'\bs="(\d+)"', m.group(1))
        s_attr = f' s="{sm.group(1)}"' if sm else ''
        cell = f'<c r="{ref}"{s_attr}><f>{esc(formule)}</f></c>'
        return bloc[:m.start()] + cell + bloc[m.end():]

    def traiter(m):
        rn = m.group(1)
        bloc = m.group(0)
        sm = re.search(r'<c r="' + symcol + rn + r'"[^>]*?(?:/>|>.*?</c>)', bloc, re.S)
        if not sm or "<f" not in sm.group(0):
            return bloc
        touche = False
        for offset, (colv, _, _) in enumerate(spec["cols"], start=2):
            f = f'IFERROR(VLOOKUP(${symcol}{rn},{spec["data_name"]}!$A:${dernier},{offset},0),"")'
            nb = poser(bloc, f"{colv}{rn}", f)
            if nb != bloc:
                bloc, touche = nb, True
        if touche:
            n[0] += 1
        return bloc

    xml = re.sub(r'<row r="(\d+)"[^>]*>.*?</row>', traiter, xml, flags=re.S)
    return xml, n[0]


def _patcher_workbook(wb):
    for spec in SPECS:
        if f'name="{spec["data_name"]}"' not in wb:
            ajout = f'<sheet name="{spec["data_name"]}" sheetId="{spec["sheetid"]}" state="hidden" r:id="{spec["rid"]}"/>'
            wb = re.sub(r'(</sheets>)', ajout + r'\1', wb, count=1)
    return wb


def _patcher_rels(rels):
    rels = re.sub(r'\s*<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '', rels)
    for spec in SPECS:
        cible = spec["data_xml"].replace("xl/", "")
        if f'Target="{cible}"' not in rels:
            ajout = (f'<Relationship Id="{spec["rid"]}" '
                     'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                     f'Target="{cible}"/>')
            rels = re.sub(r'(</Relationships>)', ajout + r'\1', rels, count=1)
    return rels


def _patcher_ct(ct):
    ct = re.sub(r'\s*<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', '', ct)
    for spec in SPECS:
        part = "/" + spec["data_xml"]
        if part not in ct:
            ajout = (f'<Override PartName="{part}" '
                     'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            ct = re.sub(r'(</Types>)', ajout + r'\1', ct, count=1)
    return ct


def setup(chemin):
    z = zipfile.ZipFile(chemin)
    noms = z.namelist()
    contenu = {n: z.read(n) for n in noms}
    z.close()

    contenu["xl/workbook.xml"] = _patcher_workbook(
        contenu["xl/workbook.xml"].decode("utf-8", "replace")).encode("utf-8")
    contenu["xl/_rels/workbook.xml.rels"] = _patcher_rels(
        contenu["xl/_rels/workbook.xml.rels"].decode("utf-8", "replace")).encode("utf-8")
    contenu["[Content_Types].xml"] = _patcher_ct(
        contenu["[Content_Types].xml"].decode("utf-8", "replace")).encode("utf-8")

    data = refresh_data_calc()
    for xmlname, xmlstr in data.items():
        contenu[xmlname] = xmlstr.encode("utf-8")

    total = 0
    for spec in SPECS:
        xml = contenu[spec["visible_xml"]].decode("utf-8", "replace")
        xml, nb = _convertir_feuille(xml, spec)
        contenu[spec["visible_xml"]] = xml.encode("utf-8")
        total += nb
        print(f"  {spec['source_tab']} : {nb} lignes converties ({spec['data_name']}).")

    ajout_xml = [s["data_xml"] for s in SPECS if s["data_xml"] not in noms]
    tmp = chemin + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for n in noms:
            if n == "xl/calcChain.xml":
                continue
            dst.writestr(n, contenu[n])
        for n in ajout_xml:
            dst.writestr(n, contenu[n])
    os.replace(tmp, chemin)
    print(f"Termine. {total} lignes converties ; feuilles {[s['data_name'] for s in SPECS]}.")


def main():
    chemin = sys.argv[1] if len(sys.argv) > 1 else CHEMIN_XLSX_DEFAUT
    if not os.path.exists(chemin):
        print(f"ERREUR : introuvable {chemin}")
        return
    if chemin == CHEMIN_XLSX_DEFAUT:
        bak = chemin + ".bak_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(chemin, bak)
        print(f"Sauvegarde : {bak}")
    setup(chemin)


if __name__ == "__main__":
    main()
