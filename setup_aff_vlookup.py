"""
SETUP (une seule fois) : rend Pré Aff / MAJ Aff ROBUSTES au recalcul dans l'Excel
« Action_2026-c_New.xlsx ».

PROBLÈME résolu : les colonnes Symbole sont des formules à lien externe dont le
cache ne se rafraîchit qu'à l'ouverture d'Excel. La synchro inverse écrivait Pré Aff
comme VALEUR FIXE collée à la ligne (d'après le symbole EN CACHE). Au recalcul, les
symboles se décalent → Pré Aff ne suit pas → désalignement visible seulement dans Excel.

SOLUTION : une feuille cachée « Aff_Data » (Symbole | Pré Aff | MAJ Aff) alimentée par
Python, et les cellules Pré Aff/MAJ Aff deviennent des formules VLOOKUP qui suivent le
symbole de LEUR ligne. Toujours aligné, quel que soit le recalcul.

Ce script (chirurgie XML, préserve le cache des formules) :
  1. ajoute la feuille cachée Aff_Data (remplie depuis le Google Sheet) ;
  2. convertit Prospects D,E (clé A) et Portefeuille I,J (clé C) en formules VLOOKUP ;
  3. retire calcChain.xml (Excel le reconstruit) et nettoie rels + [Content_Types].

USAGE :  python setup_aff_vlookup.py [chemin.xlsx]
   (sans argument -> le vrai fichier ; fait une sauvegarde .bak_YYYYmmdd_HHMMSS).
Le fichier Excel doit être FERMÉ.
"""

import os
import re
import sys
import shutil
import zipfile
from datetime import datetime

import gspread

from chemins_bnc import dossier_actions   # Windows FR (« Mon Drive ») ou EN (« My Drive »)
CHEMIN_XLSX_DEFAUT = os.path.join(dossier_actions(), "Action_2026-c_New.xlsx")
CHEMIN_CRED = r"C:\Users\jfilt\bnc_secrets\compte_service.json"
NOM_GOOGLE_SHEET = "Action_2026-c_New"

# Feuille visible -> (fichier xml, colonne symbole, col Pré Aff, col MAJ Aff)
CIBLES = {
    "Prospects":        {"xml": "xl/worksheets/sheet1.xml", "sym": "A", "pa": "D", "maj": "E"},
    "Portefeuille BNC": {"xml": "xl/worksheets/sheet2.xml", "sym": "C", "pa": "I", "maj": "J"},
}
NOM_AFF = "Aff_Data"
XML_AFF = "xl/worksheets/sheet3.xml"


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ----------------------- lecture des données Affaires depuis le Sheet -----------------------
def lire_aff_data():
    """Renvoie une liste [(symbole, pre_aff, maj_aff_serie), ...] uniques, Pré Aff non vide."""
    gc = gspread.service_account(filename=CHEMIN_CRED)
    classeur = gc.open(NOM_GOOGLE_SHEET)
    table = {}
    for nom in ("Prospects", "Portefeuille BNC"):
        ws = classeur.worksheet(nom)
        vals = ws.get_values(value_render_option=gspread.utils.ValueRenderOption.unformatted)
        if not vals:
            continue
        ent = [" ".join(str(h).split()) for h in vals[0]]
        try:
            i_sym = ent.index("Symbole")
            i_pa = ent.index("Pré Aff")
        except ValueError:
            continue
        i_maj = ent.index("MAJ Aff") if "MAJ Aff" in ent else None
        for ligne in vals[1:]:
            if i_sym >= len(ligne):
                continue
            sym = str(ligne[i_sym]).strip()
            if not sym or sym == "0":
                continue
            pa = ligne[i_pa] if i_pa < len(ligne) else ""
            if pa is None or str(pa).strip() == "":
                continue
            maj = ligne[i_maj] if (i_maj is not None and i_maj < len(ligne)) else ""
            # 1re occurrence gagne (les deux onglets portent la même donnée Affaires)
            table.setdefault(sym, (pa, maj))
    return [(s, v[0], v[1]) for s, v in table.items()]


def construire_sheet_aff(rows):
    """XML complet de la feuille cachée Aff_Data (A=Symbole, B=Pré Aff, C=MAJ Aff)."""
    out = []
    out.append('<row r="1">'
               '<c r="A1" t="inlineStr"><is><t>Symbole</t></is></c>'
               '<c r="B1" t="inlineStr"><is><t>Pre Aff</t></is></c>'
               '<c r="C1" t="inlineStr"><is><t>MAJ Aff</t></is></c></row>')
    r = 2
    for sym, pa, maj in rows:
        cells = [f'<c r="A{r}" t="inlineStr"><is><t>{esc(sym)}</t></is></c>']
        # Pré Aff : numérique si possible
        try:
            cells.append(f'<c r="B{r}"><v>{float(pa)}</v></c>')
        except (ValueError, TypeError):
            cells.append(f'<c r="B{r}" t="inlineStr"><is><t>{esc(pa)}</t></is></c>')
        # MAJ Aff : sérialisée si numérique, sinon chaîne vide (-> "" via VLOOKUP)
        try:
            cells.append(f'<c r="C{r}"><v>{float(maj)}</v></c>')
        except (ValueError, TypeError):
            cells.append(f'<c r="C{r}" t="inlineStr"><is><t></t></is></c>')
        out.append(f'<row r="{r}">' + "".join(cells) + "</row>")
        r += 1
    n = max(r - 1, 1)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:C{n}"/><sheetData>' + "".join(out) + "</sheetData></worksheet>")


# ----------------------- conversion des cellules en VLOOKUP -----------------------
def convertir_feuille(xml, colsym, colpa, colmaj):
    """Remplace, sur chaque ligne DONNÉE (cellule symbole = formule), Pré Aff/MAJ Aff
    par des formules VLOOKUP sur Aff_Data. Préserve le style. Renvoie (xml, n_converties)."""
    n = [0]

    def cell_a_formule(bloc, ref, formule):
        f_esc = esc(formule)
        pat = re.compile(r'<c r="' + re.escape(ref) + r'"([^>]*?)(?:/>|>.*?</c>)', re.S)
        m = pat.search(bloc)
        if m:
            sm = re.search(r'\bs="(\d+)"', m.group(1))
            s_attr = f' s="{sm.group(1)}"' if sm else ''
            cell = f'<c r="{ref}"{s_attr}><f>{f_esc}</f></c>'
            return bloc[:m.start()] + cell + bloc[m.end():], True
        return bloc, False

    def traiter(m):
        rn = m.group(1)
        bloc = m.group(0)
        # cellule symbole : doit être une FORMULE (ligne de données, pas l'entête)
        sm = re.search(r'<c r="' + colsym + rn + r'"[^>]*?(?:/>|>.*?</c>)', bloc, re.S)
        if not sm or "<f" not in sm.group(0):
            return bloc
        f_pa = f'IFERROR(VLOOKUP(${colsym}{rn},{NOM_AFF}!$A:$C,2,0),"")'
        f_maj = f'IFERROR(VLOOKUP(${colsym}{rn},{NOM_AFF}!$A:$C,3,0),"")'
        bloc, ok1 = cell_a_formule(bloc, f"{colpa}{rn}", f_pa)
        bloc, ok2 = cell_a_formule(bloc, f"{colmaj}{rn}", f_maj)
        if ok1 or ok2:
            n[0] += 1
        return bloc

    xml = re.sub(r'<row r="(\d+)"[^>]*>.*?</row>', traiter, xml, flags=re.S)
    return xml, n[0]


# ----------------------- édition du workbook / rels / content-types -----------------------
def patcher_workbook(wb):
    if 'name="Aff_Data"' in wb:
        return wb
    ajout = f'<sheet name="{NOM_AFF}" sheetId="3" state="hidden" r:id="rId8"/>'
    return re.sub(r'(</sheets>)', ajout + r'\1', wb, count=1)


def patcher_rels(rels):
    rels = re.sub(r'\s*<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', '', rels)  # retire calcChain
    if 'Target="worksheets/sheet3.xml"' not in rels:
        ajout = ('<Relationship Id="rId8" '
                 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                 'Target="worksheets/sheet3.xml"/>')
        rels = re.sub(r'(</Relationships>)', ajout + r'\1', rels, count=1)
    return rels


def patcher_content_types(ct):
    ct = re.sub(r'\s*<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', '', ct)  # retire calcChain
    if '/xl/worksheets/sheet3.xml' not in ct:
        ajout = ('<Override PartName="/xl/worksheets/sheet3.xml" '
                 'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        ct = re.sub(r'(</Types>)', ajout + r'\1', ct, count=1)
    return ct


def main():
    chemin = sys.argv[1] if len(sys.argv) > 1 else CHEMIN_XLSX_DEFAUT
    if not os.path.exists(chemin):
        print(f"ERREUR : introuvable {chemin}")
        return

    rows = lire_aff_data()
    print(f"Aff_Data : {len(rows)} symboles avec Pré Aff.")

    # sauvegarde (seulement pour le vrai fichier, pas les copies de test)
    if chemin == CHEMIN_XLSX_DEFAUT:
        bak = chemin + ".bak_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(chemin, bak)
        print(f"Sauvegarde : {bak}")

    z = zipfile.ZipFile(chemin)
    noms = z.namelist()
    contenu = {n: z.read(n) for n in noms}
    z.close()

    wb = contenu["xl/workbook.xml"].decode("utf-8", "replace")
    rels = contenu["xl/_rels/workbook.xml.rels"].decode("utf-8", "replace")
    ct = contenu["[Content_Types].xml"].decode("utf-8", "replace")

    contenu["xl/workbook.xml"] = patcher_workbook(wb).encode("utf-8")
    contenu["xl/_rels/workbook.xml.rels"] = patcher_rels(rels).encode("utf-8")
    contenu["[Content_Types].xml"] = patcher_content_types(ct).encode("utf-8")
    contenu[XML_AFF] = construire_sheet_aff(rows).encode("utf-8")

    total = 0
    for nom, cfg in CIBLES.items():
        xml = contenu[cfg["xml"]].decode("utf-8", "replace")
        xml, nconv = convertir_feuille(xml, cfg["sym"], cfg["pa"], cfg["maj"])
        contenu[cfg["xml"]] = xml.encode("utf-8")
        total += nconv
        print(f"  {nom} : {nconv} lignes converties en VLOOKUP.")

    contenu.pop("xl/calcChain.xml", None)  # Excel le reconstruit

    tmp = chemin + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        # on garde l'ordre original + la nouvelle feuille à la fin
        for n in noms:
            if n == "xl/calcChain.xml":
                continue
            dst.writestr(n, contenu[n])
        dst.writestr(XML_AFF, contenu[XML_AFF])
    os.replace(tmp, chemin)
    print(f"Termine. {total} lignes converties, feuille cachée {NOM_AFF} ajoutée.")


if __name__ == "__main__":
    main()
