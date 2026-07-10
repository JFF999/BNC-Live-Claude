import math
import time
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import yfinance as yf
import gspread
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed

st.set_page_config(page_title="Portefeuille BNC", layout="wide")

# === v6 : source de données = Google Sheet « Action_2026-c_New » ===
# Ce Sheet est alimenté côté source par l'Apps Script (Excel -> Sheet : Portefeuille
# A–H, Prospects A–C). L'app lit le Sheet, calcule les données Yahoo / scores, puis
# réécrit UNIQUEMENT dans la zone autorisée. Fonctionne en local ET sur Streamlit
# Cloud (identifiants via st.secrets["gcp_service_account"]).
NOM_GOOGLE_SHEET = "Action_2026-c_New"

# Indice (0-based) de la PREMIÈRE colonne en lecture/écriture, par feuille.
# Tout ce qui est AVANT reste en LECTURE SEULE (données source synchronisées).
#   Portefeuille BNC : A–H lecture (0–7), I–P écriture (>= 8)
#   Prospects        : A–C lecture (0–2), D–I écriture (>= 3)
SEUIL_ECRITURE = {
    "Portefeuille BNC": 8,
    "Prospects": 3,
}

# Catégories de signal (du plus fort au plus faible)
SIGNAUX = ["Priorité", "À surveiller", "À valider", "Risque élevé", "Secondaire", "Objectif atteint"]

# --- ASTUCE CSS : Optimisation totale de l'espace sur mobile ---
st.markdown("""
    <style>
        [data-testid="stHeader"], #MainMenu, footer { display: none !important; }
        .block-container { padding-top: 1.5rem !important; padding-bottom: 1rem !important; }
        [data-testid="stElementToolbar"] { display: none !important; }

        /* Ajustement largeur minimale pour le menu des boutons en haut */
        div[data-testid="stHorizontalBlock"]:has(div[data-testid="stPopover"]) {
            flex-direction: row !important; flex-wrap: nowrap !important; gap: 10px !important;
        }
        div[data-testid="stHorizontalBlock"]:has(div[data-testid="stPopover"]) > div {
            width: auto !important; min-width: 0 !important; flex: none !important;
        }

        /* Ajustement largeur minimale pour la ligne des statistiques (Gain Total, etc.) */
        div[data-testid="stHorizontalBlock"]:has(div.stats-block) {
            flex-direction: row !important; flex-wrap: nowrap !important; align-items: center !important; gap: 15px !important;
        }
        div[data-testid="stHorizontalBlock"]:has(div.stats-block) > div {
            width: auto !important; min-width: 0 !important; flex: none !important;
        }

        /* Ajustement largeur minimale pour le bandeau des marchés */
        div[data-testid="stHorizontalBlock"]:has(div.market-block) {
            flex-direction: row !important; flex-wrap: nowrap !important; align-items: center !important; gap: 15px !important;
        }
        div[data-testid="stHorizontalBlock"]:has(div.market-block) > div {
            width: auto !important; min-width: 0 !important; flex: none !important;
        }

        .alert-box {
            background-color: rgba(255, 215, 0, 0.1); border-left: 4px solid #FFD700;
            padding: 10px 15px; margin-bottom: 15px; border-radius: 4px; font-size: 14px;
        }
        .alert-item { margin: 2px 0px; }

        /* === v8 : cartes de statistiques (Gain total / Gain du jour / Valeur) === */
        .stats-block {
            background: rgba(255, 255, 255, .045);
            border: 1px solid rgba(255, 255, 255, .09);
            border-radius: 10px;
            padding: 6px 14px 8px 14px;
        }

        /* === v8 : cartes Top 5 (onglet Décision) === */
        .rangee-top { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 6px; }
        .carte-top {
            flex: 1 1 165px; max-width: 240px;
            background: rgba(255, 255, 255, .05);
            border: 1px solid rgba(255, 255, 255, .10);
            border-radius: 12px; padding: 10px 12px;
        }
        .carte-top.possede { border-color: rgba(255, 215, 0, .55); }
        .ct-sym  { font-weight: 700; font-size: 16px; margin-bottom: 2px; }
        .ct-rang { color: #00A65A; font-weight: 600; font-size: 13px; }
        .ct-preg { font-size: 14px; font-weight: 600; margin: 2px 0px; }
        .ct-sig  { font-size: 12px; }
        .ct-pq   { font-size: 11px; color: gray; margin-top: 4px; line-height: 1.35; }

        /* === v8 : sur MOBILE, police des tableaux réduite (zoom) pour que TOUTES les
           colonnes du Portefeuille tiennent à l'écran ; même taille pour les Prospects. */
        @media (max-width: 640px) {
            div[data-testid="stDataFrame"] { zoom: 0.78; }

            /* Bandeau des marchés : même police que les autres lignes d'entête,
               passage sur 2 lignes si trop long, avec un espacement vertical net
               (sinon la 2e ligne — TSX — touche la 1re). */
            div.market-block { font-size: 13px; line-height: 1.5; }
            div[data-testid="stHorizontalBlock"]:has(div.market-block) {
                flex-wrap: wrap !important; row-gap: 14px !important;
            }

            /* Cartes de stats (Gain total / Gain du jour / Valeur totale) un cran
               plus petites pour que les TROIS tiennent dans la largeur de l'écran.
               (!important : les tailles sont posées en style inline par le markdown) */
            div[data-testid="stHorizontalBlock"]:has(div.stats-block) { gap: 8px !important; }
            div.stats-block { padding: 4px 8px 5px 8px !important; }
            div.stats-block p:first-child { font-size: 11px !important; }
            div.stats-block p:nth-child(2) { font-size: 13px !important; }
        }

        /* === v7 : sur MOBILE, garder Score min / Risque max côte à côte (Streamlit
           empile les colonnes sur écran étroit). Les Signaux passent en pleine largeur. */
        @media (max-width: 640px) {
            div[data-testid="stHorizontalBlock"]:has(div[data-testid="stNumberInput"]) {
                flex-direction: row !important; flex-wrap: wrap !important; gap: 10px !important;
            }
            div[data-testid="stHorizontalBlock"]:has(div[data-testid="stNumberInput"])
                > div:has(div[data-testid="stNumberInput"]) {
                width: calc(50% - 10px) !important; flex: 0 0 auto !important; min-width: 110px !important;
            }
            div[data-testid="stHorizontalBlock"]:has(div[data-testid="stNumberInput"])
                > div:has(div[data-testid="stMultiSelect"]) {
                width: 100% !important; flex: 1 1 100% !important;
            }
        }
    </style>
""", unsafe_allow_html=True)

# === v7 : rechargement du PARENT depuis un composant (iframe SANDBOXÉE sans
# allow-top-navigation : location.replace/reload sur le parent est BLOQUÉ par le
# navigateur). Stratégie à 3 étages : navigation directe si permise, sinon clic sur
# un lien créé dans le document PARENT (même origine -> navigation attribuée au
# parent), sinon meta refresh injecté dans le <head> du parent.
JS_RECHARGER_PARENT = """
    function rechargerParent(urlCible) {
        const p = window.parent;
        try { p.location.replace(urlCible); return; }
        catch (e) { try { p.console.log('BNC: nav directe bloquee -', e.message); } catch (_) {} }
        try {
            const a = p.document.createElement('a');
            a.href = urlCible; a.style.display = 'none';
            p.document.body.appendChild(a);
            a.click();
            a.remove();
            return;
        } catch (e) { try { p.console.log('BNC: lien parent bloque -', e.message); } catch (_) {} }
        try {
            const m = p.document.createElement('meta');
            m.httpEquiv = 'refresh';
            m.content = '0;url=' + urlCible;
            p.document.head.appendChild(m);
        } catch (e) { try { p.console.log('BNC: meta refresh bloque -', e.message); } catch (_) {} }
    }
"""

def est_mobile():
    # === v7 : détection téléphone via le User-Agent du navigateur (st.context,
    # Streamlit >= 1.37). Les navigateurs mobiles annoncent "Mobi" / "Android" /
    # "iPhone". En cas de doute (vieille version, en-tête absent) -> ordinateur.
    # ?mobile=1 dans l'URL force le mode téléphone (utile pour tester depuis un PC).
    try:
        if st.query_params.get("mobile", "") == "1":
            return True
    except Exception:
        pass
    try:
        ua = str(st.context.headers.get("User-Agent", "")).lower()
    except Exception:
        return False
    return any(m in ua for m in ("mobi", "android", "iphone", "ipod"))

@st.cache_data(ttl=300)
def heure_mise_a_jour():
    return datetime.now(ZoneInfo("America/Toronto")).strftime("%H:%M")

@st.cache_data(ttl=300, show_spinner=False)
def signature_donnees(symboles):
    # Empreinte d'un rafraîchissement RÉEL des données : capturée au moment du
    # téléchargement (mise en cache 5 min, comme le fetch Yahoo). Sert à déclencher
    # la sauvegarde auto UNE fois par rafraîchissement (et non à chaque rerun Streamlit).
    return datetime.now(ZoneInfo("America/Toronto")).isoformat()

@st.cache_data(ttl=300, show_spinner=False)
def obtenir_taux_change():
    try:
        return yf.Ticker("USDCAD=X").history(period="1d")['Close'].iloc[-1]
    except Exception:
        return 1.35

# --- CONNEXION GOOGLE SHEETS ---
def connecter_google_sheets():
    info_cles = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(info_cles)
    return gc.open(NOM_GOOGLE_SHEET)

def _nettoyer_entetes(entetes):
    # Nettoie + déduplique les en-têtes par sécurité : si deux colonnes portent le même
    # libellé, les doublons reçoivent un suffixe .1, .2 … (évite un plantage pandas).
    propres, vus = [], {}
    for h in entetes:
        h = ' '.join(str(h).replace('\n', ' ').replace('\r', '').split())
        if h in vus:
            vus[h] += 1
            propres.append(f"{h}.{vus[h]}")
        else:
            vus[h] = 0
            propres.append(h)
    return propres

# --- LECTURE DU GOOGLE SHEET ---
@st.cache_data(ttl=300)
def charger_donnees_base(nom_feuille):
    sh = connecter_google_sheets()
    feuille = sh.worksheet(nom_feuille)
    valeurs = feuille.get_all_values()
    if not valeurs:
        return pd.DataFrame()
    entetes = _nettoyer_entetes(valeurs[0])
    df = pd.DataFrame(valeurs[1:], columns=entetes)

    # --- LE NETTOYEUR DE NOMBRES FLOTTANTS ---
    colonnes_flottantes = [
        'Prix $', 'Achat $', 'Pré YF', 'Pré Aff',
        'Var %', 'Gain %', 'Gain $', 'Pré G %'
    ]

    for col in colonnes_flottantes:
        if col in df.columns:
            # Enleve symbole monetaire, pourcent et espaces (incl. insecables =
            # separateur de milliers FR : "3 329,55" -> "3329"), puis virgule -> point.
            # Remplacements LITTERAUX (regex=False) : pandas 3 / Arrow refuse les
            # motifs regex avec echappements unicode.
            df[col] = (df[col].astype(str)
                       .str.replace('$', '', regex=False)
                       .str.replace('%', '', regex=False)
                       .str.replace(' ', '', regex=False)
                       .str.replace(' ', '', regex=False)
                       .str.replace(' ', '', regex=False)
                       .str.replace(',', '.', regex=False))
            df[col] = pd.to_numeric(df[col], errors='coerce').astype(float)

    # --- LE NETTOYEUR D'ENTIERS (Quantité) ---
    if 'Qtée' in df.columns:
        df['Qtée'] = (df['Qtée'].astype(str)
                      .str.replace(' ', '', regex=False)
                      .str.replace(' ', '', regex=False)
                      .str.replace(' ', '', regex=False)
                      .str.replace(',', '', regex=False))
        df['Qtée'] = pd.to_numeric(df['Qtée'], errors='coerce').astype('Int64') # Int64 accepte les cases vides

    # --- 'No.' en numérique (get_all_values renvoie du texte) ---
    # Indispensable pour le filtre des lignes actives (No. != 0) du Portefeuille :
    # les lignes vides/modèles (No. = "0" ou vide) deviennent 0 et seront écartées.
    if 'No.' in df.columns:
        df['No.'] = pd.to_numeric(
            df['No.'].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '', regex=False),
            errors='coerce'
        ).fillna(0).astype('Int64')

    return df

def sauvegarder_donnees_dans_sheets(df_live, nom_feuille):
    # Écrit les colonnes calculées dans le Google Sheet, UNIQUEMENT dans la zone
    # autorisée (>= SEUIL_ECRITURE). Les colonnes source (A–H / A–C, synchronisées
    # depuis l'Excel) restent intactes. Retourne (succès: bool, message: str).
    seuil = SEUIL_ECRITURE.get(nom_feuille, 9999)
    try:
        sh = connecter_google_sheets()
        feuille = sh.worksheet(nom_feuille)
        valeurs = feuille.get_all_values()
        if not valeurs:
            return False, "Feuille vide."
        entetes = _nettoyer_entetes(valeurs[0])
        if 'Symbole' not in entetes:
            return False, "Colonne 'Symbole' introuvable."
        col_symbole_idx = entetes.index('Symbole')

        # Colonnes calculées (nom feuille -> nom dans df), seulement si en zone d'écriture
        colonnes_cibles = {
            'Prix $': 'Prix $',
            'Pré G %': 'Pré G %',
            'Gain %': 'Gain %',
            'Var %': 'Var %',
            'Pré YF': 'Pré 1an $ Yahoo',
            'Gain $': 'Gain $',
        }
        colonnes_a_ecrire = {}
        for sheet_col, df_col in colonnes_cibles.items():
            if sheet_col in entetes:
                idx = entetes.index(sheet_col)
                if idx >= seuil:   # zone d'écriture seulement (lecture seule sinon)
                    colonnes_a_ecrire[idx + 1] = df_col   # gspread : indices 1-based

        # Horodatage Yahoo : on écrit UNIQUEMENT « MAJ YF » (côté Yahoo).
        # « MAJ Aff » (côté Les Affaires) n'est jamais touchée par l'app.
        idx_maj_yf = entetes.index('MAJ YF') if 'MAJ YF' in entetes else None
        if idx_maj_yf is not None and idx_maj_yf < seuil:
            idx_maj_yf = None  # hors zone d'écriture -> on n'y touche pas

        if not colonnes_a_ecrire and idx_maj_yf is None:
            return False, "Aucune colonne en zone d'écriture pour cette feuille."

        # Index mémoire des données live par symbole
        dict_live = {}
        for _, r in df_live.iterrows():
            sym = r.get('Symbole Brut')
            if pd.notna(sym) and str(sym).strip() != "":
                dict_live[str(sym).strip()] = r

        horodatage = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M")
        mises_a_jour = []
        for i, row_data in enumerate(valeurs):
            if i == 0:
                continue  # ligne d'en-tête
            if len(row_data) <= col_symbole_idx:
                continue
            sym = str(row_data[col_symbole_idx]).strip()
            if not sym or sym not in dict_live:
                continue
            row_live = dict_live[sym]
            ligne = i + 1   # gspread : ligne 1-based (en-tête = ligne 1)
            ecrit_ligne = False
            for col1, df_col in colonnes_a_ecrire.items():
                valeur = row_live.get(df_col)
                if pd.notna(valeur):
                    # Pourcentage -> stocké en décimal (0.XX), comme dans la feuille
                    if '%' in df_col:
                        valeur_finale = round(float(valeur) / 100.0, 4)
                    else:
                        valeur_finale = round(float(valeur), 2)
                    mises_a_jour.append({
                        'range': gspread.utils.rowcol_to_a1(ligne, col1),
                        'values': [[valeur_finale]],
                    })
                    ecrit_ligne = True
            if ecrit_ligne and idx_maj_yf is not None:
                mises_a_jour.append({
                    'range': gspread.utils.rowcol_to_a1(ligne, idx_maj_yf + 1),
                    'values': [[horodatage]],
                })

        if not mises_a_jour:
            return False, "Aucune donnée à écrire."
        feuille.batch_update(mises_a_jour)
        return True, f"{len(mises_a_jour)} cellules mises à jour dans « {nom_feuille} »."

    except Exception as e:
        return False, f"Erreur Google Sheets : {e}"

def preparer_export_csv(df):
    df_export = df.copy()
    if 'Symbole Brut' in df_export.columns:
        df_export['Symbole'] = df_export['Symbole Brut']
        df_export = df_export.drop(columns=['Symbole Brut'])
    if 'Tendance' in df_export.columns:
        df_export = df_export.drop(columns=['Tendance'])
    return df_export.to_csv(index=False, sep=';').encode('utf-8-sig')

# === v7 : JOURNAL DES SIGNAUX ====================================================
# Archive une fois par jour, dans l'onglet « Journal » du Sheet, les signaux du moment :
#   - Achat : prospects en « Priorité » (avec prix, rang, potentiel)
#   - Vente : titres du portefeuille en « Vendre »
# Sert à (1) montrer les CHANGEMENTS vs la dernière séance dans l'onglet Décision et
# (2) mesurer a posteriori la performance des signaux (les Priorité ont-ils monté ?).
# =================================================================================
ENTETE_JOURNAL = ["Date", "Type", "Symbole", "Signal", "Prix", "Rang", "Pré G %"]

def _num(v, defaut=None):
    n = pd.to_numeric(pd.Series([v]), errors='coerce').iloc[0]
    return defaut if pd.isna(n) else float(n)

def journaliser_signaux(df_port, df_pros, valeur_port=None, indices=None):
    """Ajoute les signaux du jour à l'onglet Journal (créé au besoin, 1 fois/jour).
    v7+ : archive aussi la VALEUR du portefeuille et les indices (Type=Valeur) pour la
    courbe « Portefeuille vs marché ». Renvoie (journal_complet, ecrit_aujourdhui)."""
    try:
        sh = connecter_google_sheets()
        try:
            ws = sh.worksheet("Journal")
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet("Journal", rows=4000, cols=len(ENTETE_JOURNAL))
            ws.append_row(ENTETE_JOURNAL)
        vals = ws.get_all_values()
        aujourdhui = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d")
        if any(r and r[0] == aujourdhui for r in vals[1:]):
            return vals, False   # déjà journalisé aujourd'hui

        lignes = []
        if df_pros is not None and not df_pros.empty:
            for _, r in df_pros.iterrows():
                if r.get("Signal") == "Priorité":
                    prix = _num(r.get("Prix $"))
                    if prix is None:
                        continue
                    lignes.append([aujourdhui, "Achat", str(r.get("Symbole Brut") or ""), "Priorité",
                                   round(prix, 2), round(_num(r.get("Achat Rang"), 0), 1),
                                   round(_num(r.get("Pré G %"), 0), 1)])
        if df_port is not None and not df_port.empty:
            for _, r in df_port.iterrows():
                if r.get("Signal") == "Vendre":
                    prix = _num(r.get("Prix $"))
                    if prix is None:
                        continue
                    lignes.append([aujourdhui, "Vente", str(r.get("Symbole Brut") or ""), "Vendre",
                                   round(prix, 2), "", round(_num(r.get("Pré G %"), 0), 1)])
        # === Valeur du portefeuille + indices (courbe « vs marché ») ===
        v = _num(valeur_port)
        if v is not None and v > 0:
            lignes.append([aujourdhui, "Valeur", "PORTEFEUILLE", "", round(v, 2), "", ""])
        for sym_i, val_i in (indices or {}).items():
            vi = _num(val_i)
            if vi is not None and vi > 0:
                lignes.append([aujourdhui, "Valeur", sym_i, "", round(vi, 2), "", ""])
        if lignes:
            ws.append_rows(lignes, value_input_option="USER_ENTERED")
        return vals + lignes, bool(lignes)
    except Exception:
        return [], False

# === v7 : PRÉFÉRENCES PERSISTÉES (onglet « Config » du Sheet) =====================
# Les réglages de ⚙️ Paramètres sont enregistrés à chaque changement et rechargés à
# l'ouverture : plus besoin de tout recocher à chaque session.
# ==================================================================================
def charger_config_app():
    try:
        sh = connecter_google_sheets()
        ws = sh.worksheet("Config")
        vals = ws.get_all_values()
        return {r[0]: r[1] for r in vals[1:] if len(r) >= 2 and r[0]}
    except Exception:
        return {}

def sauvegarder_config_app(cfg):
    try:
        sh = connecter_google_sheets()
        try:
            ws = sh.worksheet("Config")
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet("Config", rows=80, cols=2)
        ws.clear()
        ws.update([["Paramètre", "Valeur"]] + [[k, v] for k, v in sorted(cfg.items())])
        return True
    except Exception:
        return False

if 'config_app' not in st.session_state:
    st.session_state['config_app'] = charger_config_app()
CFG_APP = st.session_state['config_app']

def pref_bool(nom, defaut):
    v = CFG_APP.get(nom)
    return (v == "1") if v in ("0", "1") else defaut

def pref_int(nom, defaut):
    try:
        return int(float(CFG_APP.get(nom)))
    except (TypeError, ValueError):
        return defaut

def pref_index(nom, options, defaut):
    v = CFG_APP.get(nom)
    return options.index(v) if v in options else defaut

# === v7 : CACHE YF (onglet « CacheYF » du Sheet) ==================================
# Sauvegarde les champs issus du .info Yahoo (chaleur, dividende, volatilité, secteur,
# fondamentaux) + le Rang d'achat. Deux usages :
#   1. SECOURS quand Yahoo bloque .info : les champs manquants sont repris du cache,
#      donc Score et Rang restent STABLES au lieu de changer artificiellement.
#   2. Le script bnc_alertes.py lit le Rang ici pour l'alerte « nouveau Top 5 ».
# ==================================================================================
ENTETE_CACHE_YF = ["Symbole", "Devise", "Possede", "Achat Rang", "Chaleur 52s", "Div %",
                   "Volatilité 1m", "Nb Analystes", "Secteur", "P/E", "Croiss Rev %",
                   "Marge %", "MAJ"]
CHAMPS_NUM_CACHE = ["Chaleur 52s", "Div %", "Volatilité 1m", "Nb Analystes", "P/E",
                    "Croiss Rev %", "Marge %"]

def charger_cache_yf():
    try:
        sh = connecter_google_sheets()
        ws = sh.worksheet("CacheYF")
        vals = ws.get_all_values()
        if not vals:
            return {}
        ent = vals[0]
        return {r[0]: dict(zip(ent, r)) for r in vals[1:] if r and r[0]}
    except Exception:
        return {}

def appliquer_cache_yf(df, cache):
    """Complète les champs .info MANQUANTS avec le cache (jamais d'écrasement)."""
    if not cache or df.empty:
        return df
    for idx, row in df.iterrows():
        sym = str(row.get('Symbole Brut') or '').strip()
        c = cache.get(sym)
        if not c:
            continue
        for col in CHAMPS_NUM_CACHE:
            if col in df.columns and pd.isna(row.get(col)):
                v = pd.to_numeric(pd.Series([c.get(col)]), errors='coerce').iloc[0]
                if pd.notna(v):
                    df.at[idx, col] = float(v)
        if 'Secteur' in df.columns and not str(row.get('Secteur') or '').strip():
            if str(c.get('Secteur') or '').strip():
                df.at[idx, 'Secteur'] = c['Secteur']
    return df

def sauvegarder_cache_yf(df_port, df_pros):
    try:
        sh = connecter_google_sheets()
        try:
            ws = sh.worksheet("CacheYF")
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet("CacheYF", rows=600, cols=len(ENTETE_CACHE_YF))
        maj = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M")
        lignes = [ENTETE_CACHE_YF]
        vus = set()
        for df_ in (df_pros, df_port):
            if df_ is None or df_.empty:
                continue
            for _, r in df_.iterrows():
                sym = str(r.get('Symbole Brut') or '').strip()
                if not sym or sym in vus:
                    continue
                vus.add(sym)
                def n(col):
                    v = pd.to_numeric(pd.Series([r.get(col)]), errors='coerce').iloc[0]
                    return "" if pd.isna(v) else round(float(v), 4)
                lignes.append([sym, str(r.get('Devise') or ''),
                               "1" if bool(r.get('Possede')) else "0",
                               n('Achat Rang'), n('Chaleur 52s'), n('Div %'),
                               n('Volatilité 1m'), n('Nb Analystes'),
                               str(r.get('Secteur') or ''), n('P/E'),
                               n('Croiss Rev %'), n('Marge %'), maj])
        ws.clear()
        ws.update(lignes)
        return True
    except Exception:
        return False

# --- TITRE (v8 : compact, sur la même ligne que les boutons) ---
heure_actuelle = heure_mise_a_jour()
taux_usdcad = obtenir_taux_change()

# === v7 : ORIENTATION mobile — portrait = mode mobile, paysage = mode complet ===
# L'orientation n'est connue que du navigateur : un petit script la synchronise dans
# l'URL (?orient=portrait|paysage) et recharge la page quand l'appareil pivote.
# Injecté UNIQUEMENT sur téléphone (sur PC, redimensionner la fenêtre ne doit
# surtout pas recharger la page).
if est_mobile():
    components.html("""
    <script>
    """ + JS_RECHARGER_PARENT + """
    (function() {
        function verifier() {
            try {
                const p = window.parent;
                const actuel = (p.innerWidth > p.innerHeight) ? 'paysage' : 'portrait';
                const url = new URL(p.location.href);
                if (url.searchParams.get('orient') !== actuel) {
                    url.searchParams.set('orient', actuel);
                    rechargerParent(url.toString());
                }
            } catch (e) { try { window.parent.console.log('BNC orient:', e.message); } catch (_) {} }
        }
        verifier();
        let t = null;
        window.parent.addEventListener('resize', function() {
            clearTimeout(t);
            t = setTimeout(verifier, 400);
        });
    })();
    </script>
    """, height=0)

def orientation_paysage():
    try:
        return st.query_params.get("orient", "") == "paysage"
    except Exception:
        return False

# Libellés compacts sur mobile (mode connu AVANT le popover grâce à la préférence
# persistée ; « Auto » retombe sur la détection User-Agent).
_mode_pref = CFG_APP.get('mode_affichage', 'Auto (détection)')
if _mode_pref == 'Auto (détection)':
    # téléphone EN PORTRAIT = mode mobile ; pivoté en PAYSAGE = affichage complet
    mobile_ui = est_mobile() and not orientation_paysage()
else:
    mobile_ui = (_mode_pref == 'Mobile (essentiel)')

# Favicon Les Affaires (24x24 PNG, intégré en base64 -> aucun chargement externe)
ICONE_LA_B64 = ("iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAFFElEQVR42p2WX2xcVxHGfzPn7HW89jppQiyHJP0nh7QxEYhYDZCAHQqtIiSQCmtaNVIEFfCCKKIUFaEqqgQUaEEu4gFVAgkJULUWVRtS9R9lbaUWNKRVnYQG5QGaqnYIdu3YsTe7e+85w8N603WgCeq8XOnee76Z851vvjlwmTCQEkVXBn/ptzJ4o+gMhHcTJXArXuQ8RycP5o9OHsyT85f/tyXkncCHIBww01v6d+710+c/E6vZh7CwvrHITUvevxLXX/XUM38dP/SASDRQgXjFBCVwQyJhrH/3rR3/nPqeX0z7kzRSj5FoDeIUaHNKPeeod+aOXujdcP/Hj4w/cyCaPnBJErmEcycq4XDv9vs635x70FUiVZFgllkkUyEHooAhRgRkFeasI8fSlg0P7Tpx9NulLLghCE1Mba1chFC+ru87607PPphVQliiFuq24KyQ88n2rUohr2qZSmOdR3BVkVBfqoe1f5+89/C2HT8cUgklim7FDpqcv3jTx27pPH762exCyFIqbtWe3bLxvq+z6vpryV93LSc/+XkWR0dxroCFeBHBBHMxhlw+8Yvvv+bW3UcOP9fEVECKEEsnSom8PjmsF1IyRbzmpXr8GNrVSaG3F+c8IoI16RQQ5xB1KCrmE9FqhNcnh0snSkmxcRai5YEBJ2Cb73rk5s6F7MYKRInmTD1xZopzTzwN0TCzBvcAoqgZaZijHmZJ4zxp9pZbjAuxsBBu3PzVR24WsPLAgPOMjYEITM/f5uuZCRIRFCJIgiZJQ4AturCsTvDGmjtuZ/XeT9DW00M6Ncn0Y09GPVQWzi7ehsrTjI2heyCzGL2EOFCPJiZogweB5aovHpYIRiB2d9L73GP0/mKYs7/+LZokbLhzH9v+MKK68wMS5+YHLES/BzIF+BtonFvyLaJaod3mQ9SRcZ51X7qd9XsGmX3+BRaefYozP/45FgIeKHx4J+GtGd9U6Ns9ry1o7+RNMeJZxfmDf+J0dw///tmjODyuey3iGsoUt7LTPEAfxJdWd2Scm0VwIPa/PcUiSp7aa8eYOnCa7rv2kfR9Ewp5LATEOcwM7erMWM6jZfDiNMP50UTVMFvR6iYtTC1TtHrfEDveOMb6r+znzE9/wtzvft/cQWzzzqyelSXJZWXwOlgsGtGI3V2Pp4kXw/Rts27UYTFiMUI9IL6Dq79/P8matcw8cYiFk8cbIrNGCakXqaSVx8kCgxTNy8hINJCX/zz8wlLPF14rVLNtFQiKOLNIcs1GRBUBZNM6YpYSaykAV99zN+19Wwj1SBQJClrf2vvqr0Llj2YmIhIVsBHQfulPZcumu2Pe400sDRVr/+gukhuuZ/7YBPPHJ+j63F78hh7+8eVvMD8xQTozjW/rZOpHw3bm0V9Smzojm/bv/9pIw+y0heGmTRPKvX3f6n5j/qHFWhqtq8Mspi7W6w2bbu9AqobV5rAkj+vIU5v7V0jw0u68zn1k+3cHXxz9QdOH/suuS+CGVMLolu33dL05/7BfqlJBo6DRWO5uAOeRECJEzeM1K3gWNq29d+DkKw+XikU3NDJy0a5XjLoRsJLhPj17dnzfTR8cl1r1fW2RzXlDxUwUxIMkZpJ3qrR7qV2VvHR+68YvDk4c+c2l4FccmZjJX3bt+pRMnfusVbMdFrP3EMFUZ1w+edm9d82T/ePjzyNirbS866F/yk61nbJTbXj3fw/9K0XrtWWFpZbBL0+uy15b/gNnp1KI5ff4ngAAAABJRU5ErkJggg==")

# === v8 : bouton 🤖 (PC seulement) -> la tâche planifiée Claude « Surveiller les
# affaires » (routine claude.ai). Un site web ne peut pas cliquer « Run now » à ta
# place : le bouton ouvre la PAGE DE LA TÂCHE dans un nouvel onglet — un seul clic
# sur Run now lance la surveillance.
URL_COWORK_AFF = "https://claude.ai/"   # TODO: URL exacte de la tâche (voir barre d'adresse)

# --- HAUT DE PAGE : Titre + Paramètres + Rafraîchir + Sheet + Affaires + LesAffaires.com ---
# (le CSS du bloc contenant stPopover force la rangée horizontale, même sur mobile)
if mobile_ui:
    col_titre, col_param, col_refresh, col_sheet, col_aff, col_la = st.columns(6)
    col_cw = None
else:
    col_titre, col_param, col_refresh, col_sheet, col_aff, col_la, col_cw = st.columns(7)
with col_titre:
    # Sur mobile : « 📈 » seul, sinon la rangée déborde et le bouton Sheet est coupé.
    st.markdown(
        f"<h3 style='margin: 0px; padding-top: 2px; white-space: nowrap;'>"
        f"{'📈' if mobile_ui else '📈 BNC LIVE v8'}</h3>",
        unsafe_allow_html=True
    )

with col_param:
    with st.popover("⚙️", help="Paramètres"):
        OPTIONS_GAIN = ["Yahoo", "Affaires", "Moyenne"]
        source_gain = st.selectbox("Calcul du Gain", OPTIONS_GAIN,
                                   index=pref_index('source_gain', OPTIONS_GAIN, 2))

        # Tri du Portefeuille (déplacé depuis l'entête de l'onglet ; persisté).
        # Pré G % se trie en ordre CROISSANT (proches de l'objectif en haut).
        OPTIONS_TRI_PORT = ["Pré G %", "Gain %"]
        colonne_tri = st.selectbox("Tri du Portefeuille", OPTIONS_TRI_PORT,
                                   index=pref_index('tri_portefeuille', OPTIONS_TRI_PORT, 1))

        # === v7 : mode d'affichage — sur téléphone, colonnes ESSENTIELLES seulement ===
        OPTIONS_MODE = ["Auto (détection)", "Ordinateur (complet)", "Mobile (essentiel)"]
        mode_affichage = st.selectbox(
            "Mode d'affichage", OPTIONS_MODE,
            index=pref_index('mode_affichage', OPTIONS_MODE, 0)
        )
        if mode_affichage == "Mobile (essentiel)":
            mode_mobile = True
        elif mode_affichage == "Ordinateur (complet)":
            mode_mobile = False
        else:
            # Auto : téléphone en PORTRAIT = mobile ; en PAYSAGE = complet
            mode_mobile = est_mobile() and not orientation_paysage()

        st.markdown("---")
        st.markdown("**Affichage des Colonnes**")
        afficher_no = st.checkbox("Afficher No.", value=pref_bool('afficher_no', True))
        afficher_desc = st.checkbox("Afficher Description", value=pref_bool('afficher_desc', True))
        afficher_dev = st.checkbox("Afficher Devise (Dev.)", value=pref_bool('afficher_dev', False))
        afficher_compte = st.checkbox("Afficher Compte", value=pref_bool('afficher_compte', True))

        afficher_var = st.checkbox("Afficher Var %", value=pref_bool('afficher_var', True))
        afficher_tendance = st.checkbox("Afficher Tendance (1m)", value=pref_bool('afficher_tendance', False))
        afficher_chaleur = st.checkbox("Afficher Chaleur 52 sem.", value=pref_bool('afficher_chaleur', False))
        afficher_div = st.checkbox("Afficher Dividendes (Div %)", value=pref_bool('afficher_div', False))
        afficher_analystes = st.checkbox("Afficher Nb d'analystes", value=pref_bool('afficher_analystes', False))  # === V4 ===

        st.markdown("---")
        st.markdown("**Moteur de décision (v5)**")
        afficher_signal = st.checkbox("Afficher Signal", value=pref_bool('afficher_signal', True))
        afficher_score = st.checkbox("Afficher Score", value=pref_bool('afficher_score', True))
        afficher_confiance = st.checkbox("Afficher Confiance", value=pref_bool('afficher_confiance', True))
        afficher_risque = st.checkbox("Afficher Risque", value=pref_bool('afficher_risque', True))
        afficher_volatilite = st.checkbox("Afficher Volatilité", value=pref_bool('afficher_volatilite', False))

        st.markdown("---")
        st.markdown("**Fonctionnalités Avancées**")
        activer_taux_change = st.checkbox("Taux de change actif", value=pref_bool('activer_taux_change', False))
        afficher_gain_jour = st.checkbox("Calculer le Gain du Jour", value=pref_bool('afficher_gain_jour', True))
        afficher_bandeau = st.checkbox("Afficher le Bandeau des Marchés", value=pref_bool('afficher_bandeau', False))
        afficher_alertes = st.checkbox("Activer les Alertes Intelligentes", value=pref_bool('afficher_alertes', False))
        rafraichir_auto = st.checkbox("Rafraîchir auto (à l'ouverture + aux 10 min en séance)",
                                      value=pref_bool('rafraichir_auto', True))

        # === V4 : garde-fou sur la fiabilité de l'objectif Yahoo ===
        # Un targetMeanPrice basé sur 1 seul analyste ne vaut rien. On peut exiger
        # un minimum d'analystes : en dessous, l'objectif Yahoo est ignoré.
        min_analystes = st.number_input(
            "Min. d'analystes pour l'objectif Yahoo (0 = désactivé)",
            min_value=0, max_value=50, value=pref_int('min_analystes', 0), step=1
        )

        # Pré Aff périmée : ignorée dans le calcul du Pré G % si la prévision (MAJ Aff)
        # date de plus de N mois. La valeur reste affichée mais grisée. 0 = jamais ignorée.
        mois_max_aff = st.number_input(
            "Ignorer Pré Aff si la prévision date de plus de (mois, 0 = jamais)",
            min_value=0, max_value=60, value=pref_int('mois_max_aff', 6), step=1
        )

        # === v7 : garde-fou anti-aberration ===
        # Un Pré G % énorme (> seuil) est presque toujours une erreur de données/échelle
        # (ex. cible $US vs prix CDR). Au-dessus, la valeur est grisée et neutralisée dans
        # le score / le rang d'achat. 0 = jamais.
        plafond_preg = st.number_input(
            "Signaler Pré G % au-dessus de (%, 0 = jamais)",
            min_value=0, max_value=1000, value=pref_int('plafond_preg', 200), step=25
        )

        st.markdown("---")
        st.markdown("**Aide à la décision (v7)**")
        afficher_rang = st.checkbox("Afficher Rang d'achat", value=pref_bool('afficher_rang', True))
        afficher_pourquoi = st.checkbox("Afficher Pourquoi", value=pref_bool('afficher_pourquoi', True))
        afficher_percentile = st.checkbox("Afficher Rang % (percentile)", value=pref_bool('afficher_percentile', False))
        afficher_concordance = st.checkbox("Afficher Concordance", value=pref_bool('afficher_concordance', False))
        afficher_entree = st.checkbox("Afficher Qualité d'entrée", value=pref_bool('afficher_entree', False))
        afficher_secteur = st.checkbox("Afficher Secteur", value=pref_bool('afficher_secteur', False))
        trier_par_rang = st.checkbox("Trier les Prospects par Rang d'achat", value=pref_bool('trier_par_rang', True))
        afficher_baisse = st.checkbox("Afficher Baisse depuis sommet 52s (Portefeuille)", value=pref_bool('afficher_baisse', False))
        seuil_baisse = st.number_input(
            "Alerte si baisse depuis le sommet dépasse (%)",
            min_value=5, max_value=50, value=pref_int('seuil_baisse', 15), step=5
        )
        afficher_fondamentaux = st.checkbox("Afficher Fondamentaux (P/E, croiss., marge)", value=pref_bool('afficher_fondamentaux', False))
        journaliser = st.checkbox("Journaliser les signaux (onglet Journal du Sheet)", value=pref_bool('journaliser', True))

    # === v7 : sauvegarde silencieuse des préférences quand elles changent ===
    cfg_courant = {
        'source_gain': source_gain, 'mode_affichage': mode_affichage,
        'tri_portefeuille': colonne_tri,
        'min_analystes': str(min_analystes), 'mois_max_aff': str(mois_max_aff),
        'plafond_preg': str(plafond_preg), 'seuil_baisse': str(seuil_baisse),
    }
    for nom_p, val_p in (('afficher_no', afficher_no), ('afficher_desc', afficher_desc),
                         ('afficher_dev', afficher_dev), ('afficher_compte', afficher_compte),
                         ('afficher_var', afficher_var), ('afficher_tendance', afficher_tendance),
                         ('afficher_chaleur', afficher_chaleur), ('afficher_div', afficher_div),
                         ('afficher_analystes', afficher_analystes), ('afficher_signal', afficher_signal),
                         ('afficher_score', afficher_score), ('afficher_confiance', afficher_confiance),
                         ('afficher_risque', afficher_risque), ('afficher_volatilite', afficher_volatilite),
                         ('activer_taux_change', activer_taux_change), ('afficher_gain_jour', afficher_gain_jour),
                         ('afficher_bandeau', afficher_bandeau), ('afficher_alertes', afficher_alertes),
                         ('afficher_rang', afficher_rang), ('afficher_pourquoi', afficher_pourquoi),
                         ('afficher_percentile', afficher_percentile), ('afficher_concordance', afficher_concordance),
                         ('afficher_entree', afficher_entree), ('afficher_secteur', afficher_secteur),
                         ('trier_par_rang', trier_par_rang), ('afficher_baisse', afficher_baisse),
                         ('afficher_fondamentaux', afficher_fondamentaux), ('journaliser', journaliser),
                         ('rafraichir_auto', rafraichir_auto)):
        cfg_courant[nom_p] = "1" if val_p else "0"
    if cfg_courant != {k: CFG_APP.get(k) for k in cfg_courant}:
        if sauvegarder_config_app(cfg_courant):
            st.session_state['config_app'] = dict(cfg_courant)
            CFG_APP = st.session_state['config_app']
            st.toast("⚙️ Préférences enregistrées.", icon="💾")

# === v7 : état des bourses (US et TSX : 9 h 30 – 16 h, heure de l'Est, jours ouvrables).
# La différence entre les deux vient des jours fériés propres à chaque pays (listes 2026).
FERIES_US_2026 = {"2026-01-01", "2026-01-19", "2026-02-16", "2026-04-03", "2026-05-25",
                  "2026-06-19", "2026-07-03", "2026-09-07", "2026-11-26", "2026-12-25"}
FERIES_CA_2026 = {"2026-01-01", "2026-02-16", "2026-04-03", "2026-05-18", "2026-07-01",
                  "2026-08-03", "2026-09-07", "2026-10-12", "2026-12-25", "2026-12-28"}

def statut_bourses():
    maintenant = datetime.now(ZoneInfo("America/Toronto"))
    date_jour = maintenant.strftime("%Y-%m-%d")
    en_heures = (9, 30) <= (maintenant.hour, maintenant.minute) < (16, 0)
    jour_ouvrable = maintenant.weekday() < 5
    ouvert_us = jour_ouvrable and en_heures and date_jour not in FERIES_US_2026
    ouvert_ca = jour_ouvrable and en_heures and date_jour not in FERIES_CA_2026
    return ouvert_us, ouvert_ca

def prochaine_ouverture():
    """Prochain moment où AU MOINS une des deux bourses ouvre (9 h 30 + 30 s de marge)."""
    maintenant = datetime.now(ZoneInfo("America/Toronto"))
    for j in range(0, 8):
        d = (maintenant + timedelta(days=j)).date()
        ds = d.strftime("%Y-%m-%d")
        if d.weekday() >= 5:
            continue
        if ds in FERIES_US_2026 and ds in FERIES_CA_2026:
            continue
        ouverture = datetime(d.year, d.month, d.day, 9, 30, 30, tzinfo=ZoneInfo("America/Toronto"))
        if ouverture > maintenant:
            return ouverture
    return None

def synchroniser_affaires():
    """v8 : bouton 📰 — copie Pré Aff / MAJ Aff depuis l'onglet « LesAffaires » vers
    l'onglet « Prospects » du MÊME classeur. Réutilise les helpers du script local
    (clé de symbole unifiée BBD.B/BBD-B.TO, la cible la plus récente gagne).
    Renvoie (nb mis à jour, nb vidés)."""
    from sync_affaires_vers_gsheet import parse_nombre, cle_symbole, date_key

    sh = connecter_google_sheets()
    src = sh.worksheet("LesAffaires").get_all_values()
    affaires = {}
    for row in src[1:]:
        if len(row) <= 3:
            continue
        sym = str(row[2]).strip().upper()          # C = Symbole
        if not sym:
            continue
        cible = parse_nombre(row[3])               # D = Cours cible
        date = str(row[0]).strip()                 # A = Date
        if cible is None:
            continue
        cle = cle_symbole(sym)
        ancien = affaires.get(cle)
        if ancien is None or date_key(date) >= date_key(ancien[0]):
            affaires[cle] = (date, cible)

    ws = sh.worksheet("Prospects")
    vals = ws.get_all_values()
    ent = [' '.join(str(h).split()) for h in vals[0]]
    i_sym, i_pa, i_maj = ent.index("Symbole"), ent.index("Pré Aff"), ent.index("MAJ Aff")

    updates, n_maj, n_vides = [], 0, 0
    for r, row in enumerate(vals[1:], start=2):
        if len(row) <= i_sym:
            continue
        sym = str(row[i_sym]).strip().upper()
        if not sym or sym == '0':
            continue
        entree = affaires.get(cle_symbole(sym))
        if entree:
            date, cible = entree
            updates.append({'range': gspread.utils.rowcol_to_a1(r, i_pa + 1), 'values': [[cible]]})
            updates.append({'range': gspread.utils.rowcol_to_a1(r, i_maj + 1), 'values': [[date]]})
            n_maj += 1
        else:
            # symbole absent de LesAffaires : on vide une éventuelle vieille valeur
            pa_actuel = str(row[i_pa]).strip() if len(row) > i_pa else ''
            if pa_actuel:
                updates.append({'range': gspread.utils.rowcol_to_a1(r, i_pa + 1), 'values': [['']]})
                updates.append({'range': gspread.utils.rowcol_to_a1(r, i_maj + 1), 'values': [['']]})
                n_vides += 1
    if updates:
        ws.batch_update(updates, value_input_option='USER_ENTERED')
    return n_maj, n_vides

def url_google_sheet():
    # URL du Google Sheet pour le bouton « Ouvrir Sheet ».
    # - construite depuis l'ID (Spreadsheet.url n'existe pas dans les vieux gspread) ;
    # - un échec n'est JAMAIS mis en cache (sinon le bouton disparaît pour 1 h) :
    #   on mémorise seulement le succès, et on réessaie à chaque rerun sinon.
    u = st.session_state.get('url_sheet', '')
    if u:
        return u
    try:
        sh = connecter_google_sheets()
        u = getattr(sh, 'url', '') or ''
        if not u:
            u = f"https://docs.google.com/spreadsheets/d/{sh.id}"
    except Exception:
        return ""
    st.session_state['url_sheet'] = u
    return u

if col_refresh.button("🔄", help=f"Rafraîchir (dernière heure : {heure_actuelle})"):
    st.cache_data.clear()
    st.rerun()
url_sheet = url_google_sheet()
if url_sheet:
    col_sheet.link_button("📗", url_sheet, help="Ouvrir le Google Sheet")
if col_aff.button("📰", help="Importer Les Affaires (onglet LesAffaires → Prospects)"):
    try:
        with st.spinner("Import Les Affaires..."):
            n_maj, n_vides = synchroniser_affaires()
        st.toast(f"📰 Les Affaires : {n_maj} mis à jour, {n_vides} vidé(s).", icon="✅")
        st.cache_data.clear()   # recharge les Pré Aff fraîches
        st.rerun()
    except Exception as e:
        st.error(f"Import Les Affaires : {type(e).__name__} - {e}")
if col_cw is not None:
    # PC seulement : ouvre la page des tâches Claude -> cliquer « Run now » sur
    # « Surveiller les affaires ».
    col_cw.link_button("🤖", URL_COWORK_AFF,
                       help="Tâche Claude « Surveiller les affaires » (cliquer Run now)")
with col_la:
    # Lien vers la page « À surveiller » de LesAffaires.com, avec leur icône,
    # habillé pour ressembler aux autres boutons de la rangée.
    st.markdown(
        f"<a href='https://www.lesaffaires.com/bourse/a-surveiller/' target='_blank' "
        f"title='Les Affaires — titres à surveiller' "
        f"style='display: inline-flex; align-items: center; justify-content: center; "
        f"height: 38px; min-width: 46px; padding: 0 4px; "
        f"border: 1px solid rgba(250, 250, 250, .2); border-radius: 8px; "
        f"text-decoration: none;'>"
        f"<img src='data:image/png;base64,{ICONE_LA_B64}' width='20' height='20' "
        f"style='border-radius: 3px; display: block;'/></a>",
        unsafe_allow_html=True
    )

# Indicateur d'état des bourses (sous la rangée Paramètres / boutons)
ouvert_us, ouvert_ca = statut_bourses()

# === v7 : rafraîchissement AUTO ===
# - Bourse OUVERTE : recharge toutes les 10 min (cache 5 min expiré -> données fraîches).
# - Bourse FERMÉE : si l'app est ouverte avant 9 h 30 (fenêtre max 3 h), recharge pile
#   à l'ouverture. Un onglet oublié la veille ne recharge pas au milieu de la nuit.
note_auto = ""
if rafraichir_auto:
    if ouvert_us or ouvert_ca:
        # Rechargement aux 5 min ; l'ordonnanceur (jetons_fetch) ne rafraîchit qu'UN
        # groupe par rechargement : Portefeuille 10 min, Pros CAD 25 min, Pros US 35 min.
        components.html(
            "<script>" + JS_RECHARGER_PARENT +
            "setTimeout(function() { rechargerParent(window.parent.location.href); }, "
            f"{5 * 60 * 1000});</script>",
            height=0
        )
        if not mobile_ui:   # v8 : note masquée sur mobile (le minuteur tourne quand même)
            note_auto = (" &nbsp;·&nbsp; ⏱ <span style='color: gray;'>auto : Portef. 10 min · "
                         "Pros CAD 25 min · Pros US 35 min</span>")
    else:
        _prochaine = prochaine_ouverture()
        if _prochaine is not None:
            _delai_s = (_prochaine - datetime.now(ZoneInfo("America/Toronto"))).total_seconds()
            if 0 < _delai_s <= 3 * 3600:
                components.html(
                    "<script>" + JS_RECHARGER_PARENT +
                    "setTimeout(function() { rechargerParent(window.parent.location.href); }, "
                    f"{int(_delai_s * 1000)});</script>",
                    height=0
                )
                if not mobile_ui:   # v8 : note masquée sur mobile (le minuteur tourne quand même)
                    note_auto = (" &nbsp;·&nbsp; ⏱ <span style='color: gray;'>rafraîchissement auto à "
                                 f"{_prochaine.strftime('%H:%M')}</span>")

# (v8 : 💵/🍁 au lieu des drapeaux emoji — Windows ne rend pas 🇺🇸/🇨🇦 ;
#  les heures d'ouverture ne sont affichées que sur PC pour économiser l'espace mobile)
heures_txt = "" if mobile_ui else " <span style='color: gray;'>(9 h 30 – 16 h, heure de l'Est)</span>"
st.markdown(
    f"<div style='font-size: 13px; margin-top: -6px; margin-bottom: 6px;'>"
    f"💵 Bourse US : {'🟢 <b>Ouverte</b>' if ouvert_us else '🔴 Fermée'}"
    f" &nbsp;·&nbsp; 🍁 TSX : {'🟢 <b>Ouverte</b>' if ouvert_ca else '🔴 Fermée'}"
    f"{heures_txt}{note_auto}</div>",
    unsafe_allow_html=True
)

# === v7 : dernière synchro Yahoo RÉUSSIE par groupe (rempli après la phase 2) ===
ph_sync = st.empty()

@st.cache_resource
def _dernieres_synchros():
    # Survit aux rechargements de page ; un blocage Yahoo n'écrase pas la dernière réussite.
    return {}

def _fmt_sync(h):
    if not h:
        return "—"
    auj = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d")
    return h[len(auj) + 1:] if h.startswith(auj + " ") else h   # aujourd'hui -> HH:MM seul

# --- MOTEUR TURBO ---
# === V4 : récupération Yahoo refondue ===========================================
# Avant : 1 appel .history() + 1 appel .info PAR symbole, 10 threads en parallèle.
#         -> .info se faisait régulièrement bloquer par Yahoo (429), d'où des trous
#            intermittents dans Pré 1an / Div % / Chaleur 52s.
# Maintenant :
#   1) PRIX & HISTORIQUE : un SEUL appel groupé yf.download() pour tous les symboles.
#                          === v5 : période 1 mois (au lieu de 5j) pour la volatilité.
#   2) INFOS (.info)     : appel plus lourd, donc moins de threads (4 au lieu de 10)
#                          pour réduire fortement le risque de blocage.
# L'interface de sortie reste identique : { sym: {'hist': df, 'info': dict} }
# ================================================================================
@st.cache_data(ttl=28800, show_spinner=False)
def telecharger_yahoo(groupes, retry_premier=False, jeton=None):
    # `jeton` ne sert qu'à contrôler le CACHE par groupe (clé de cache) :
    #  - mode normal : jeton = tranche de 5 min -> comportement classique (ttl 5 min) ;
    #  - mode auto en séance : jeton = compteur de version par groupe, incrémenté par
    #    jetons_fetch() quand SON intervalle est écoulé (P1 10 min, P2 25, P3 35).
    # === v6 : récupération PAR PRIORITÉ, appelable par phase ===
    # groupes = liste ORDONNÉE de groupes de symboles (priorité décroissante).
    # PRIX : un seul appel groupé pour TOUS (robuste, jamais bloqué).
    # .info : récupéré groupe par groupe ; si Yahoo bloque (429 -> .info quasi vide),
    # on s'ARRÊTE pour les groupes suivants (reprises plus tard).
    # retry_premier : ré-essaie les manquants du 1er groupe après une pause (Portefeuille).
    groupes = [tuple(g) for g in groupes]
    tous = list(dict.fromkeys(s for g in groupes for s in g))
    resultats = {sym: {'hist': pd.DataFrame(), 'info': {}} for sym in tous}

    # --- 1) PRIX & HISTORIQUE : un seul appel groupé (rapide, robuste) ---
    if tous:
        try:
            data = yf.download(
                tous, period="1mo", interval="1d",
                group_by="ticker", threads=True, progress=False, auto_adjust=True
            )
            for sym in tous:
                try:
                    hist = data if len(tous) == 1 else data[sym]
                    hist = hist.dropna(how="all")
                    if not hist.empty:
                        resultats[sym]['hist'] = hist
                except Exception:
                    pass
        except Exception:
            pass

    # --- 2) INFOS (.info) PAR PRIORITÉ ---
    def fetch_info(sym):
        try:
            info = yf.Ticker(sym).info or {}
            # un vrai .info a des dizaines de clés ; une réponse 429/échec en a très peu
            return sym, info, (len(info) > 5)
        except Exception:
            return sym, {}, False

    def recuperer(groupe):
        echecs = 0
        with ThreadPoolExecutor(max_workers=4) as executor:
            for future in as_completed([executor.submit(fetch_info, s) for s in groupe]):
                sym, info, ok = future.result()
                if ok:
                    resultats[sym]['info'] = info
                else:
                    echecs += 1
        return echecs

    niveaux_ok = []
    bloque = False
    for niveau, groupe in enumerate(groupes, 1):
        if bloque:
            break
        if not groupe:
            niveaux_ok.append(niveau)
            continue
        echecs = recuperer(list(groupe))
        # 1er groupe (Portefeuille en phase 1) : on retente les manquants après une pause.
        if niveau == 1 and retry_premier and echecs > 0:
            time.sleep(3)
            manquants = [s for s in groupe if len(resultats[s]['info']) <= 5]
            recuperer(manquants)
            echecs = sum(1 for s in groupe if len(resultats[s]['info']) <= 5)
        # Trop d'échecs sur ce groupe => Yahoo bloque -> on saute les priorités SUIVANTES.
        if len(groupe) >= 5 and echecs > len(groupe) * 0.5:
            bloque = True
        else:
            niveaux_ok.append(niveau)

    resultats['__statut__'] = {'niveaux_ok': niveaux_ok, 'bloque': bloque}
    # Horodatage du TÉLÉCHARGEMENT réel : voyage avec le cache (un cache hit renvoie
    # l'heure de la récupération d'origine, pas celle du rechargement de page).
    resultats['__horodatage__'] = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M")
    return resultats

# === V4 : plus de @st.cache_data ici =============================================
# construire_donnees recevait un DataFrame ET un dict contenant des DataFrames :
# Streamlit devait les HASHER à chaque appel (coûteux + cache-miss silencieux).
# Comme telecharger_tous_les_prix_yahoo est déjà caché, on recalcule simplement.
# ================================================================================
def _meme_societe(nom_cad, nom_us):
    # Garde-fou anti-collision pour la règle de trois : confirme que le ticker CAD et le
    # ticker US déduit désignent la MÊME société (via le nom Yahoo). Pour les vraies
    # interlistées/CDR, le longName est identique (« Alphabet Inc. »).
    def norm(x):
        return ' '.join(str(x or '').lower().replace('.', ' ').replace(',', ' ').split())
    a, b = norm(nom_cad), norm(nom_us)
    if len(a) < 3 or len(b) < 3:
        return False
    return a == b or a in b or b in a

def construire_donnees(df, dict_yahoo, est_portefeuille=True, symboles_portefeuille=None, plafond_scaling=200):
    df = df.copy()
    if 'Description' not in df.columns:   # === Description : garantit que la colonne existe ===
        df['Description'] = ""
    df['Devise'] = 'USD'
    df['Possede'] = False
    df['Pré 1an $ Yahoo'] = np.nan
    df['Chaleur 52s'] = np.nan
    df['Div %'] = np.nan
    df['Nb Analystes'] = np.nan  # === V4 ===
    df['Volatilité 1m'] = np.nan  # === v5 ===
    df['Données OK'] = False       # === v5 : prix bien récupéré ? ===
    df['Gain Jour $'] = 0.0
    df['Symbole Brut'] = ""
    df['Secteur'] = ""      # === v7 : secteur Yahoo (diversification) ===
    df['Baisse 52s %'] = np.nan   # === v7 : % depuis le sommet 52 sem (protection des gains) ===
    df['P/E'] = np.nan            # === v7 : fondamentaux légers ===
    df['Croiss Rev %'] = np.nan
    df['Marge %'] = np.nan
    tendances = []

    for index, row in df.iterrows():
        symbole = row.get('Symbole')
        if pd.notna(symbole) and str(symbole).strip() not in ("", "0"):
            symbole_clean = str(symbole).strip()
            df.at[index, 'Symbole Brut'] = symbole_clean

            # === V4 : détection devise par suffixe réel (.endswith) au lieu de "in" ===
            if symbole_clean.endswith(('.TO', '.V', '.NE', '.CN')):
                df.at[index, 'Devise'] = 'CAD'
            else:
                df.at[index, 'Devise'] = 'USD'

            if symboles_portefeuille and symbole_clean in symboles_portefeuille:
                df.at[index, 'Possede'] = True

            donnees_y = dict_yahoo.get(symbole_clean, {})
            infos = donnees_y.get('hist', pd.DataFrame())
            infos_gen = donnees_y.get('info', {})

            prix_actuel = None

            if not infos.empty and len(infos) >= 2:
                serie_close = infos['Close'].dropna()
                prix_actuel = serie_close.iloc[-1]
                prix_veille = serie_close.iloc[-2]

                df.at[index, 'Prix $'] = prix_actuel
                df.at[index, 'Var %'] = (prix_actuel - prix_veille) / prix_veille
                df.at[index, 'Données OK'] = True

                tendances.append(serie_close.tolist())

                # === v5 : volatilité annualisée (écart-type des rendements * sqrt(252)) ===
                rendements = serie_close.pct_change().dropna()
                if len(rendements) >= 5:
                    df.at[index, 'Volatilité 1m'] = float(rendements.std() * math.sqrt(252) * 100)

                if est_portefeuille and 'Achat $' in row and pd.notna(row['Achat $']) and str(row['Achat $']).strip() != "":
                    achat = float(row['Achat $'])
                    qte = float(row['Qtée']) if 'Qtée' in row and pd.notna(row['Qtée']) and str(row['Qtée']).strip() != "" else 0
                    df.at[index, 'Gain %'] = (prix_actuel - achat) / achat
                    df.at[index, 'Gain $'] = (prix_actuel - achat) * qte
                    df.at[index, 'Gain Jour $'] = (prix_actuel - prix_veille) * qte
            else:
                tendances.append(None)

            prevision_1an = infos_gen.get('targetMeanPrice')
            if prevision_1an is not None:
                df.at[index, 'Pré 1an $ Yahoo'] = prevision_1an
            elif symbole_clean.endswith(('.TO', '.V', '.NE', '.CN')) and prix_actuel is not None and prix_actuel > 0:
                # === Action CAD sans objectif Yahoo : on emprunte celui du ticker US ===
                # équivalent (souvent déjà dans les Prospects, donc déjà téléchargé), mis à
                # l'échelle CAD par règle de trois sur les prix actuels :
                #   Pré YF_CAD = Objectif_US × (Prix_CAD / Prix_US)
                us_candidat = symbole_clean
                for suff in ('.TO', '.V', '.NE', '.CN'):
                    if us_candidat.endswith(suff):
                        us_candidat = us_candidat[:-len(suff)]
                        break
                donnees_us = dict_yahoo.get(us_candidat, {})
                info_us = donnees_us.get('info', {})
                cible_us = info_us.get('targetMeanPrice')
                hist_us = donnees_us.get('hist', pd.DataFrame())
                # Garde-fou anti-collision : même société (nom Yahoo) côté CAD et US.
                # Si le .info du ticker CAD manque (blocage Yahoo), on se rabat sur la
                # Description du Sheet (toujours disponible).
                nom_cad = (infos_gen.get('longName') or infos_gen.get('shortName')
                           or row.get('Description'))
                nom_us = info_us.get('longName') or info_us.get('shortName')
                if (cible_us is not None and not hist_us.empty and 'Close' in hist_us.columns
                        and _meme_societe(nom_cad, nom_us)):
                    close_us = hist_us['Close'].dropna()
                    if len(close_us) >= 1 and float(close_us.iloc[-1]) > 0:
                        prix_us = float(close_us.iloc[-1])
                        df.at[index, 'Pré 1an $ Yahoo'] = float(cible_us) * (float(prix_actuel) / prix_us)

            # === v7 : cibles en $US ? Mise à l'échelle CAD par PLAUSIBILITÉ, appliquée aux
            # TROIS cibles : Pré Aff (synchro Affaires par symbole de base), cible Yahoo LIVE
            # (Yahoo publie parfois la cible $US de la société sur la page du CDR .TO !) et
            # Pré YF du Sheet (peut contenir une ancienne valeur non convertie).
            # Règle : convertir par prix_CAD/prix_US UNIQUEMENT si le gain brut dépasse le
            # seuil ET que le gain converti redevient plausible (ne touche jamais une cible
            # déjà en CAD, ex. MDA.TO).
            if symbole_clean.endswith(('.TO', '.V', '.NE', '.CN')) and prix_actuel is not None and prix_actuel > 0:
                seuil = float(plafond_scaling) if plafond_scaling and plafond_scaling > 0 else 200.0
                prix_f = float(prix_actuel)
                ratio_cad = None   # None = pas encore calculé ; False = indisponible

                def _ratio_us():
                    us_c = symbole_clean
                    for suff in ('.TO', '.V', '.NE', '.CN'):
                        if us_c.endswith(suff):
                            us_c = us_c[:-len(suff)]
                            break
                    d_us = dict_yahoo.get(us_c, {})
                    h_us = d_us.get('hist', pd.DataFrame())
                    i_us = d_us.get('info', {})
                    # nom Yahoo, sinon Description du Sheet (le .info du CAD peut manquer)
                    nom_c = (infos_gen.get('longName') or infos_gen.get('shortName')
                             or row.get('Description'))
                    nom_u = i_us.get('longName') or i_us.get('shortName')
                    if h_us.empty or 'Close' not in h_us.columns or not _meme_societe(nom_c, nom_u):
                        return False
                    c_us = h_us['Close'].dropna()
                    if len(c_us) < 1 or float(c_us.iloc[-1]) <= 0:
                        return False
                    return prix_f / float(c_us.iloc[-1])

                for col_cible in ('Pré Aff', 'Pré 1an $ Yahoo', 'Pré YF'):
                    if col_cible not in df.columns:
                        continue
                    num = pd.to_numeric(pd.Series([df.at[index, col_cible]]), errors='coerce').iloc[0]
                    if pd.isna(num) or num <= 0:
                        continue
                    if (float(num) - prix_f) / prix_f * 100 <= seuil:
                        continue   # plausible -> on ne touche pas
                    if ratio_cad is None:
                        ratio_cad = _ratio_us()
                    if not ratio_cad:
                        continue
                    conv = float(num) * ratio_cad
                    if (conv - prix_f) / prix_f * 100 <= seuil:
                        df.at[index, col_cible] = conv

            # === V4 : nombre d'analystes derrière l'objectif (fiabilité du signal) ===
            nb_analystes = infos_gen.get('numberOfAnalystOpinions')
            if nb_analystes is not None:
                df.at[index, 'Nb Analystes'] = nb_analystes

            # === Description : nom de l'entreprise depuis Yahoo si la cellule est vide ===
            # On ne remplit QUE les cases vides : toute description saisie à la main est conservée.
            desc_existante = row.get('Description')
            if pd.isna(desc_existante) or str(desc_existante).strip() == "":
                nom_entreprise = infos_gen.get('longName') or infos_gen.get('shortName')
                if nom_entreprise:
                    df.at[index, 'Description'] = str(nom_entreprise)

            # === V4 : correctif dividende ===
            # yfinance renvoie selon les versions une FRACTION (0.025) ou déjà un % (2.5).
            # Heuristique : une valeur < 1 est presque toujours une fraction -> *100.
            div_yield = infos_gen.get('dividendYield')
            if div_yield is not None and div_yield > 0:
                df.at[index, 'Div %'] = div_yield * 100 if div_yield < 1 else div_yield

            low_52 = infos_gen.get('fiftyTwoWeekLow')
            high_52 = infos_gen.get('fiftyTwoWeekHigh')
            if low_52 is not None and high_52 is not None and prix_actuel is not None:
                if high_52 > low_52:
                    chaleur = ((prix_actuel - low_52) / (high_52 - low_52)) * 100
                    df.at[index, 'Chaleur 52s'] = max(0, min(100, chaleur))
                else:
                    df.at[index, 'Chaleur 52s'] = 50.0

            # === v7 : baisse depuis le sommet 52 sem (négatif = sous le sommet) ===
            if high_52 is not None and prix_actuel is not None and float(high_52) > 0:
                df.at[index, 'Baisse 52s %'] = (float(prix_actuel) - float(high_52)) / float(high_52) * 100

            # === v7 : fondamentaux légers (déjà présents dans le .info téléchargé) ===
            pe = infos_gen.get('trailingPE') or infos_gen.get('forwardPE')
            try:
                if pe is not None and float(pe) > 0:
                    df.at[index, 'P/E'] = float(pe)
            except (ValueError, TypeError):
                pass
            rg = infos_gen.get('revenueGrowth')       # fraction (0.15 = +15 %)
            try:
                if rg is not None:
                    df.at[index, 'Croiss Rev %'] = float(rg) * 100
            except (ValueError, TypeError):
                pass
            pm = infos_gen.get('profitMargins')       # fraction
            try:
                if pm is not None:
                    df.at[index, 'Marge %'] = float(pm) * 100
            except (ValueError, TypeError):
                pass

            devise_off = infos_gen.get('currency')
            if devise_off:
                df.at[index, 'Devise'] = str(devise_off).upper()

            # === v7 : secteur (pour la diversification vs portefeuille) ===
            secteur = infos_gen.get('sector')
            if secteur:
                df.at[index, 'Secteur'] = str(secteur)

            df.at[index, 'Symbole'] = f"https://ca.finance.yahoo.com/quote/{symbole_clean}"
        else:
            tendances.append(None)

    df['Tendance'] = tendances
    return df

def calculer_potentiel_gain(df, source, est_portefeuille=True, min_analystes=0, mois_max_aff=0, plafond_preg=0):  # === V4 : min_analystes ; v6 : mois_max_aff ; v7 : plafond_preg ===
    df = df.copy()
    if 'Prix $' not in df.columns:
        return df

    prix = pd.to_numeric(df['Prix $'], errors='coerce')

    # Récupération ultra-sécurisée des prévisions
    if 'Pré 1an $ Yahoo' in df.columns:
        yahoo_live = pd.to_numeric(df['Pré 1an $ Yahoo'], errors='coerce')
    else:
        yahoo_live = pd.Series(np.nan, index=df.index)

    if 'Pré YF' in df.columns:
        yahoo_base = pd.to_numeric(df['Pré YF'], errors='coerce')
    else:
        yahoo_base = pd.Series(np.nan, index=df.index)

    yahoo = yahoo_live.fillna(yahoo_base)

    # === V4 : on ignore l'objectif Yahoo s'il repose sur trop peu d'analystes ===
    if min_analystes and min_analystes > 0 and 'Nb Analystes' in df.columns:
        nb = pd.to_numeric(df['Nb Analystes'], errors='coerce').fillna(0)
        yahoo = yahoo.where(nb >= min_analystes, np.nan)

    if 'Pré Aff' in df.columns:
        affaires = pd.to_numeric(df['Pré Aff'], errors='coerce').replace(0, np.nan)
    else:
        affaires = pd.Series(np.nan, index=df.index)

    # === v6 : Pré Aff PÉRIMÉE — ignorée dans le calcul si MAJ Aff date de plus de N mois ===
    perime = pd.Series(False, index=df.index)
    if mois_max_aff and mois_max_aff > 0 and 'MAJ Aff' in df.columns:
        dates_aff = pd.to_datetime(df['MAJ Aff'], errors='coerce')
        seuil = pd.Timestamp.now() - pd.DateOffset(months=int(mois_max_aff))
        perime = dates_aff.notna() & (dates_aff < seuil) & affaires.notna()
    affaires_calc = affaires.where(~perime, np.nan)   # ignorée dans le calcul si périmée

    # === v7 : cible individuellement ABERRANTE écartée de la moyenne ===
    # Si UNE des deux cibles implique un gain invraisemblable (> plafond) alors que
    # l'AUTRE est plausible, l'aberrante est exclue du calcul (et grisée pour Pré Aff)
    # au lieu de polluer la moyenne. Filet de sécurité quand la conversion CAD n'a pas
    # pu s'exécuter (ex. NVDA.TO : Aff 300 $US non convertie vs Yahoo 67 $ plausible).
    yahoo_calc = yahoo
    if plafond_preg and plafond_preg > 0:
        seuil_ab = float(plafond_preg)
        gain_y = (yahoo - prix) / prix * 100
        gain_a = (affaires_calc - prix) / prix * 100
        y_ab = yahoo.notna() & (prix > 0) & (gain_y > seuil_ab)
        a_ab = affaires_calc.notna() & (prix > 0) & (gain_a > seuil_ab)
        excl_a = a_ab & yahoo.notna() & ~y_ab
        excl_y = y_ab & affaires_calc.notna() & ~a_ab
        affaires_calc = affaires_calc.where(~excl_a, np.nan)
        yahoo_calc = yahoo.where(~excl_y, np.nan)
        perime = perime | excl_a   # grise la Pré Aff écartée à l'affichage

    if source == "Yahoo":
        cible = yahoo_calc.fillna(affaires_calc)
    elif source == "Affaires":
        cible = affaires_calc.fillna(yahoo_calc)
    else:
        temp = pd.DataFrame({'Y': yahoo_calc, 'A': affaires_calc})
        cible = temp.mean(axis=1, skipna=True)

    mask = (prix > 0) & cible.notna()
    df.loc[mask, 'Pré G %'] = (cible[mask] - prix[mask]) / prix[mask]

    # === v7 : garde-fou anti-aberration — Pré G % > seuil = presque toujours une erreur
    # de données/échelle. On le marque (grisé à l'affichage, neutralisé dans le score/rang).
    preg = pd.to_numeric(df.get('Pré G %'), errors='coerce')  # fraction ici (×100 plus tard)
    if plafond_preg and plafond_preg > 0:
        df['Pré G Aberrant'] = preg.notna() & (preg * 100 > float(plafond_preg))
    else:
        df['Pré G Aberrant'] = pd.Series(False, index=df.index)

    # Enregistrement pour l'affichage (valeur Pré Aff RÉELLE, grisée si périmée)
    df['Pré YF Display'] = yahoo
    df['Pré Aff Display'] = affaires
    df['Pré Aff Périmé'] = perime

    return df

# === v5 : MOTEUR DE DÉCISION (Score / Confiance / Risque / Signal) ==============
# Inspiré de la version Codex, mais CORRIGÉ : on calcule sur des valeurs déjà en %
# (Pré G %, Var %, Div %, Volatilité, Chaleur), donc les seuils ci-dessous ont du
# sens. De plus, le Score est normalisé par les POIDS DISPONIBLES : une donnée
# manquante ne tire pas le score vers le bas, elle est simplement ignorée.
# ================================================================================
def pct_vers_points(valeur, bas, haut):
    """Convertit une valeur en note 0..100 selon une échelle [bas, haut]."""
    if pd.isna(valeur):
        return np.nan
    if haut == bas:
        return 0.0
    return float(np.clip((valeur - bas) / (haut - bas) * 100, 0, 100))

def classifier_signal(score, confiance, risque, potentiel):
    if pd.isna(score):
        return "Données insuffisantes"
    if confiance < 45:
        return "À valider"
    if risque >= 75 and score < 80:
        return "Risque élevé"
    if pd.notna(potentiel) and potentiel <= 0:
        return "Objectif atteint"
    if score >= 75 and confiance >= 60:
        return "Priorité"
    if score >= 60:
        return "À surveiller"
    return "Secondaire"

def signal_portefeuille(potentiel):
    # Signal orienté DÉCISION DE VENTE pour un titre détenu, basé sur le potentiel
    # restant (Pré G %). Aligné sur les seuils de couleur (≤5 % vendre, <15 % surveiller).
    if pd.isna(potentiel):
        return ""
    if potentiel <= 5:
        return "Vendre"
    if potentiel < 15:
        return "À surveiller"
    return "Attendre"

def _pourquoi_achat(r):
    # === v7 : explication lisible des facteurs saillants pour choisir un achat ===
    bits = []
    p = r.get("Pré G %")
    if pd.notna(p):
        if r.get("Pré G Aberrant"):
            bits.append("⚠️ potentiel invraisemblable")
        elif p > 0:
            bits.append(f"+{p:.0f}% potentiel")
        else:
            bits.append("objectif atteint")
    c = r.get("Concordance")
    if pd.notna(c):
        if c >= 70:
            bits.append("2 sources concordantes")
        elif c < 40:
            bits.append("⚠️ cibles divergentes")
    ch = r.get("Chaleur 52s")
    if pd.notna(ch):
        if ch <= 25:
            bits.append("proche creux 52s")
        elif ch >= 85:
            bits.append("proche sommet 52s")
    rq = r.get("Risque")
    if pd.notna(rq) and rq >= 75:
        bits.append("risque élevé")
    conc = r.get("Concentration Secteur")
    sect = r.get("Secteur")
    if pd.notna(conc) and conc and conc >= 2 and pd.notna(sect) and str(sect).strip():
        bits.append(f"secteur chargé ({sect}×{int(conc)})")
    return " · ".join(bits)

def calculer_score_decision(df, pour_portefeuille=False, secteurs_portefeuille=None):
    df = df.copy()

    potentiel = pd.to_numeric(df.get("Pré G %"), errors="coerce")     # en %
    chaleur = pd.to_numeric(df.get("Chaleur 52s"), errors="coerce")   # 0..100
    dividende = pd.to_numeric(df.get("Div %"), errors="coerce")       # en %
    volatilite = pd.to_numeric(df.get("Volatilité 1m"), errors="coerce")  # en %
    var_jour = pd.to_numeric(df.get("Var %"), errors="coerce")        # en %
    yahoo = pd.to_numeric(df.get("Pré YF Display"), errors="coerce")  # cible $ Yahoo
    affaires = pd.to_numeric(df.get("Pré Aff Display"), errors="coerce")  # cible $ Affaires

    # === v7 : neutralise le potentiel ABERRANT dans le score/rang (mais pas à l'affichage) ===
    aberrant = df.get("Pré G Aberrant")
    if aberrant is None:
        aberrant = pd.Series(False, index=df.index)
    aberrant = aberrant.fillna(False).astype(bool)
    potentiel = potentiel.where(~aberrant)

    donnees_ok = df.get("Données OK")
    if donnees_ok is None:
        donnees_ok = pd.Series(False, index=df.index)

    # --- Composantes du score (0..100) ---
    pts_potentiel = potentiel.apply(lambda v: pct_vers_points(v, 0, 80))
    pts_creux = chaleur.apply(lambda v: 100 - pct_vers_points(v, 0, 100) if pd.notna(v) else np.nan)
    pts_div = dividende.apply(lambda v: pct_vers_points(v, 0, 6))
    pts_momentum = var_jour.apply(lambda v: pct_vers_points(v, -5, 5))
    penalite_vol = volatilite.apply(lambda v: pct_vers_points(v, 15, 80))

    composantes = pd.DataFrame({
        "potentiel": pts_potentiel,
        "creux": pts_creux,
        "div": pts_div,
        "momentum": pts_momentum,
    })
    somme_ponderee = (
        composantes["potentiel"].fillna(0) * 0.50
        + composantes["creux"].fillna(0) * 0.18
        + composantes["div"].fillna(0) * 0.12
        + composantes["momentum"].fillna(50) * 0.08
        + (100 - penalite_vol.fillna(50)) * 0.12
    )
    # Normalisation par les poids RÉELLEMENT disponibles (clé de la robustesse).
    # Momentum ET volatilité sont TOUJOURS comptés (défaut neutre 50) : leur valeur
    # par défaut est ajoutée au numérateur, donc leur poids doit l'être au dénominateur
    # — sinon le score peut dépasser 100 quand presque tout est manquant.
    poids_disponible = (
        composantes["potentiel"].notna() * 0.50
        + composantes["creux"].notna() * 0.18
        + composantes["div"].notna() * 0.12
        + 0.08  # momentum : toujours compté (défaut neutre 50)
        + 0.12  # volatilité : toujours comptée (défaut neutre 50)
    )
    df["Score"] = np.where(poids_disponible > 0, somme_ponderee / poids_disponible, np.nan)

    # === v7 : Qualité du point d'ENTRÉE (potentiel + proche creux 52s + momentum court) ===
    entree_num = (pts_potentiel.fillna(0) * 0.45 + pts_creux.fillna(0) * 0.35 + pts_momentum.fillna(50) * 0.20)
    entree_den = (pts_potentiel.notna() * 0.45 + pts_creux.notna() * 0.35 + 0.20)
    df["Entrée"] = np.where(entree_den > 0, entree_num / entree_den, np.nan)

    # --- Confiance : complétude des données + accord des cibles ---
    confiance = pd.Series(20.0, index=df.index)
    confiance += donnees_ok.astype(bool) * 20
    confiance += yahoo.notna() * 20
    confiance += affaires.notna() * 15
    confiance += chaleur.notna() * 10
    confiance += volatilite.notna() * 10
    confiance += dividende.notna() * 5

    deux_cibles = yahoo.notna() & affaires.notna() & (yahoo > 0) & (affaires > 0)
    desaccord = (abs(yahoo - affaires) / pd.concat([yahoo, affaires], axis=1).mean(axis=1)).where(deux_cibles)
    confiance -= desaccord.fillna(0).clip(0, 1) * 25
    df["Confiance"] = confiance.clip(0, 100)

    # === v7 : Concordance des deux cibles (0..100), NaN si une seule source ===
    df["Concordance"] = (100 * (1 - desaccord.clip(0, 1))).where(deux_cibles)

    # --- Risque : volatilité + proximité du sommet 52s + cible atteinte + faible confiance ---
    risque = pd.Series(30.0, index=df.index)
    risque += volatilite.apply(lambda v: pct_vers_points(v, 20, 90)).fillna(25) * 0.45
    risque += chaleur.apply(lambda v: pct_vers_points(v, 70, 100)).fillna(15) * 0.25
    risque += potentiel.apply(lambda v: 25 if pd.notna(v) and v < 0 else 0)
    risque += df["Confiance"].apply(lambda v: pct_vers_points(60 - v, 0, 60)).fillna(0) * 0.25
    df["Risque"] = risque.clip(0, 100)

    if pour_portefeuille:
        # Portefeuille : signal de VENTE (Vendre / À surveiller / Attendre)
        df["Signal"] = potentiel.apply(signal_portefeuille)
    else:
        # Prospects : signal d'ACHAT (Priorité / À surveiller / ...)
        df["Signal"] = [
            classifier_signal(s, c, r, p)
            for s, c, r, p in zip(df["Score"], df["Confiance"], df["Risque"], potentiel)
        ]

    # === v7 : RANG D'ACHAT composite (Score + Entrée + Concordance + faible Risque),
    # pénalisé par la sur-concentration sectorielle du portefeuille. ===
    score_s = pd.to_numeric(df["Score"], errors="coerce")
    risque_s = pd.to_numeric(df["Risque"], errors="coerce")
    entree_s = pd.to_numeric(df["Entrée"], errors="coerce")
    conc_s = pd.to_numeric(df["Concordance"], errors="coerce")
    rang = (score_s.fillna(0) * 0.45
            + entree_s.fillna(score_s).fillna(0) * 0.20
            + conc_s.fillna(50) * 0.15
            + (100 - risque_s.fillna(50)) * 0.20)
    concentration = pd.Series(0.0, index=df.index)
    if secteurs_portefeuille and "Secteur" in df.columns:
        concentration = df["Secteur"].map(
            lambda s: secteurs_portefeuille.get(str(s), 0) if pd.notna(s) and str(s).strip() != "" else 0
        )
        concentration = pd.to_numeric(concentration, errors="coerce").fillna(0)
        rang = rang - concentration.clip(0, 4) * 5   # -5 par titre déjà détenu dans ce secteur
    df["Concentration Secteur"] = concentration
    df["Achat Rang"] = rang.clip(0, 100)

    # Explication lisible ("Pourquoi") — utile surtout pour les Prospects.
    df["Pourquoi"] = df.apply(_pourquoi_achat, axis=1)

    # === v8 : version AFFICHAGE du signal avec icône. La colonne logique 'Signal'
    # reste intacte (filtres, journal, alertes comparent les libellés exacts).
    if pour_portefeuille:
        icones = {"Vendre": "🔴", "À surveiller": "🟡", "Attendre": "🟢"}
    else:
        icones = {"Priorité": "⭐", "À surveiller": "🔵", "À valider": "🟡",
                  "Risque élevé": "🔴", "Objectif atteint": "⚪", "Secondaire": "▫️",
                  "Données insuffisantes": "⚪"}
    df["Signal Aff"] = df["Signal"].map(lambda s: f"{icones[s]} {s}" if s in icones else s)

    return df

def couleur_var(valeur):
    if pd.isna(valeur): return ''
    if valeur > 0: return 'color: #00cc00;'
    elif valeur < 0: return 'color: #ff4d4d;'
    return ''

def couleur_baisse_fabrique(seuil):
    # === v7 : rouge quand le titre a reculé de plus de `seuil` % depuis son sommet 52s ===
    def couleur_baisse(valeur):
        if pd.isna(valeur): return ''
        if valeur <= -abs(seuil): return 'background-color: rgba(217, 75, 75, .28);'
        if valeur <= -abs(seuil) / 2: return 'background-color: rgba(255, 209, 102, .25);'
        return ''
    return couleur_baisse

def couleur_preg_prospect(valeur):
    # === v8 : dégradé continu sur le potentiel des Prospects ===
    # <= 0 % : rouge ; 0 -> 40 % : interpolation jaune -> vert ; > 40 % : vert plein.
    if pd.isna(valeur):
        return ''
    v = float(valeur)
    if v <= 0:
        return 'background-color: rgba(217, 75, 75, .28);'
    t = min(v, 40.0) / 40.0
    r = int(255 + (0 - 255) * t)
    g = int(209 + (166 - 209) * t)
    b = int(102 + (90 - 102) * t)
    return f'background-color: rgba({r}, {g}, {b}, .25);'

def couleur_alerte_vente(valeur):
    if pd.isna(valeur): return ''
    if valeur <= 5: return 'background-color: rgba(255, 0, 0, 0.3)'
    elif valeur < 15: return 'background-color: rgba(255, 255, 0, 0.3)'
    else: return 'background-color: rgba(0, 255, 0, 0.3)'

def couleur_signal(valeur):
    # v8 : correspondance par INCLUSION (la valeur affichée porte une icône, ex. « ⭐ Priorité »)
    couleurs = {
        "Priorité": "background-color: rgba(0, 166, 90, .22);",
        "À surveiller": "background-color: rgba(106, 169, 255, .20);",
        "À valider": "background-color: rgba(255, 209, 102, .25);",
        "Risque élevé": "background-color: rgba(217, 75, 75, .22);",
        "Objectif atteint": "background-color: rgba(127, 127, 127, .18);",
    }
    v = str(valeur)
    for cle, style in couleurs.items():
        if cle in v:
            return style
    return ""

def couleur_signal_portefeuille(valeur):
    # Vendre = rouge, À surveiller = jaune, Attendre = vert (comme la colonne Pré G %).
    couleurs = {
        "Vendre": "background-color: rgba(217, 75, 75, .28);",
        "À surveiller": "background-color: rgba(255, 209, 102, .28);",
        "Attendre": "background-color: rgba(0, 166, 90, .22);",
    }
    v = str(valeur)
    for cle, style in couleurs.items():
        if cle in v:
            return style
    return ""

def surligner_prospects(row):
    if row.get('Possede') == True: return ['background-color: rgba(255, 215, 0, 0.4)'] * len(row)
    return [''] * len(row)

def griser_pre_aff_perime(row):
    # Grise la Pré Aff périmée ET (v7) le Pré G % aberrant : valeurs non prises en compte.
    styles = [''] * len(row)
    gris = 'color: #9aa0a6; font-style: italic;'
    if row.get('Pré Aff Périmé') and 'Pré Aff Display' in row.index:
        styles[row.index.get_loc('Pré Aff Display')] = gris
    if row.get('Pré G Aberrant') and 'Pré G %' in row.index:
        styles[row.index.get_loc('Pré G %')] = gris
    return styles

def config_largeur_description(df, afficher, px_par_char=8, largeur_min=120, largeur_max=600):
    # Largeur MINIMALE de la colonne Description, calée sur la plus longue description
    # RÉELLEMENT affichée dans CET onglet (optimise l'espace onglet par onglet).
    # px_par_char / largeur_min / largeur_max permettent d'être plus ou moins agressif.
    if not afficher or 'Description' not in df.columns:
        return {}
    longueurs = df['Description'].dropna().astype(str).map(len)
    max_len = int(longueurs.max()) if len(longueurs) > 0 else 0
    if max_len <= 0:
        return {}
    largeur = int(min(max(max_len * px_par_char + 16, largeur_min), largeur_max))
    try:
        return {"Description": st.column_config.TextColumn("Description", width=largeur)}
    except Exception:
        return {"Description": st.column_config.TextColumn("Description", width="large")}

def config_largeur_pourquoi(df, largeur_max=1100):
    # === v7 : largeur de la colonne Pourquoi calée sur le motif le plus long, pour
    # occuper l'écran d'un PC (sur mobile la colonne n'est pas affichée de toute façon).
    if 'Pourquoi' not in df.columns:
        return {}
    longueurs = df['Pourquoi'].dropna().astype(str).map(len)
    max_len = int(longueurs.max()) if len(longueurs) > 0 else 0
    if max_len <= 0:
        return {}
    largeur = int(min(max(max_len * 7 + 16, 220), largeur_max))
    try:
        return {"Pourquoi": st.column_config.TextColumn("Pourquoi", width=largeur)}
    except Exception:
        return {"Pourquoi": st.column_config.TextColumn("Pourquoi", width="large")}

def hauteur_tableau(nb_lignes, max_lignes=18):
    # Hauteur plafonnée à max_lignes visibles : au-delà, le tableau défile à
    # l'interne et sa ligne d'en-tête reste figée (comportement natif de
    # st.dataframe). En deçà, la hauteur épouse le contenu comme avant.
    return (min(nb_lignes, max_lignes) * 35) + 43

def config_colonnes_communes():
    # Configuration d'affichage partagée par tous les tableaux (idée reprise de Codex).
    cfg = {
        "No.": st.column_config.NumberColumn("No.", format="%d"),
        "Qtée": st.column_config.NumberColumn("Qtée", format="%d"),
        "Symbole": st.column_config.LinkColumn("Symbole", display_text=r"https://ca\.finance\.yahoo\.com/quote/(.*)"),
        "Prix $": st.column_config.NumberColumn("Prix", format="$ %.2f"),
        "Achat $": st.column_config.NumberColumn("Achat", format="$ %.2f"),
        "Gain $": st.column_config.NumberColumn("Gain $", format="$ %.2f"),
        "Gain %": st.column_config.NumberColumn("Gain %", format="%.1f %%"),
        "Var %": st.column_config.NumberColumn("Var %", format="%.1f %%"),
        "Pré G %": st.column_config.NumberColumn("Pré G %", format="%.1f %%"),
        "Pré YF Display": st.column_config.NumberColumn("Pré YF", format="$ %.2f"),
        "Pré Aff Display": st.column_config.NumberColumn("Pré Aff", format="$ %.2f"),
        "Tendance": st.column_config.LineChartColumn("Tendance (1m)"),
        "Chaleur 52s": st.column_config.ProgressColumn("♨️ 52 sem.", format="%.0f %%", min_value=0, max_value=100),
        "Div %": st.column_config.NumberColumn("Div %", format="%.2f %%"),
        "Volatilité 1m": st.column_config.NumberColumn("Volat.", format="%.1f %%"),
        "Nb Analystes": st.column_config.NumberColumn("Nb An.", format="%d"),
        "Score": st.column_config.ProgressColumn("Score", format="%.0f", min_value=0, max_value=100),
        "Confiance": st.column_config.ProgressColumn("Conf.", format="%.0f", min_value=0, max_value=100),
        "Risque": st.column_config.ProgressColumn("Risque", format="%.0f", min_value=0, max_value=100),
        "Achat Rang": st.column_config.ProgressColumn("🏆 Rang", format="%.0f", min_value=0, max_value=100),
        "Baisse 52s %": st.column_config.NumberColumn("↘ Sommet", format="%.1f %%"),
        "P/E": st.column_config.NumberColumn("P/E", format="%.1f"),
        "Croiss Rev %": st.column_config.NumberColumn("Cr. Rev", format="%.1f %%"),
        "Marge %": st.column_config.NumberColumn("Marge", format="%.1f %%"),
        "Rang %": st.column_config.NumberColumn("Rang %", format="%.0f %%"),
        "Concordance": st.column_config.ProgressColumn("Concord.", format="%.0f", min_value=0, max_value=100),
        "Entrée": st.column_config.ProgressColumn("Entrée", format="%.0f", min_value=0, max_value=100),
        "Secteur": st.column_config.TextColumn("Secteur", width="small"),
        "Pourquoi": st.column_config.TextColumn("Pourquoi", width="medium"),
        "Signal": st.column_config.TextColumn("Signal", width="small"),
        "Signal Aff": st.column_config.TextColumn("Signal", width="small"),
        "Date Achat": st.column_config.DatetimeColumn("Date Achat", format="YYYY-MM-DD"),
        "MAJ YF": st.column_config.TextColumn("Date YF", width="small"),
        "MAJ Aff": st.column_config.TextColumn("Date Aff", width="small"),
    }
    # === v7 : sur MOBILE, le Rang en simple chiffre très étroit (la barre de
    # progression est trop large pour un petit écran). Sur ordinateur, barre inchangée.
    if globals().get('mode_mobile'):
        try:
            cfg["Achat Rang"] = st.column_config.NumberColumn("🏆", format="%.0f", width=45)
            # Avec le zoom mobile (0.78), il y a la place d'afficher le signal en entier.
            cfg["Signal Aff"] = st.column_config.TextColumn("Signal", width=118)
        except Exception:
            cfg["Achat Rang"] = st.column_config.NumberColumn("🏆", format="%.0f", width="small")
    else:
        # === v8 : sur PC, colonne Signal JUSTE assez large pour le plus long libellé
        # (« ⚪ Objectif atteint ») — on maximise l'espace des autres colonnes.
        try:
            cfg["Signal Aff"] = st.column_config.TextColumn("Signal", width=132)
        except Exception:
            cfg["Signal Aff"] = st.column_config.TextColumn("Signal", width="medium")
    return cfg

# === v7 : ORDONNANCEUR de rafraîchissement PAR GROUPE (mode auto en séance) =======
# Cadences : Portefeuille 10 min, Prospects CAD 25 min, Prospects US 35 min.
# La page se recharge aux 5 min ; à chaque rechargement, AU PLUS UN groupe (le plus
# prioritaire dont l'intervalle est écoulé) est réellement rafraîchi -> jamais deux
# groupes en même temps, toujours >= 5 min entre deux appels Yahoo.
# Les horodatages vivent dans un cache_resource (survivent aux rechargements de page).
# ==================================================================================
INTERVALLES_FETCH = {"P1": 10 * 60, "P2": 25 * 60, "P3": 35 * 60}

@st.cache_resource
def _etat_fetch_auto():
    return {"last": {}, "ver": {"P1": 0, "P2": 0, "P3": 0}}

def jetons_fetch(auto_seance, marche_us=True, marche_ca=True):
    maintenant = time.time()
    if not auto_seance:
        # Mode normal : tout le monde partage une tranche de 5 min (= ancien ttl 300).
        b = int(maintenant // 300)
        return {g: ("t", b) for g in ("P1", "P2", "P3")}
    # === v8 : un groupe n'est rafraîchi que si SA bourse est ouverte ===
    # P2 (Pros CAD) <-> TSX ; P3 (Pros US) <-> bourse US ; P1 (Portefeuille, mixte
    # CAD+US) dès qu'une des deux est ouverte. Un groupe dont la bourse est fermée
    # garde ses valeurs en cache (sa « Dernière synchro » ne bouge pas).
    autorise = {"P1": marche_us or marche_ca, "P2": marche_ca, "P3": marche_us}
    etat = _etat_fetch_auto()
    last, ver = etat["last"], etat["ver"]
    if not last:
        # Premier chargement : les 3 groupes partent maintenant (cache vide de toute
        # façon) ; leurs prochaines échéances (10/25/35 min) sont naturellement décalées.
        for g in ("P1", "P2", "P3"):
            last[g] = maintenant
    else:
        for g in ("P1", "P2", "P3"):   # priorité P1 > P2 > P3 ; UN seul par rechargement
            if not autorise.get(g, True):
                continue
            if maintenant - last.get(g, 0) >= INTERVALLES_FETCH[g]:
                ver[g] += 1
                last[g] = maintenant
                break
    return {g: ("v", ver[g]) for g in ("P1", "P2", "P3")}

try:
    with st.spinner("Connexion à Google Sheets..."):
        df_base_portefeuille = charger_donnees_base('Portefeuille BNC')
        df_base_prospects = charger_donnees_base('Prospects')

    if 'No.' in df_base_portefeuille.columns:
        df_portefeuille_actif = df_base_portefeuille[df_base_portefeuille['No.'] != 0].reset_index(drop=True)
    else:
        df_portefeuille_actif = df_base_portefeuille.copy()

    # === v6 : GROUPES DE PRIORITÉ pour la récupération Yahoo ===
    #   P1 = Portefeuille (+ indices) ; P2 = Prospects CAD ; P3 = Prospects US.
    # On ajoute pour chaque action CAD son équivalent US (règle de trois sur Pré YF).
    SUFFIXES_CAD = ('.TO', '.V', '.NE', '.CN')

    def candidat_us(s):
        for suff in SUFFIXES_CAD:
            if s.endswith(suff):
                return s[:-len(suff)]
        return None

    def symboles_de(df_):
        if 'Symbole' not in df_.columns:
            return []
        return [str(s).strip() for s in df_['Symbole'].dropna() if str(s).strip() not in ("", "0")]

    def avec_us(liste):
        out = list(liste)
        for s in liste:
            c = candidat_us(s)
            if c:
                out.append(c)
        return out

    syms_port = symboles_de(df_portefeuille_actif)
    syms_pros = symboles_de(df_base_prospects)
    grp1 = avec_us(syms_port) + ["^GSPC", "^IXIC", "^GSPTSE"]
    grp2 = avec_us([s for s in syms_pros if s.endswith(SUFFIXES_CAD)])
    grp3 = avec_us([s for s in syms_pros if not s.endswith(SUFFIXES_CAD)])

    # Dédupliquer en RESPECTANT la priorité (un symbole déjà en P1 n'est pas refait en P2/P3)
    vus = set()
    def dedup(g):
        r = []
        for s in g:
            if s and s not in vus:
                vus.add(s)
                r.append(s)
        return tuple(r)
    g1, g2, g3 = dedup(grp1), dedup(grp2), dedup(grp3)
    symboles_possedes = tuple(set(df_portefeuille_actif['Symbole'].dropna().astype(str).str.strip()))

    # === PHASE 1 : PORTEFEUILLE (priorité 1) — récupéré et AFFICHÉ en premier ===
    jetons = jetons_fetch(rafraichir_auto and (ouvert_us or ouvert_ca), ouvert_us, ouvert_ca)
    with st.spinner("Chargement du Portefeuille..."):
        yahoo_p1 = telecharger_yahoo((g1,), retry_premier=True, jeton=jetons["P1"])

    # === v7 : cache YF (secours quand Yahoo bloque .info -> scores stables) ===
    if 'cache_yf' not in st.session_state:
        st.session_state['cache_yf'] = charger_cache_yf()
    cache_yf = st.session_state['cache_yf']

    df_live = construire_donnees(df_portefeuille_actif, yahoo_p1, est_portefeuille=True, plafond_scaling=plafond_preg)
    df_live = appliquer_cache_yf(df_live, cache_yf)
    df_live = calculer_potentiel_gain(df_live, source_gain, est_portefeuille=True, min_analystes=min_analystes, mois_max_aff=mois_max_aff, plafond_preg=plafond_preg)
    for col in ["Pré G %", "Gain %", "Var %"]:
        if col in df_live.columns: df_live[col] = pd.to_numeric(df_live[col], errors='coerce') * 100
    df_live = calculer_score_decision(df_live, pour_portefeuille=True)  # === v5 : signal de vente ===

    # === v7 : concentration sectorielle du PORTEFEUILLE (nb de titres détenus par secteur),
    # utilisée pour pénaliser la sur-concentration dans le Rang d'achat des Prospects. ===
    secteurs_portefeuille = {}
    if 'Secteur' in df_live.columns:
        _sect = df_live['Secteur'].dropna().astype(str)
        _sect = _sect[_sect.str.strip() != ""]
        secteurs_portefeuille = _sect.value_counts().to_dict()

    # Sauvegarde auto du Portefeuille (une fois par rafraîchissement, via signature)
    sig_port = signature_donnees(("PORT",) + g1)
    if st.session_state.get('sig_save_port') != sig_port:
        ok_p, msg_p = sauvegarder_donnees_dans_sheets(df_live, 'Portefeuille BNC')
        st.session_state['sig_save_port'] = sig_port
        if ok_p: st.toast("💾 Portefeuille synchronisé.", icon="✅")
        else: st.warning(f"Sauvegarde Portefeuille : {msg_p}")

    if afficher_bandeau:
        indices_marches = {"S&P 500": "^GSPC", "NASDAQ": "^IXIC", "TSX": "^GSPTSE"}
        cols_m = st.columns(3)
        for idx, (nom_m, sym_m) in enumerate(indices_marches.items()):
            m_data = yahoo_p1.get(sym_m, {}).get('hist', pd.DataFrame())
            if not m_data.empty and len(m_data) >= 2:
                m_actuel = m_data['Close'].iloc[-1]
                m_veille = m_data['Close'].iloc[-2]
                m_var = (m_actuel - m_veille) / m_veille * 100
                m_signe = "+" if m_var > 0 else ""
                cols_m[idx].markdown(f"<div class='market-block'><b>{nom_m}</b> : {m_actuel:,.2f} (<span style='color:{'#00cc00' if m_var > 0 else '#ff4d4d'}'>{m_signe}{m_var:.2f}%</span>)</div>", unsafe_allow_html=True)
            else:
                cols_m[idx].markdown(f"<div class='market-block'><b>{nom_m}</b> : Indisponible</div>", unsafe_allow_html=True)
        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

    # Emplacement réservé en HAUT (rempli quand les prospects sont chargés, phase 2)
    ph_alertes = st.empty()

    # --- ARCHITECTURE DES COLONNES (Unifiée pour tous les onglets) ---
    colonnes_base_port = []
    if afficher_no: colonnes_base_port.append("No.")
    colonnes_base_port.append("Symbole")
    if afficher_desc: colonnes_base_port.append("Description")
    if afficher_dev: colonnes_base_port.append("Dev.")
    if afficher_compte: colonnes_base_port.append("Compte")

    # NOUVEL ORDRE DES COLONNES PRINCIPALES
    colonnes_base_port.append("Prix $")
    colonnes_base_port.append("Gain $")
    colonnes_base_port.append("Gain %")

    # Portefeuille : seul le Signal (Vendre / À surveiller / Attendre) est affiché.
    # Score / Confiance / Risque restent réservés aux onglets Pros.
    if afficher_signal: colonnes_base_port.append("Signal Aff")

    if afficher_var: colonnes_base_port.append("Var %")
    if afficher_baisse: colonnes_base_port.append("Baisse 52s %")     # === v7 ===
    if afficher_tendance: colonnes_base_port.append("Tendance")
    if afficher_chaleur: colonnes_base_port.append("Chaleur 52s")
    if afficher_div: colonnes_base_port.append("Div %")
    if afficher_volatilite: colonnes_base_port.append("Volatilité 1m")
    if afficher_analystes: colonnes_base_port.append("Nb Analystes")  # === V4 ===
    if afficher_fondamentaux: colonnes_base_port.extend(["P/E", "Croiss Rev %", "Marge %"])  # === v7 ===

    colonnes_base_port.extend(["Pré YF Display", "Pré Aff Display", "Pré G %", "Achat $", "Qtée", "Date Achat", "MAJ YF", "MAJ Aff"])

    # On utilise la même logique d'affichage de base pour les prospects
    colonnes_base_pros = []
    colonnes_base_pros.append("Symbole")
    if afficher_desc: colonnes_base_pros.append("Description")
    if afficher_signal: colonnes_base_pros.append("Signal Aff")
    if afficher_score: colonnes_base_pros.append("Score")
    if afficher_confiance: colonnes_base_pros.append("Confiance")
    if afficher_risque: colonnes_base_pros.append("Risque")
    if afficher_rang: colonnes_base_pros.append("Achat Rang")          # === v7 ===
    if afficher_percentile: colonnes_base_pros.append("Rang %")        # === v7 ===
    if afficher_concordance: colonnes_base_pros.append("Concordance")  # === v7 ===
    if afficher_entree: colonnes_base_pros.append("Entrée")            # === v7 ===
    if afficher_secteur: colonnes_base_pros.append("Secteur")          # === v7 ===
    if afficher_dev: colonnes_base_pros.append("Dev.")
    if afficher_compte: colonnes_base_pros.append("Compte")
    colonnes_base_pros.append("Prix $")

    if afficher_var: colonnes_base_pros.append("Var %")
    colonnes_base_pros.append("Pré G %")
    if afficher_tendance: colonnes_base_pros.append("Tendance")
    if afficher_chaleur: colonnes_base_pros.append("Chaleur 52s")
    if afficher_div: colonnes_base_pros.append("Div %")
    if afficher_volatilite: colonnes_base_pros.append("Volatilité 1m")
    if afficher_analystes: colonnes_base_pros.append("Nb Analystes")  # === V4 ===
    if afficher_fondamentaux: colonnes_base_pros.extend(["P/E", "Croiss Rev %", "Marge %"])  # === v7 ===
    colonnes_base_pros.extend(["Pré YF Display", "MAJ YF", "Pré Aff Display", "MAJ Aff"])
    if afficher_pourquoi: colonnes_base_pros.append("Pourquoi")        # === v7 ===

    # === v7 : MODE MOBILE — on ne garde que les colonnes ESSENTIELLES (l'ordre des
    # listes est préservé). Sur ordinateur, tous les détails restent affichés.
    if mode_mobile:
        ESSENTIEL_PORT = {"Symbole", "Prix $", "Gain %", "Signal Aff", "Pré G %"}
        ESSENTIEL_PROS = {"Symbole", "Signal Aff", "Achat Rang", "Prix $", "Pré G %", "Pré Aff Display"}
        colonnes_base_port = [c for c in colonnes_base_port if c in ESSENTIEL_PORT]
        colonnes_base_pros = [c for c in colonnes_base_pros if c in ESSENTIEL_PROS]
        # (l'indicateur 📱 est intégré au début de la ligne des heures de synchro)

    tab1, tab_dec, tab2, tab3, tab4 = st.tabs(["💰 Portefeuille", "🧭 Décision", "🎯 Pros CAD", "🎯 Pros US", "📘 Méthode"])

    # --- ONGLET 1 : PORTEFEUILLE ---
    with tab1:
        if 'Prix $' in df_live.columns and 'Qtée' in df_live.columns:
            valeurs_brutes = pd.to_numeric(df_live['Prix $'], errors='coerce') * pd.to_numeric(df_live['Qtée'], errors='coerce').fillna(0)
            gains_bruts = pd.to_numeric(df_live['Gain $'], errors='coerce').fillna(0)
            gains_jour_bruts = pd.to_numeric(df_live['Gain Jour $'], errors='coerce').fillna(0)

            if activer_taux_change:
                valeurs_converties = np.where(df_live['Devise'] == 'USD', valeurs_brutes * taux_usdcad, valeurs_brutes)
                gains_convertis = np.where(df_live['Devise'] == 'USD', gains_bruts * taux_usdcad, gains_bruts)
                gains_jour_convertis = np.where(df_live['Devise'] == 'USD', gains_jour_bruts * taux_usdcad, gains_jour_bruts)
                titre_gain = "Gain net ($ CA)"
                titre_gain_j = "Gain Jour ($ CA)"
                titre_valeur = "Valeur Nette ($ CA)"
                symbole_devise = "$ CA"
                texte_taux = f"<p style='margin: 0px; font-size: 11px; color: gray;'>1 USD = {taux_usdcad:.3f} CAD</p>"
            else:
                valeurs_converties = valeurs_brutes
                gains_convertis = gains_bruts
                gains_jour_convertis = gains_jour_bruts
                titre_gain = "Gain total"
                titre_gain_j = "Gain du jour"
                symbole_devise = "$"
                titre_valeur = "Valeur totale"
                texte_taux = ""

            valeur_totale_nette = valeurs_converties.sum()
            gain_total_net = gains_convertis.sum()
            gain_jour_total_net = gains_jour_convertis.sum()
        else:
            valeur_totale_nette = 0
            gain_total_net = 0
            gain_jour_total_net = 0
            titre_gain = "Gain total"
            titre_gain_j = "Gain du jour"
            titre_valeur = "Valeur totale"
            symbole_devise = "$"
            texte_taux = ""

        gain_formate = f"{gain_total_net:,.2f} {symbole_devise}".replace(',', ' ')
        gain_j_formate = f"{gain_jour_total_net:,.2f} {symbole_devise}".replace(',', ' ')
        valeur_formate = f"{valeur_totale_nette:,.2f} {symbole_devise}".replace(',', ' ')

        cols_s = st.columns(3) if afficher_gain_jour else st.columns(2)

        # === v8 : cartes colorées — le Gain total est vert/rouge avec flèche, comme le Gain du jour ===
        coul_gain = '#00cc00' if gain_total_net >= 0 else '#ff4d4d'
        fleche_gain = '▲ ' if gain_total_net > 0 else ('▼ ' if gain_total_net < 0 else '')
        coul_jour = '#00cc00' if gain_jour_total_net >= 0 else '#ff4d4d'
        fleche_jour = '▲ ' if gain_jour_total_net > 0 else ('▼ ' if gain_jour_total_net < 0 else '')

        with cols_s[0]:
            st.markdown(f"<div class='stats-block' style='text-align: left;'><p style='margin: 0px; font-size: 13px; color: gray;'>{titre_gain}</p><p style='margin: 0px; font-size: 16px; font-weight: bold; color: {coul_gain};'>{fleche_gain}{gain_formate}</p></div>", unsafe_allow_html=True)

        if afficher_gain_jour:
            with cols_s[1]:
                st.markdown(f"<div class='stats-block' style='text-align: center;'><p style='margin: 0px; font-size: 13px; color: gray;'>{titre_gain_j}</p><p style='margin: 0px; font-size: 16px; font-weight: bold; color: {coul_jour};'>{fleche_jour}{gain_j_formate}</p></div>", unsafe_allow_html=True)

        idx_val = 2 if afficher_gain_jour else 1

        with cols_s[idx_val]:
            st.markdown(f"<div class='stats-block' style='text-align: center;'><p style='margin: 0px; font-size: 13px; color: gray;'>{titre_valeur}</p><p style='margin: 0px; font-size: 16px; font-weight: bold;'>{valeur_formate}</p>{texte_taux}</div>", unsafe_allow_html=True)

        # Pré G % se trie en ordre CROISSANT (titres proches/au-dessus de l'objectif = à surveiller en haut).
        if colonne_tri == "Pré G %":
            df_live = df_live.sort_values(by="Pré G %", ascending=True, na_position="last")
        else:
            df_live = df_live.sort_values(by=colonne_tri, ascending=False, na_position="last")

        colonnes_a_afficher = [c for c in colonnes_base_port if c in df_live.columns]
        config_description = config_largeur_description(df_live, afficher_desc)

        styled_port = df_live.style
        if 'Pré G %' in df_live.columns:
            styled_port = styled_port.map(couleur_alerte_vente, subset=['Pré G %'])
        if afficher_var and 'Var %' in df_live.columns:
            styled_port = styled_port.map(couleur_var, subset=['Var %'])
        if afficher_baisse and 'Baisse 52s %' in df_live.columns:
            styled_port = styled_port.map(couleur_baisse_fabrique(seuil_baisse), subset=['Baisse 52s %'])
        if 'Signal Aff' in df_live.columns:
            styled_port = styled_port.map(couleur_signal_portefeuille, subset=['Signal Aff'])
        if 'Pré Aff Périmé' in df_live.columns:
            styled_port = styled_port.apply(griser_pre_aff_perime, axis=1)

        st.dataframe(
            styled_port,
            use_container_width=False, hide_index=True, height=hauteur_tableau(len(df_live)),
            column_order=colonnes_a_afficher,
            column_config={**config_description, **config_colonnes_communes()}
        )

    # === PHASE 2 : PROSPECTS (priorités 2 et 3) — chargés APRÈS l'affichage du Portefeuille ===
    # Appels SÉPARÉS par groupe : chacun a son propre cache/jeton (cadences 25/35 min
    # en mode auto ; en mode normal les jetons partagent la même tranche de 5 min).
    with st.spinner("Chargement des Prospects (CAD puis US)..."):
        yahoo_p2 = telecharger_yahoo((g2,), jeton=jetons["P2"])
        yahoo_p3 = telecharger_yahoo((g3,), jeton=jetons["P3"])
    yahoo_p23 = {**{k: v for k, v in yahoo_p2.items() if not k.startswith('__')},
                 **{k: v for k, v in yahoo_p3.items() if not k.startswith('__')}}
    yahoo_data = {**yahoo_p1, **yahoo_p23}   # équivalents US de P1 partagés (règle de trois)

    df_live_prospects = construire_donnees(df_base_prospects, yahoo_data, est_portefeuille=False, symboles_portefeuille=symboles_possedes, plafond_scaling=plafond_preg)
    df_live_prospects = appliquer_cache_yf(df_live_prospects, cache_yf)
    df_live_prospects = calculer_potentiel_gain(df_live_prospects, source_gain, est_portefeuille=False, min_analystes=min_analystes, mois_max_aff=mois_max_aff, plafond_preg=plafond_preg)
    for col in ["Pré G %", "Var %"]:
        if col in df_live_prospects.columns: df_live_prospects[col] = pd.to_numeric(df_live_prospects[col], errors='coerce') * 100
    df_live_prospects = calculer_score_decision(df_live_prospects, secteurs_portefeuille=secteurs_portefeuille)  # === v5/v7 ===

    # Sauvegarde auto des Prospects
    sig_pros = signature_donnees(("PROS",) + g2 + g3)
    if st.session_state.get('sig_save_pros') != sig_pros:
        ok_r, msg_r = sauvegarder_donnees_dans_sheets(df_live_prospects, 'Prospects')
        st.session_state['sig_save_pros'] = sig_pros
        if ok_r: st.toast("💾 Prospects synchronisés.", icon="✅")
        else: st.warning(f"Sauvegarde Prospects : {msg_r}")

    # Statuts par groupe (dans chaque appel mono-groupe, niveaux_ok == [1] = réussi).
    # v8 : plus de message « Yahoo a limité les requêtes » — la ligne « Dernière synchro
    # réussie » suffit (une heure qui ne bouge pas = groupe non rafraîchi).
    stat2 = yahoo_p2.get('__statut__', {})
    stat3 = yahoo_p3.get('__statut__', {})

    # === v7 : mémorise et affiche la dernière synchro RÉUSSIE par groupe ===
    synchros = _dernieres_synchros()
    if 1 in yahoo_p1.get('__statut__', {}).get('niveaux_ok', []):
        synchros['Portefeuille'] = yahoo_p1.get('__horodatage__', '')
    if 1 in stat2.get('niveaux_ok', []):
        synchros['Pros CAD'] = yahoo_p2.get('__horodatage__', '')
    if 1 in stat3.get('niveaux_ok', []):
        synchros['Pros US'] = yahoo_p3.get('__horodatage__', '')
    if mobile_ui:
        # v8 : version compacte — 📱 (mode mobile) ouvre la ligne des heures de synchro
        ph_sync.markdown(
            f"<div style='font-size: 12px; color: gray; margin-top: -4px; margin-bottom: 6px;'>"
            f"📱 Portef. <b>{_fmt_sync(synchros.get('Portefeuille'))}</b>"
            f" · CAD <b>{_fmt_sync(synchros.get('Pros CAD'))}</b>"
            f" · US <b>{_fmt_sync(synchros.get('Pros US'))}</b></div>",
            unsafe_allow_html=True
        )
    else:
        ph_sync.markdown(
            f"<div style='font-size: 12px; color: gray; margin-top: -4px; margin-bottom: 6px;'>"
            f"🕒 Dernière synchro réussie — Portefeuille : <b>{_fmt_sync(synchros.get('Portefeuille'))}</b>"
            f" &nbsp;·&nbsp; Pros CAD : <b>{_fmt_sync(synchros.get('Pros CAD'))}</b>"
            f" &nbsp;·&nbsp; Pros US : <b>{_fmt_sync(synchros.get('Pros US'))}</b></div>",
            unsafe_allow_html=True
        )

    # Alertes (Portefeuille + Prospects) -> emplacement réservé en haut
    if afficher_alertes:
        with ph_alertes.container():
            alertes_generees = []
            if not df_live.empty:
                for _, row in df_live.iterrows():
                    sym = row.get('Symbole Brut', 'Action')
                    if pd.notna(row.get('Pré G %')) and row['Pré G %'] <= 0:
                        alertes_generees.append(f"🎯 **{sym}** a atteint son objectif de prix !")
                    if pd.notna(row.get('Var %')):
                        if row['Var %'] >= 5.0: alertes_generees.append(f"🚀 **{sym}** s'envole (+{row['Var %']:.1f}%)")
                        elif row['Var %'] <= -5.0: alertes_generees.append(f"🔻 **{sym}** chute ({row['Var %']:.1f}%)")
            if not df_live_prospects.empty:
                for _, row in df_live_prospects.iterrows():
                    sym = row.get('Symbole Brut', 'Action')
                    if row.get('Signal') == "Priorité":
                        alertes_generees.append(f"⭐ **{sym}** (Prospect) ressort en PRIORITÉ.")
                    if pd.notna(row.get('Chaleur 52s')) and row['Chaleur 52s'] <= 5.0:
                        alertes_generees.append(f"🔥 **{sym}** (Prospect) est au plus bas sur 1 an !")
            if alertes_generees:
                html_alertes = "<div class='alert-box'><strong>🚨 Alertes Actives :</strong><br>"
                for alerte in alertes_generees: html_alertes += f"<p class='alert-item'>{alerte}</p>"
                st.markdown(html_alertes + "</div>", unsafe_allow_html=True)

    # === v7 : JOURNAL (1 fois/jour) puis ONGLET DÉCISION =========================
    journal_rows = []
    if journaliser:
        sig_journal = signature_donnees(("JOURNAL",) + g1 + g2 + g3)
        if st.session_state.get('sig_journal') != sig_journal:
            # Clôtures des indices pour la courbe « Portefeuille vs marché »
            indices_cloture = {}
            for sym_i in ("^GSPTSE", "^GSPC"):
                h_i = yahoo_p1.get(sym_i, {}).get('hist', pd.DataFrame())
                if not h_i.empty and 'Close' in h_i.columns:
                    c_i = h_i['Close'].dropna()
                    if len(c_i):
                        indices_cloture[sym_i] = float(c_i.iloc[-1])
            journal_rows, ecrit = journaliser_signaux(df_live, df_live_prospects,
                                                      valeur_totale_nette, indices_cloture)
            st.session_state['sig_journal'] = sig_journal
            st.session_state['journal_rows'] = journal_rows
            if ecrit:
                st.toast("📓 Signaux du jour journalisés.", icon="✅")
        else:
            journal_rows = st.session_state.get('journal_rows', [])

    # === v7 : sauvegarde du cache YF (1 fois par rafraîchissement) ===
    sig_cache = signature_donnees(("CACHE",) + g1 + g2 + g3)
    if st.session_state.get('sig_cache_yf') != sig_cache:
        sauvegarder_cache_yf(df_live, df_live_prospects)
        st.session_state['sig_cache_yf'] = sig_cache

    with tab_dec:
        # --- Top 5 par devise, en TABLEAUX ; CAD séparé détenu / non détenu ---
        if not df_live_prospects.empty and "Achat Rang" in df_live_prospects.columns:
            cols_top = [c for c in ["Symbole", "Description", "Achat Rang", "Signal Aff", "Prix $",
                                    "Pré G %", "Pourquoi"] if c in df_live_prospects.columns]
            possede = df_live_prospects.get("Possede")
            if possede is None:
                possede = pd.Series(False, index=df_live_prospects.index)
            possede = possede.fillna(False).astype(bool)

            # Colonne Pourquoi élargie pour montrer TOUT le texte (helper partagé).
            config_dec = {**config_colonnes_communes(), **config_largeur_pourquoi(df_live_prospects)}

            sections = (
                ("🏆 Top 5 achats CAD 🍁 — non détenus", "CAD", False),
                ("💼 Top 5 CAD 🍁 — déjà détenus (renforcer ?)", "CAD", True),
                ("🏆 Top 5 achats US 💵", "USD", None),
            )
            for titre_top, devise_top, filtre_possede in sections:
                st.markdown(f"#### {titre_top}")
                sel = df_live_prospects["Devise"] == devise_top
                if filtre_possede is not None:
                    sel = sel & (possede == filtre_possede)
                top5 = (df_live_prospects[sel]
                        .drop_duplicates(subset="Symbole Brut", keep="first")
                        .sort_values("Achat Rang", ascending=False, na_position="last").head(5))
                if top5.empty:
                    st.info("Aucun titre dans cette catégorie.")
                    continue
                # Style sur le DataFrame COMPLET (surligner_prospects lit 'Possede'),
                # affichage restreint aux colonnes utiles via column_order.
                styled_top = top5.style.apply(surligner_prospects, axis=1)
                if 'Pré G %' in top5.columns:
                    styled_top = styled_top.map(couleur_preg_prospect, subset=['Pré G %'])
                if 'Signal Aff' in top5.columns:
                    styled_top = styled_top.map(couleur_signal, subset=['Signal Aff'])
                st.dataframe(
                    styled_top,
                    use_container_width=False, hide_index=True,
                    column_order=cols_top,
                    column_config=config_dec
                )
            st.caption("🟡 surligné = déjà détenu. Détails et filtres dans les onglets Pros.")
        else:
            st.info("Prospects non disponibles.")

        # --- Titres à VENDRE (portefeuille) ---
        st.markdown("#### 🔴 À vendre / objectif atteint")
        ventes = df_live[df_live.get("Signal", pd.Series(dtype=object)) == "Vendre"] if not df_live.empty else pd.DataFrame()
        if not ventes.empty:
            cols_v = [c for c in ["Symbole", "Description", "Prix $", "Gain %", "Pré G %", "Baisse 52s %"] if c in ventes.columns]
            st.dataframe(ventes[cols_v], use_container_width=False, hide_index=True,
                         column_config=config_colonnes_communes())
        else:
            st.success("Aucun titre du portefeuille en signal Vendre. ✅")

        # --- Protection des gains : fortes baisses depuis le sommet 52 sem ---
        if 'Baisse 52s %' in df_live.columns and not df_live.empty:
            baisses = df_live[pd.to_numeric(df_live['Baisse 52s %'], errors='coerce') <= -abs(seuil_baisse)]
            if not baisses.empty:
                st.markdown(f"#### ↘ Repli de plus de {seuil_baisse} % depuis le sommet 52 sem")
                cols_b = [c for c in ["Symbole", "Description", "Prix $", "Gain %", "Baisse 52s %", "Signal Aff"] if c in baisses.columns]
                st.dataframe(
                    baisses.sort_values('Baisse 52s %')[cols_b],
                    use_container_width=False, hide_index=True,
                    column_config=config_colonnes_communes()
                )

        # --- Changements vs dernière séance + performance des signaux (via Journal) ---
        if journal_rows and len(journal_rows) > 1:
            jdf = pd.DataFrame(journal_rows[1:], columns=journal_rows[0])
            jdf["Prix"] = pd.to_numeric(jdf.get("Prix"), errors="coerce")
            dates = sorted(jdf["Date"].dropna().unique())
            aujourdhui = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d")

            # Changements : Priorité d'aujourd'hui vs dernière date PRÉCÉDENTE
            dates_prec = [d for d in dates if d < aujourdhui]
            if dates_prec:
                prec = dates_prec[-1]
                set_avant = set(jdf[(jdf["Date"] == prec) & (jdf["Type"] == "Achat")]["Symbole"])
                set_now = set(jdf[(jdf["Date"] == aujourdhui) & (jdf["Type"] == "Achat")]["Symbole"])
                nouveaux = sorted(set_now - set_avant)
                sortis = sorted(set_avant - set_now)
                if nouveaux or sortis:
                    st.markdown(f"#### 🔄 Changements depuis le {prec}")
                    if nouveaux:
                        st.markdown("**Nouvelles Priorités :** " + ", ".join(f"`{s}`" for s in nouveaux))
                    if sortis:
                        st.markdown("**Sorties de Priorité :** " + ", ".join(f"`{s}`" for s in sortis))

            # Performance des signaux Achat passés (1re apparition -> prix actuel)
            achats = jdf[jdf["Type"] == "Achat"].dropna(subset=["Prix"])
            if not achats.empty and not df_live_prospects.empty:
                premiers = achats.sort_values("Date").groupby("Symbole").first().reset_index()
                premiers = premiers[premiers["Date"] < aujourdhui]   # au moins 1 jour de recul
                if not premiers.empty:
                    # Index UNIQUE requis par .map : on écarte les lignes sans symbole
                    # (doublons de chaîne vide) et les éventuels symboles répétés.
                    base_prix = df_live_prospects[
                        df_live_prospects["Symbole Brut"].astype(str).str.strip() != ""
                    ].drop_duplicates(subset="Symbole Brut", keep="first")
                    prix_actuels = pd.to_numeric(
                        base_prix.set_index("Symbole Brut")["Prix $"], errors="coerce")
                    premiers["Prix actuel"] = premiers["Symbole"].map(prix_actuels)
                    premiers["Perf %"] = (premiers["Prix actuel"] - premiers["Prix"]) / premiers["Prix"] * 100
                    perf = premiers.dropna(subset=["Perf %"])
                    if not perf.empty:
                        st.markdown("#### 📓 Performance des signaux « Priorité » passés")
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Moyenne", f"{perf['Perf %'].mean():+.1f} %")
                        meilleur = perf.loc[perf['Perf %'].idxmax()]
                        pire = perf.loc[perf['Perf %'].idxmin()]
                        c2.metric(f"Meilleur ({meilleur['Symbole']})", f"{meilleur['Perf %']:+.1f} %")
                        c3.metric(f"Pire ({pire['Symbole']})", f"{pire['Perf %']:+.1f} %")
                        st.caption(f"{len(perf)} signaux suivis depuis leur 1re apparition au Journal.")
            # --- Évolution : valeur du portefeuille vs indices (base 100) ---
            valeurs_j = jdf[jdf["Type"] == "Valeur"].copy()
            if not valeurs_j.empty:
                valeurs_j["Prix"] = pd.to_numeric(valeurs_j["Prix"], errors="coerce")
                pivot = valeurs_j.pivot_table(index="Date", columns="Symbole",
                                              values="Prix", aggfunc="last").sort_index()
                if len(pivot) >= 2:
                    norm = pivot.div(pivot.iloc[0]) * 100   # base 100 au 1er jour
                    norm = norm.rename(columns={"PORTEFEUILLE": "Portefeuille",
                                                "^GSPTSE": "TSX", "^GSPC": "S&P 500"})
                    st.markdown("#### 📈 Portefeuille vs marché (base 100)")
                    st.line_chart(norm)
                    dern = norm.iloc[-1].dropna()
                    if len(dern):
                        cparts = st.columns(len(dern))
                        for i_c, (nom_c, val_c) in enumerate(dern.items()):
                            cparts[i_c].metric(str(nom_c), f"{val_c - 100:+.1f} %")
                    st.caption(f"Depuis le {pivot.index[0]} ({len(pivot)} jour(s) enregistré(s)).")
                else:
                    st.caption("📈 La courbe Portefeuille vs marché apparaîtra dès le 2e jour de données.")
        elif journaliser:
            st.caption("📓 Le Journal se remplit à chaque jour d'utilisation — les changements et "
                       "la performance des signaux apparaîtront ici dès la 2e séance.")

    # --- ONGLET 2 : PROSPECTS CAD ---
    with tab2:
        col_min, col_max, col_sig, _ = st.columns([1, 1, 4, 2])
        min_score_cad = col_min.number_input("Score min", min_value=0, max_value=100, value=55, step=5, key="cad_min_score")
        max_risque_cad = col_max.number_input("Risque max", min_value=0, max_value=100, value=85, step=5, key="cad_max_risk")
        filtre_signal_cad = col_sig.multiselect(
            "Signaux", SIGNAUX,
            default=["Priorité"], key="cad_signal_filter"
        )
        voir_aff_cad = st.checkbox(
            "Ajouter les titres ayant une prévision Les Affaires (Pré G % ≥ 5 %)",
            value=True, key="cad_voir_aff"
        )

        df_prospects_cad = df_live_prospects[df_live_prospects['Devise'] == 'CAD'].copy()
        if "Score" in df_prospects_cad.columns:
            masque_cad = (
                df_prospects_cad["Score"].fillna(0).ge(min_score_cad)
                & df_prospects_cad["Risque"].fillna(100).le(max_risque_cad)
                & df_prospects_cad["Signal"].isin(filtre_signal_cad)
            )
            if voir_aff_cad:
                aff_cad = pd.to_numeric(df_prospects_cad.get("Pré Aff Display"), errors="coerce")
                perime_cad = df_prospects_cad.get("Pré Aff Périmé")
                if perime_cad is None:
                    perime_cad = pd.Series(False, index=df_prospects_cad.index)
                preg_cad = pd.to_numeric(df_prospects_cad.get("Pré G %"), errors="coerce")
                # prévisions NON périmées seulement, et potentiel d'au moins 5 %
                masque_cad = masque_cad | (aff_cad.notna() & (aff_cad != 0)
                                           & ~perime_cad.fillna(False).astype(bool)
                                           & (preg_cad >= 5))
            df_prospects_cad = df_prospects_cad[masque_cad]
            if trier_par_rang and "Achat Rang" in df_prospects_cad.columns:
                df_prospects_cad = df_prospects_cad.sort_values(by="Achat Rang", ascending=False, na_position="last")
            else:
                df_prospects_cad = df_prospects_cad.sort_values(by=["Score", "Confiance"], ascending=[False, False], na_position="last")
            if "Achat Rang" in df_prospects_cad.columns:   # === v7 : percentile dans la liste affichée ===
                df_prospects_cad["Rang %"] = df_prospects_cad["Achat Rang"].rank(pct=True) * 100

        colonnes_a_afficher_pros = [c for c in colonnes_base_pros if c in df_prospects_cad.columns]
        config_description = config_largeur_description(df_prospects_cad, afficher_desc, px_par_char=6, largeur_min=80, largeur_max=320)

        styled_cad = df_prospects_cad.style.apply(surligner_prospects, axis=1)
        if afficher_var and 'Var %' in df_prospects_cad.columns:
            styled_cad = styled_cad.map(couleur_var, subset=['Var %'])
        if 'Pré G %' in df_prospects_cad.columns:
            styled_cad = styled_cad.map(couleur_preg_prospect, subset=['Pré G %'])   # === v8 : dégradé ===
        if 'Signal Aff' in df_prospects_cad.columns:
            styled_cad = styled_cad.map(couleur_signal, subset=['Signal Aff'])
        if 'Pré Aff Périmé' in df_prospects_cad.columns:
            styled_cad = styled_cad.apply(griser_pre_aff_perime, axis=1)

        st.dataframe(
            styled_cad,
            use_container_width=False, hide_index=True, height=hauteur_tableau(len(df_prospects_cad)),
            column_order=colonnes_a_afficher_pros,
            column_config={**config_description, **config_colonnes_communes(),
                           **config_largeur_pourquoi(df_prospects_cad)}
        )

    # --- ONGLET 3 : PROSPECTS US ---
    with tab3:
        col_min_us, col_max_us, col_sig_us, _ = st.columns([1, 1, 4, 2])
        min_score_us = col_min_us.number_input("Score min", min_value=0, max_value=100, value=55, step=5, key="usd_min_score")
        max_risque_us = col_max_us.number_input("Risque max", min_value=0, max_value=100, value=85, step=5, key="usd_max_risk")
        filtre_signal_us = col_sig_us.multiselect(
            "Signaux", SIGNAUX,
            default=["Priorité"], key="usd_signal_filter"
        )
        voir_aff_us = st.checkbox(
            "Ajouter les titres ayant une prévision Les Affaires (Pré G % ≥ 5 %)",
            value=True, key="usd_voir_aff"
        )

        df_prospects_usd = df_live_prospects[df_live_prospects['Devise'] == 'USD'].copy()
        if "Score" in df_prospects_usd.columns:
            masque_us = (
                df_prospects_usd["Score"].fillna(0).ge(min_score_us)
                & df_prospects_usd["Risque"].fillna(100).le(max_risque_us)
                & df_prospects_usd["Signal"].isin(filtre_signal_us)
            )
            if voir_aff_us:
                aff_us = pd.to_numeric(df_prospects_usd.get("Pré Aff Display"), errors="coerce")
                perime_us = df_prospects_usd.get("Pré Aff Périmé")
                if perime_us is None:
                    perime_us = pd.Series(False, index=df_prospects_usd.index)
                preg_us = pd.to_numeric(df_prospects_usd.get("Pré G %"), errors="coerce")
                # prévisions NON périmées seulement, et potentiel d'au moins 5 %
                masque_us = masque_us | (aff_us.notna() & (aff_us != 0)
                                         & ~perime_us.fillna(False).astype(bool)
                                         & (preg_us >= 5))
            df_prospects_usd = df_prospects_usd[masque_us]
            if trier_par_rang and "Achat Rang" in df_prospects_usd.columns:
                df_prospects_usd = df_prospects_usd.sort_values(by="Achat Rang", ascending=False, na_position="last")
            else:
                df_prospects_usd = df_prospects_usd.sort_values(by=["Score", "Confiance"], ascending=[False, False], na_position="last")
            if "Achat Rang" in df_prospects_usd.columns:   # === v7 : percentile dans la liste affichée ===
                df_prospects_usd["Rang %"] = df_prospects_usd["Achat Rang"].rank(pct=True) * 100

        colonnes_a_afficher_pros_us = [c for c in colonnes_base_pros if c in df_prospects_usd.columns]
        config_description = config_largeur_description(df_prospects_usd, afficher_desc, px_par_char=6, largeur_min=80, largeur_max=320)

        styled_usd = df_prospects_usd.style.apply(surligner_prospects, axis=1)
        if afficher_var and 'Var %' in df_prospects_usd.columns:
            styled_usd = styled_usd.map(couleur_var, subset=['Var %'])
        if 'Pré G %' in df_prospects_usd.columns:
            styled_usd = styled_usd.map(couleur_preg_prospect, subset=['Pré G %'])   # === v8 : dégradé ===
        if 'Signal Aff' in df_prospects_usd.columns:
            styled_usd = styled_usd.map(couleur_signal, subset=['Signal Aff'])
        if 'Pré Aff Périmé' in df_prospects_usd.columns:
            styled_usd = styled_usd.apply(griser_pre_aff_perime, axis=1)

        st.dataframe(
            styled_usd,
            use_container_width=False, hide_index=True, height=hauteur_tableau(len(df_prospects_usd)),
            column_order=colonnes_a_afficher_pros_us,
            column_config={**config_description, **config_colonnes_communes(),
                           **config_largeur_pourquoi(df_prospects_usd)}
        )

    # --- ONGLET 4 : MÉTHODE ---
    with tab4:
        st.markdown(
            """
            ## 💰 Onglet Portefeuille — Signal de VENTE
            Pour un titre que tu **détiens déjà**, la colonne **Signal** indique quoi en faire,
            selon le **potentiel restant** (Pré G %) :

            | Signal | Condition | Idée |
            |---|---|---|
            | 🔴 **Vendre** | potentiel ≤ 5 % | objectif quasi atteint, peu de hausse restante |
            | 🟡 **À surveiller** | 5 % < potentiel < 15 % | se rapproche de l'objectif |
            | 🟢 **Attendre** | potentiel ≥ 15 % | belle marge de hausse, on conserve |

            *Les seuils et couleurs sont alignés sur la colonne Pré G %. Score / Confiance /
            Risque ne sont pas affichés dans le Portefeuille (ils servent à choisir quoi
            **acheter**, pas quoi vendre).*

            ---
            ## 🎯 Onglets Prospects — Score d'ACHAT
            Pour un titre que tu envisages d'**acheter**, on calcule trois indices.

            ### Score (0–100)
            Combine plusieurs facteurs, le **potentiel de gain restant le principal** :

            | Facteur | Échelle | Poids |
            |---|---|---|
            | Potentiel de gain (Pré G %) | 0 % → 80 % | 50 % |
            | Proximité du creux 52 sem. | sommet → creux | 18 % |
            | Dividende (Div %) | 0 % → 6 % | 12 % |
            | Momentum (Var % du jour) | −5 % → +5 % | 8 % |
            | Faible volatilité | 15 % → 80 % (inversé) | 12 % |

            👉 Point clé : le score est **normalisé par les poids des données disponibles**.
            Une donnée manquante (ex. pas de dividende) n'est pas comptée comme zéro — elle est
            simplement ignorée, donc elle ne pénalise pas injustement le titre.

            ### 🛡️ Confiance (0–100)
            Monte quand les données sont complètes (prix, cible Yahoo, cible Affaires, chaleur,
            volatilité, dividende). **Baisse fortement quand les cibles Yahoo et Affaires se
            contredisent** : si les deux analyses ne sont pas d'accord, on a moins confiance.

            ### ⚠️ Risque (0–100)
            Monte avec la **volatilité**, la **proximité du sommet 52 sem.**, un **objectif déjà
            atteint** (potentiel négatif) et une **faible confiance**.

            ### 🏷️ Signal (achat)
            Étiquette lisible déduite des trois indices :
            - **Priorité** : Score ≥ 75 et Confiance ≥ 60
            - **À surveiller** : Score ≥ 60
            - **À valider** : confiance trop faible (< 45)
            - **Risque élevé** : Risque ≥ 75 sans score exceptionnel
            - **Objectif atteint** : potentiel ≤ 0
            - **Secondaire** : le reste

            ---
            ## 🆕 Aide à la décision (v7)
            Six ajouts pour **mieux choisir la prochaine action à acheter** :

            - **🏆 Rang d'achat (0–100)** — classement composite : `45 % Score + 20 % Entrée +
              15 % Concordance + 20 % (100 − Risque)`, **moins 5 pts par titre déjà détenu dans
              le même secteur**. Les Prospects sont triés par ce rang par défaut.
            - **Concordance (0–100)** — degré d'accord entre les deux cibles (Yahoo & Les
              Affaires). 100 = cibles identiques ; faible = elles se contredisent (⚠️).
            - **Entrée (0–100)** — qualité du point d'entrée : potentiel + proximité du **creux
              52 sem.** + momentum court. Élevé = bon moment pour entrer.
            - **Garde-fou anti-aberration** — un Pré G % au-dessus du seuil (Paramètres, défaut
              200 %) est **grisé** et **neutralisé** dans le Score et le Rang (souvent une erreur
              de devise/échelle, ex. cible $US vs prix CDR).
            - **Diversification secteur** — le secteur Yahoo est récupéré ; un prospect dont le
              secteur est **déjà chargé** dans ton portefeuille est pénalisé au Rang et annoté.
            - **Pourquoi** — résumé lisible des facteurs saillants d'un titre (ex.
              « +28 % potentiel · 2 sources concordantes · proche creux 52s »).
            - **Rang %** — percentile du titre dans la liste affichée (option).

            ### 🧭 Onglet Décision
            La synthèse du jour en un coup d'œil : **Top 5 achats** (meilleur Rang), titres à
            **vendre**, replis de plus de N % depuis le **sommet 52 sem** (protection des gains),
            **changements** depuis la dernière séance et **performance des signaux passés**.

            ### 📓 Journal des signaux
            Chaque jour d'utilisation, les prospects en « Priorité » et les titres en « Vendre »
            sont archivés (avec prix) dans l'onglet **Journal** du Google Sheet. Avec le temps,
            la section Décision montre si les signaux « Priorité » ont réellement monté.

            ### ↘ Baisse depuis le sommet 52 sem (Portefeuille)
            Un titre encore gagnant peut avoir reculé fortement depuis son sommet : la colonne
            « ↘ Sommet » (rouge au-delà du seuil, Paramètres, défaut 15 %) aide à **protéger les
            gains** — l'angle mort du signal Vendre basé sur le potentiel seul.

            ### 📊 Fondamentaux légers (option)
            **P/E**, **croissance des revenus** et **marge nette** (Yahoo), en colonnes optionnelles
            — informatif, non intégré au Score.

            *Tous ces indicateurs s'activent/se masquent dans ⚙️ Paramètres → Aide à la décision.*

            ---
            *Cet outil sert à **prioriser une liste de surveillance**. Il n'est ni un conseil
            financier ni un substitut à ton jugement.*
            """
        )

except Exception as e:
    st.error(f"Erreur interceptée : {type(e).__name__} - {e}")
    with st.expander("Détails techniques"):
        import traceback
        st.code(traceback.format_exc())
