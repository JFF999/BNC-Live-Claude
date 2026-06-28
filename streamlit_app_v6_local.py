import math
import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import gspread
from datetime import datetime
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
    </style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=300)
def heure_mise_a_jour():
    return datetime.now(ZoneInfo("America/Toronto")).strftime("%H:%M")

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
            # 1. Enlève le $, le % et les espaces
            df[col] = df[col].astype(str).str.replace('$', '', regex=False).str.replace('%', '', regex=False).str.replace(r'\s+', '', regex=True)
            # 2. Change la virgule en point
            df[col] = df[col].str.replace(',', '.', regex=False)
            # 3. Convertit en vrai chiffre mathématique ET force le mode "décimal" (float)
            df[col] = pd.to_numeric(df[col], errors='coerce').astype(float)

    # --- LE NETTOYEUR D'ENTIERS (Quantité) ---
    if 'Qtée' in df.columns:
        df['Qtée'] = df['Qtée'].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '', regex=False)
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

# --- TITRE PRINCIPAL ---
st.title("📈 BNC LIVE v6 (Google Sheet)")

heure_actuelle = heure_mise_a_jour()
taux_usdcad = obtenir_taux_change()

# --- HAUT DE PAGE : Paramètres ---
col_param, col_btn = st.columns(2)

with col_param:
    with st.popover("⚙️ Paramètres"):
        source_gain = st.selectbox("Calcul du Gain", ["Yahoo", "Affaires", "Moyenne"], index=2)

        st.markdown("---")
        st.markdown("**Affichage des Colonnes**")
        afficher_no = st.checkbox("Afficher No.", value=False)
        afficher_desc = st.checkbox("Afficher Description", value=False)
        afficher_dev = st.checkbox("Afficher Devise (Dev.)", value=False)
        afficher_compte = st.checkbox("Afficher Compte", value=False)

        afficher_var = st.checkbox("Afficher Var %", value=True)
        afficher_tendance = st.checkbox("Afficher Tendance (1m)", value=False)
        afficher_chaleur = st.checkbox("Afficher Chaleur 52 sem.", value=False)
        afficher_div = st.checkbox("Afficher Dividendes (Div %)", value=False)
        afficher_analystes = st.checkbox("Afficher Nb d'analystes", value=False)  # === V4 ===

        st.markdown("---")
        st.markdown("**Moteur de décision (v5)**")
        afficher_signal = st.checkbox("Afficher Signal", value=True)
        afficher_score = st.checkbox("Afficher Score", value=True)
        afficher_confiance = st.checkbox("Afficher Confiance", value=True)
        afficher_risque = st.checkbox("Afficher Risque", value=True)
        afficher_volatilite = st.checkbox("Afficher Volatilité", value=False)

        st.markdown("---")
        st.markdown("**Fonctionnalités Avancées**")
        activer_taux_change = st.checkbox("Taux de change actif", value=False)
        afficher_gain_jour = st.checkbox("Calculer le Gain du Jour", value=True)
        afficher_bandeau = st.checkbox("Afficher le Bandeau des Marchés", value=False)
        afficher_alertes = st.checkbox("Activer les Alertes Intelligentes", value=False)

        # === V4 : garde-fou sur la fiabilité de l'objectif Yahoo ===
        # Un targetMeanPrice basé sur 1 seul analyste ne vaut rien. On peut exiger
        # un minimum d'analystes : en dessous, l'objectif Yahoo est ignoré.
        min_analystes = st.number_input(
            "Min. d'analystes pour l'objectif Yahoo (0 = désactivé)",
            min_value=0, max_value=50, value=0, step=1
        )

with col_btn:
    if st.button(f"🔄 Rafraîchir ({heure_actuelle})"):
        st.cache_data.clear()
        st.rerun()

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
@st.cache_data(ttl=300, show_spinner=False)
def telecharger_tous_les_prix_yahoo(symboles):
    symboles = list(symboles)
    resultats = {sym: {'hist': pd.DataFrame(), 'info': {}} for sym in symboles}

    # --- 1) PRIX & HISTORIQUE : un seul appel groupé (rapide, robuste) ---
    try:
        data = yf.download(
            symboles, period="1mo", interval="1d",  # === v5 : 5d -> 1mo (volatilité) ===
            group_by="ticker", threads=True, progress=False, auto_adjust=True
        )
        for sym in symboles:
            try:
                hist = data if len(symboles) == 1 else data[sym]
                hist = hist.dropna(how="all")
                if not hist.empty:
                    resultats[sym]['hist'] = hist
            except Exception:
                pass
    except Exception:
        pass

    # --- 2) INFOS FONDAMENTALES : moins de threads pour éviter le rate-limit ---
    def fetch_info(sym):
        try:
            return sym, yf.Ticker(sym).info
        except Exception:
            return sym, {}

    with ThreadPoolExecutor(max_workers=4) as executor:  # === V4 : 10 -> 4 ===
        futures = [executor.submit(fetch_info, sym) for sym in symboles]
        for future in as_completed(futures):
            sym, info = future.result()
            resultats[sym]['info'] = info

    return resultats

# === V4 : plus de @st.cache_data ici =============================================
# construire_donnees recevait un DataFrame ET un dict contenant des DataFrames :
# Streamlit devait les HASHER à chaque appel (coûteux + cache-miss silencieux).
# Comme telecharger_tous_les_prix_yahoo est déjà caché, on recalcule simplement.
# ================================================================================
def construire_donnees(df, dict_yahoo, est_portefeuille=True, symboles_portefeuille=None):
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
                cible_us = donnees_us.get('info', {}).get('targetMeanPrice')
                hist_us = donnees_us.get('hist', pd.DataFrame())
                if cible_us is not None and not hist_us.empty and 'Close' in hist_us.columns:
                    close_us = hist_us['Close'].dropna()
                    if len(close_us) >= 1 and float(close_us.iloc[-1]) > 0:
                        prix_us = float(close_us.iloc[-1])
                        df.at[index, 'Pré 1an $ Yahoo'] = float(cible_us) * (float(prix_actuel) / prix_us)

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

            devise_off = infos_gen.get('currency')
            if devise_off:
                df.at[index, 'Devise'] = str(devise_off).upper()

            df.at[index, 'Symbole'] = f"https://ca.finance.yahoo.com/quote/{symbole_clean}"
        else:
            tendances.append(None)

    df['Tendance'] = tendances
    return df

def calculer_potentiel_gain(df, source, est_portefeuille=True, min_analystes=0):  # === V4 : param min_analystes ===
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

    if source == "Yahoo":
        cible = yahoo.fillna(affaires)
    elif source == "Affaires":
        cible = affaires.fillna(yahoo)
    else:
        temp = pd.DataFrame({'Y': yahoo, 'A': affaires})
        cible = temp.mean(axis=1, skipna=True)

    mask = (prix > 0) & cible.notna()
    df.loc[mask, 'Pré G %'] = (cible[mask] - prix[mask]) / prix[mask]

    # Enregistrement pour l'affichage
    df['Pré YF Display'] = yahoo
    df['Pré Aff Display'] = affaires

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

def calculer_score_decision(df, pour_portefeuille=False):
    df = df.copy()

    potentiel = pd.to_numeric(df.get("Pré G %"), errors="coerce")     # en %
    chaleur = pd.to_numeric(df.get("Chaleur 52s"), errors="coerce")   # 0..100
    dividende = pd.to_numeric(df.get("Div %"), errors="coerce")       # en %
    volatilite = pd.to_numeric(df.get("Volatilité 1m"), errors="coerce")  # en %
    var_jour = pd.to_numeric(df.get("Var %"), errors="coerce")        # en %
    yahoo = pd.to_numeric(df.get("Pré YF Display"), errors="coerce")  # cible $ Yahoo
    affaires = pd.to_numeric(df.get("Pré Aff Display"), errors="coerce")  # cible $ Affaires

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

    return df

def couleur_var(valeur):
    if pd.isna(valeur): return ''
    if valeur > 0: return 'color: #00cc00;'
    elif valeur < 0: return 'color: #ff4d4d;'
    return ''

def couleur_alerte_vente(valeur):
    if pd.isna(valeur): return ''
    if valeur <= 5: return 'background-color: rgba(255, 0, 0, 0.3)'
    elif valeur < 15: return 'background-color: rgba(255, 255, 0, 0.3)'
    else: return 'background-color: rgba(0, 255, 0, 0.3)'

def couleur_signal(valeur):
    couleurs = {
        "Priorité": "background-color: rgba(0, 166, 90, .22);",
        "À surveiller": "background-color: rgba(106, 169, 255, .20);",
        "À valider": "background-color: rgba(255, 209, 102, .25);",
        "Risque élevé": "background-color: rgba(217, 75, 75, .22);",
        "Objectif atteint": "background-color: rgba(127, 127, 127, .18);",
    }
    return couleurs.get(valeur, "")

def couleur_signal_portefeuille(valeur):
    # Vendre = rouge, À surveiller = jaune, Attendre = vert (comme la colonne Pré G %).
    couleurs = {
        "Vendre": "background-color: rgba(217, 75, 75, .28);",
        "À surveiller": "background-color: rgba(255, 209, 102, .28);",
        "Attendre": "background-color: rgba(0, 166, 90, .22);",
    }
    return couleurs.get(valeur, "")

def surligner_prospects(row):
    if row.get('Possede') == True: return ['background-color: rgba(255, 215, 0, 0.4)'] * len(row)
    return [''] * len(row)

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

def config_colonnes_communes():
    # Configuration d'affichage partagée par tous les tableaux (idée reprise de Codex).
    return {
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
        "Signal": st.column_config.TextColumn("Signal", width="small"),
        "Date Achat": st.column_config.DatetimeColumn("Date Achat", format="YYYY-MM-DD"),
    }

try:
    with st.spinner("Connexion à Google Sheets..."):
        df_base_portefeuille = charger_donnees_base('Portefeuille BNC')
        df_base_prospects = charger_donnees_base('Prospects')

    if 'No.' in df_base_portefeuille.columns:
        df_portefeuille_actif = df_base_portefeuille[df_base_portefeuille['No.'] != 0].reset_index(drop=True)
    else:
        df_portefeuille_actif = df_base_portefeuille.copy()

    tous_les_symboles = set()
    for df_temp in [df_portefeuille_actif, df_base_prospects]:
        if 'Symbole' in df_temp.columns:
            tous_les_symboles.update([str(s).strip() for s in df_temp['Symbole'].dropna() if str(s).strip() not in ("", "0")])
    tous_les_symboles.update(["^GSPC", "^IXIC", "^GSPTSE"])

    symboles_liste_stricte = tuple(sorted(list(tous_les_symboles)))

    with st.spinner("Mode Turbo : Chargement des marchés mondiaux..."):
        yahoo_data = telecharger_tous_les_prix_yahoo(symboles_liste_stricte)

    symboles_possedes = tuple(set(df_portefeuille_actif['Symbole'].dropna().astype(str).str.strip()))

    df_live = construire_donnees(df_portefeuille_actif, yahoo_data, est_portefeuille=True)
    df_live = calculer_potentiel_gain(df_live, source_gain, est_portefeuille=True, min_analystes=min_analystes)  # === V4 ===
    for col in ["Pré G %", "Gain %", "Var %"]:
        if col in df_live.columns: df_live[col] = pd.to_numeric(df_live[col], errors='coerce') * 100
    df_live = calculer_score_decision(df_live, pour_portefeuille=True)  # === v5 : signal de vente ===

    df_live_prospects = construire_donnees(df_base_prospects, yahoo_data, est_portefeuille=False, symboles_portefeuille=symboles_possedes)
    df_live_prospects = calculer_potentiel_gain(df_live_prospects, source_gain, est_portefeuille=False, min_analystes=min_analystes)  # === V4 ===
    for col in ["Pré G %", "Var %"]:
        if col in df_live_prospects.columns: df_live_prospects[col] = pd.to_numeric(df_live_prospects[col], errors='coerce') * 100
    df_live_prospects = calculer_score_decision(df_live_prospects)  # === v5 ===

    if afficher_bandeau:
        indices_marches = {"S&P 500": "^GSPC", "NASDAQ": "^IXIC", "TSX": "^GSPTSE"}
        cols_m = st.columns(3)
        for idx, (nom_m, sym_m) in enumerate(indices_marches.items()):
            m_data = yahoo_data.get(sym_m, {}).get('hist', pd.DataFrame())
            if not m_data.empty and len(m_data) >= 2:
                m_actuel = m_data['Close'].iloc[-1]
                m_veille = m_data['Close'].iloc[-2]
                m_var = (m_actuel - m_veille) / m_veille * 100
                m_signe = "+" if m_var > 0 else ""
                cols_m[idx].markdown(f"<div class='market-block'>**{nom_m}** : {m_actuel:,.2f} (<span style='color:{'#00cc00' if m_var > 0 else '#ff4d4d'}'>{m_signe}{m_var:.2f}%</span>)</div>", unsafe_allow_html=True)
            else:
                cols_m[idx].markdown(f"<div class='market-block'>**{nom_m}** : Indisponible</div>", unsafe_allow_html=True)
        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)

    if afficher_alertes:
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
    if afficher_signal: colonnes_base_port.append("Signal")

    if afficher_var: colonnes_base_port.append("Var %")
    if afficher_tendance: colonnes_base_port.append("Tendance")
    if afficher_chaleur: colonnes_base_port.append("Chaleur 52s")
    if afficher_div: colonnes_base_port.append("Div %")
    if afficher_volatilite: colonnes_base_port.append("Volatilité 1m")
    if afficher_analystes: colonnes_base_port.append("Nb Analystes")  # === V4 ===

    colonnes_base_port.extend(["Pré YF Display", "Pré Aff Display", "Pré G %", "Achat $", "Qtée", "Date Achat"])

    # On utilise la même logique d'affichage de base pour les prospects
    colonnes_base_pros = []
    colonnes_base_pros.append("Symbole")
    if afficher_desc: colonnes_base_pros.append("Description")
    if afficher_signal: colonnes_base_pros.append("Signal")
    if afficher_score: colonnes_base_pros.append("Score")
    if afficher_confiance: colonnes_base_pros.append("Confiance")
    if afficher_risque: colonnes_base_pros.append("Risque")
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
    colonnes_base_pros.extend(["Pré YF Display", "Pré Aff Display"])

    tab1, tab2, tab3, tab4 = st.tabs(["💰 Portefeuille", "🎯 Pros CAD", "🎯 Pros US", "📘 Méthode"])

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

        cols_s = st.columns([2.5, 2.5, 2.5, 1.8]) if afficher_gain_jour else st.columns([3, 3, 2])

        with cols_s[0]:
            st.markdown(f"<div class='stats-block' style='text-align: left; padding-top: 5px;'><p style='margin: 0px; font-size: 13px; color: gray;'>{titre_gain}</p><p style='margin: 0px; font-size: 16px; font-weight: bold;'>{gain_formate}</p></div>", unsafe_allow_html=True)

        if afficher_gain_jour:
            with cols_s[1]:
                st.markdown(f"<div class='stats-block' style='text-align: center; padding-top: 5px;'><p style='margin: 0px; font-size: 13px; color: gray;'>{titre_gain_j}</p><p style='margin: 0px; font-size: 16px; font-weight: bold; color: {'#00cc00' if gain_jour_total_net >= 0 else '#ff4d4d'};'>{"+" if gain_jour_total_net > 0 else ""}{gain_j_formate}</p></div>", unsafe_allow_html=True)

        idx_val = 2 if afficher_gain_jour else 1
        idx_tri = 3 if afficher_gain_jour else 2

        with cols_s[idx_val]:
            st.markdown(f"<div class='stats-block' style='text-align: center; padding-top: 5px;'><p style='margin: 0px; font-size: 13px; color: gray;'>{titre_valeur}</p><p style='margin: 0px; font-size: 16px; font-weight: bold;'>{valeur_formate}</p>{texte_taux}</div>", unsafe_allow_html=True)

        with cols_s[idx_tri]:
            colonne_tri = st.selectbox("Tri", ["Pré G %", "Gain %"], key="tri_portefeuille", label_visibility="collapsed")

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
        if 'Signal' in df_live.columns:
            styled_port = styled_port.map(couleur_signal_portefeuille, subset=['Signal'])

        st.dataframe(
            styled_port,
            use_container_width=False, hide_index=True, height=(len(df_live) * 35) + 43,
            column_order=colonnes_a_afficher,
            column_config={**config_description, **config_colonnes_communes()}
        )

        col_save, col_exp = st.columns(2)
        with col_save:
            if st.button("💾 Sauvegarder vers Google Sheet (Portefeuille)"):
                with st.spinner("Écriture dans Google Sheets..."):
                    succes, message = sauvegarder_donnees_dans_sheets(df_live, 'Portefeuille BNC')
                    if succes: st.success(message)
                    else: st.error(message)

    # --- ONGLET 2 : PROSPECTS CAD ---
    with tab2:
        col_min, col_max, col_sig = st.columns([1, 1, 2])
        min_score_cad = col_min.number_input("Score min", min_value=0, max_value=100, value=55, step=5, key="cad_min_score")
        max_risque_cad = col_max.number_input("Risque max", min_value=0, max_value=100, value=85, step=5, key="cad_max_risk")
        filtre_signal_cad = col_sig.multiselect(
            "Signaux", SIGNAUX,
            default=["Priorité", "À surveiller", "À valider"], key="cad_signal_filter"
        )

        df_prospects_cad = df_live_prospects[df_live_prospects['Devise'] == 'CAD'].copy()
        if "Score" in df_prospects_cad.columns:
            df_prospects_cad = df_prospects_cad[
                df_prospects_cad["Score"].fillna(0).ge(min_score_cad)
                & df_prospects_cad["Risque"].fillna(100).le(max_risque_cad)
                & df_prospects_cad["Signal"].isin(filtre_signal_cad)
            ].sort_values(by=["Score", "Confiance"], ascending=[False, False], na_position="last")

        colonnes_a_afficher_pros = [c for c in colonnes_base_pros if c in df_prospects_cad.columns]
        config_description = config_largeur_description(df_prospects_cad, afficher_desc, px_par_char=6, largeur_min=80, largeur_max=320)

        styled_cad = df_prospects_cad.style.apply(surligner_prospects, axis=1)
        if afficher_var and 'Var %' in df_prospects_cad.columns:
            styled_cad = styled_cad.map(couleur_var, subset=['Var %'])
        if 'Signal' in df_prospects_cad.columns:
            styled_cad = styled_cad.map(couleur_signal, subset=['Signal'])

        st.dataframe(
            styled_cad,
            use_container_width=False, hide_index=True, height=(len(df_prospects_cad) * 35) + 43,
            column_order=colonnes_a_afficher_pros,
            column_config={**config_description, **config_colonnes_communes()}
        )

        if st.button("💾 Sauvegarder vers Google Sheet (Prospects CAD)"):
            with st.spinner("Écriture dans Google Sheets..."):
                succes, message = sauvegarder_donnees_dans_sheets(df_prospects_cad, 'Prospects')
                if succes: st.success(message)
                else: st.error(message)

    # --- ONGLET 3 : PROSPECTS US ---
    with tab3:
        col_min_us, col_max_us, col_sig_us = st.columns([1, 1, 2])
        min_score_us = col_min_us.number_input("Score min", min_value=0, max_value=100, value=55, step=5, key="usd_min_score")
        max_risque_us = col_max_us.number_input("Risque max", min_value=0, max_value=100, value=85, step=5, key="usd_max_risk")
        filtre_signal_us = col_sig_us.multiselect(
            "Signaux", SIGNAUX,
            default=["Priorité", "À surveiller", "À valider"], key="usd_signal_filter"
        )

        df_prospects_usd = df_live_prospects[df_live_prospects['Devise'] == 'USD'].copy()
        if "Score" in df_prospects_usd.columns:
            df_prospects_usd = df_prospects_usd[
                df_prospects_usd["Score"].fillna(0).ge(min_score_us)
                & df_prospects_usd["Risque"].fillna(100).le(max_risque_us)
                & df_prospects_usd["Signal"].isin(filtre_signal_us)
            ].sort_values(by=["Score", "Confiance"], ascending=[False, False], na_position="last")

        colonnes_a_afficher_pros_us = [c for c in colonnes_base_pros if c in df_prospects_usd.columns]
        config_description = config_largeur_description(df_prospects_usd, afficher_desc, px_par_char=6, largeur_min=80, largeur_max=320)

        styled_usd = df_prospects_usd.style.apply(surligner_prospects, axis=1)
        if afficher_var and 'Var %' in df_prospects_usd.columns:
            styled_usd = styled_usd.map(couleur_var, subset=['Var %'])
        if 'Signal' in df_prospects_usd.columns:
            styled_usd = styled_usd.map(couleur_signal, subset=['Signal'])

        st.dataframe(
            styled_usd,
            use_container_width=False, hide_index=True, height=(len(df_prospects_usd) * 35) + 43,
            column_order=colonnes_a_afficher_pros_us,
            column_config={**config_description, **config_colonnes_communes()}
        )

        if st.button("💾 Sauvegarder vers Google Sheet (Prospects US)"):
            with st.spinner("Écriture dans Google Sheets..."):
                succes, message = sauvegarder_donnees_dans_sheets(df_prospects_usd, 'Prospects')
                if succes: st.success(message)
                else: st.error(message)

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
            *Cet outil sert à **prioriser une liste de surveillance**. Il n'est ni un conseil
            financier ni un substitut à ton jugement.*
            """
        )

except Exception as e:
    st.error(f"Erreur interceptée : {type(e).__name__} - {e}")
    with st.expander("Détails techniques"):
        import traceback
        st.code(traceback.format_exc())
