import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import gspread
from datetime import datetime
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed

st.set_page_config(page_title="Portefeuille BNC", layout="wide")

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

        /* Ajustement largeur minimale pour les cases Min % et Max % */
        div[data-testid="stHorizontalBlock"]:has(div[data-testid="stNumberInput"]) {
            flex-direction: row !important; flex-wrap: nowrap !important; gap: 15px !important;
        }
        div[data-testid="stHorizontalBlock"]:has(div[data-testid="stNumberInput"]) > div {
            width: auto !important; min-width: 110px !important; flex: none !important;
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

# --- CONNEXION ET AUTHENTIFICATION GOOGLE SHEETS ---
def connecter_google_sheets():
    info_cles = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(info_cles)
    return gc.open("Action_2026-c")

@st.cache_data(ttl=300)
def charger_donnees_base(nom_feuille):
    sh = connecter_google_sheets()
    feuille = sh.worksheet(nom_feuille)
    donnees = feuille.get_all_records()
    df = pd.DataFrame(donnees)

    # Nettoyage des noms de colonnes
    df.columns = [str(c).replace('\n', ' ').replace('\r', '') for c in df.columns]
    df.columns = [' '.join(c.split()) for c in df.columns]

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

    return df

def sauvegarder_donnees_dans_sheets(df_live, nom_feuille):
    try:
        sh = connecter_google_sheets()
        feuille = sh.worksheet(nom_feuille)

        # 1. Lire la feuille Google exactement telle qu'elle est structurée
        toutes_les_valeurs = feuille.get_all_values()
        if not toutes_les_valeurs: return False

        entetes = toutes_les_valeurs[0]
        if 'Symbole' not in entetes: return False

        col_symbole_idx = entetes.index('Symbole')

        # 2. Définir quelles colonnes l'application a le droit de mettre à jour
        colonnes_cibles = {
            'Prix $': 'Prix $',
            'Pré G %': 'Pré G %',
            'Gain %': 'Gain %',
            'Var %': 'Var %',
            'Pré YF': 'Pré 1an $ Yahoo',
            'Gain $': 'Gain $'
        }

        colonnes_a_ecrire = {}
        for sheet_col_name, df_col_name in colonnes_cibles.items():
            if sheet_col_name in entetes:
                idx = entetes.index(sheet_col_name)

                # --- PARE-FEU DE SÉCURITÉ ADAPTATIF ---
                # Si c'est le portefeuille, on bloque de A à I (indices 0 à 8)
                if nom_feuille == 'Portefeuille BNC' and idx <= 8:
                    continue
                # Si c'est un onglet de prospects, on bloque de A à C (indices 0 à 2)
                elif nom_feuille == 'Prospects' and idx <= 2:
                    continue

                # Google Sheets utilise des indices basés sur 1 (A=1, B=2...)
                colonnes_a_ecrire[idx + 1] = df_col_name

        # === Description : colonne TEXTE, traitée à part ===
        # Le pare-feu numérique ci-dessus la bloquerait (elle est en zone protégée).
        # On l'autorise ici, mais on n'écrit JAMAIS par-dessus une description déjà
        # présente : seules les cellules VIDES du Sheet sont remplies (avec le nom Yahoo).
        idx_description = entetes.index('Description') if 'Description' in entetes else None

        if not colonnes_a_ecrire and idx_description is None:
            return False

        # 3. Créer un dictionnaire mémoire ultra-rapide des données actuelles
        dict_donnees_live = {}
        for _, row in df_live.iterrows():
            sym = row.get('Symbole Brut')
            if pd.notna(sym) and str(sym).strip() != "":
                dict_donnees_live[str(sym).strip()] = row

        mises_a_jour = []

        # 4. Scanner le Google Sheets ligne par ligne et préparer l'injection massive
        for row_index_list, row_data in enumerate(toutes_les_valeurs):
            if row_index_list == 0: continue # Sauter la ligne des titres

            if len(row_data) > col_symbole_idx:
                symbole_feuille = str(row_data[col_symbole_idx]).strip()

                if symbole_feuille and symbole_feuille in dict_donnees_live:
                    row_live = dict_donnees_live[symbole_feuille]
                    ligne_gspread = row_index_list + 1

                    for gspread_col_idx, df_col_name in colonnes_a_ecrire.items():
                        valeur = row_live.get(df_col_name)
                        if pd.notna(valeur):
                            # Si la colonne est un pourcentage, on la remet en mode décimal (0.XX)
                            if '%' in df_col_name:
                                valeur_finale = round(float(valeur) / 100.0, 4)
                            else:
                                valeur_finale = round(float(valeur), 2)

                            mises_a_jour.append({
                                'range': gspread.utils.rowcol_to_a1(ligne_gspread, gspread_col_idx),
                                'values': [[valeur_finale]]
                            })

                    # --- Description (texte) : on remplit UNIQUEMENT les cellules vides ---
                    if idx_description is not None:
                        valeur_desc = row_live.get('Description')
                        cellule_actuelle = row_data[idx_description] if len(row_data) > idx_description else ""
                        if (pd.notna(valeur_desc) and str(valeur_desc).strip() != ""
                                and str(cellule_actuelle).strip() == ""):
                            mises_a_jour.append({
                                'range': gspread.utils.rowcol_to_a1(ligne_gspread, idx_description + 1),
                                'values': [[str(valeur_desc)]]
                            })

        # 5. Envoyer toutes les données d'un seul coup
        if mises_a_jour:
            feuille.batch_update(mises_a_jour)
            return True

        return False

    except Exception as e:
        st.error(f"Erreur d'écriture Google Sheets : {e}")
        return False

def preparer_export_csv(df):
    df_export = df.copy()
    if 'Symbole Brut' in df_export.columns:
        df_export['Symbole'] = df_export['Symbole Brut']
        df_export = df_export.drop(columns=['Symbole Brut'])
    if 'Tendance' in df_export.columns:
        df_export = df_export.drop(columns=['Tendance'])
    return df_export.to_csv(index=False, sep=';').encode('utf-8-sig')

# --- TITRE PRINCIPAL ---
st.title("📈 BNC LIVE v4 Claude")

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
        afficher_tendance = st.checkbox("Afficher Tendance (5j)", value=False)
        afficher_chaleur = st.checkbox("Afficher Chaleur 52 sem.", value=False)
        afficher_div = st.checkbox("Afficher Dividendes (Div %)", value=False)
        afficher_analystes = st.checkbox("Afficher Nb d'analystes", value=False)  # === V4 ===

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
            symboles, period="5d", interval="1d",
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
    df['Gain Jour $'] = 0.0
    df['Symbole Brut'] = ""
    tendances = []

    for index, row in df.iterrows():
        symbole = row.get('Symbole')
        if pd.notna(symbole) and str(symbole).strip() != "":
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
                prix_actuel = infos['Close'].iloc[-1]
                prix_veille = infos['Close'].iloc[-2]

                df.at[index, 'Prix $'] = prix_actuel
                df.at[index, 'Var %'] = (prix_actuel - prix_veille) / prix_veille

                tendances.append(infos['Close'].tolist())

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

def surligner_prospects(row):
    if row.get('Possede') == True: return ['background-color: rgba(255, 215, 0, 0.4)'] * len(row)
    return [''] * len(row)

def calculer_priorite(df):
    # Signal de priorité d'un prospect, en étoiles (1 à 5), combinant :
    #   - Potentiel de gain (Pré G %)   : poids 50 %  -> 0 % = 0, 100 %+ = max
    #   - Proximité du creux 52 sem.    : poids 30 %  -> au creux = max, au sommet = 0
    #   - Fiabilité (Nb analystes)      : poids 20 %  -> 0 = 0, 20+ analystes = max
    # Échelles ABSOLUES : le score reste stable quel que soit le filtrage Min/Max.
    df = df.copy()
    if 'Pré G %' not in df.columns:
        df['Priorité'] = ""
        return df

    preg = pd.to_numeric(df.get('Pré G %'), errors='coerce')
    chaleur = pd.to_numeric(df.get('Chaleur 52s'), errors='coerce')
    nb = pd.to_numeric(df.get('Nb Analystes'), errors='coerce')

    score_gain = (preg / 100.0).clip(0, 1).fillna(0)            # 0 % -> 0 ; 100 %+ -> 1
    score_creux = ((100.0 - chaleur) / 100.0).clip(0, 1).fillna(0.5)  # manquant = neutre
    score_fiab = (nb / 20.0).clip(0, 1).fillna(0)               # 0 ou manquant -> 0

    score = 0.5 * score_gain + 0.3 * score_creux + 0.2 * score_fiab

    # Score 0..1 -> 1 à 5 étoiles
    etoiles = (score * 4 + 1).round().clip(1, 5)
    df['Priorité'] = etoiles.map(lambda n: "⭐" * int(n) if pd.notna(n) else "")
    return df

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
            tous_les_symboles.update([str(s).strip() for s in df_temp['Symbole'].dropna() if str(s).strip() != ""])
    tous_les_symboles.update(["^GSPC", "^IXIC", "^GSPTSE"])

    symboles_liste_stricte = tuple(sorted(list(tous_les_symboles)))

    with st.spinner("Mode Turbo : Chargement des marchés mondiaux..."):
        yahoo_data = telecharger_tous_les_prix_yahoo(symboles_liste_stricte)

    symboles_possedes = tuple(set(df_portefeuille_actif['Symbole'].dropna().astype(str).str.strip()))

    df_live = construire_donnees(df_portefeuille_actif, yahoo_data, est_portefeuille=True)
    df_live = calculer_potentiel_gain(df_live, source_gain, est_portefeuille=True, min_analystes=min_analystes)  # === V4 ===
    for col in ["Pré G %", "Gain %", "Var %"]:
        if col in df_live.columns: df_live[col] = pd.to_numeric(df_live[col], errors='coerce') * 100

    df_live_prospects = construire_donnees(df_base_prospects, yahoo_data, est_portefeuille=False, symboles_portefeuille=symboles_possedes)
    df_live_prospects = calculer_potentiel_gain(df_live_prospects, source_gain, est_portefeuille=False, min_analystes=min_analystes)  # === V4 ===
    for col in ["Pré G %", "Var %"]:
        if col in df_live_prospects.columns: df_live_prospects[col] = pd.to_numeric(df_live_prospects[col], errors='coerce') * 100

    # === Signal de priorité (Pros uniquement), calculé une fois pour les deux onglets ===
    df_live_prospects = calculer_priorite(df_live_prospects)

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

    if afficher_var: colonnes_base_port.append("Var %")
    if afficher_tendance: colonnes_base_port.append("Tendance")
    if afficher_chaleur: colonnes_base_port.append("Chaleur 52s")
    if afficher_div: colonnes_base_port.append("Div %")
    if afficher_analystes: colonnes_base_port.append("Nb Analystes")  # === V4 ===

    colonnes_base_port.extend(["Pré YF Display", "Pré Aff Display", "Pré G %", "Achat $", "Qtée", "Date Achat"])

    # On utilise la même logique d'affichage de base pour les prospects
    colonnes_base_pros = []
    colonnes_base_pros.append("Symbole")
    colonnes_base_pros.append("Priorité")   # signal de priorité (étoiles), bien visible
    if afficher_desc: colonnes_base_pros.append("Description")
    if afficher_dev: colonnes_base_pros.append("Dev.")
    if afficher_compte: colonnes_base_pros.append("Compte")
    colonnes_base_pros.append("Prix $")

    if afficher_var: colonnes_base_pros.append("Var %")
    if afficher_tendance: colonnes_base_pros.append("Tendance")
    if afficher_chaleur: colonnes_base_pros.append("Chaleur 52s")
    if afficher_div: colonnes_base_pros.append("Div %")
    if afficher_analystes: colonnes_base_pros.append("Nb Analystes")  # === V4 ===
    colonnes_base_pros.extend(["Pré YF Display", "Pré Aff Display", "Pré G %"])

    # === V4 : config réutilisable pour la colonne Nb Analystes ===
    config_nb_analystes = {"Nb Analystes": st.column_config.NumberColumn("Nb An.", format="%d")}

    # === Largeur de la colonne Description : calculée par onglet (voir config_largeur_description) ===

    tab1, tab2, tab3 = st.tabs(["💰 Portefeuille", "🎯 Pros CAD", "🎯 Pros US"])

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

        df_live = df_live.sort_values(by="Pré G %" if colonne_tri == "Pré G %" else "Gain %", ascending=(colonne_tri == "Pré G %"))

        colonnes_a_afficher = [c for c in colonnes_base_port if c in df_live.columns]
        config_description = config_largeur_description(df_live, afficher_desc)

        st.dataframe(
            df_live.style.map(couleur_alerte_vente, subset=['Pré G %']).map(couleur_var, subset=['Var %'] if afficher_var else []),
            use_container_width=False, hide_index=True, height=(len(df_live) * 35) + 43,
            column_order=colonnes_a_afficher,
            column_config={
                **config_description,   # === largeur Description ===
                "No.": st.column_config.NumberColumn("No.", format="%d"),
                "Qtée": st.column_config.NumberColumn("Qtée", format="%d"),
                "Symbole": st.column_config.LinkColumn("Symbole", display_text=r"https://ca\.finance\.yahoo\.com/quote/(.*)"),
                "Pré G %": st.column_config.NumberColumn(format="%.1f %%"), "Prix $": st.column_config.NumberColumn(format="$ %.2f"),
                "Var %": st.column_config.NumberColumn(format="%.1f %%"), "Tendance": st.column_config.LineChartColumn("Tendance (5j)"),
                "Chaleur 52s": st.column_config.ProgressColumn("♨️ 52 sem.", format="%.0f %%", min_value=0, max_value=100),
                "Div %": st.column_config.NumberColumn("Div %", format="%.2f %%"),
                **config_nb_analystes,  # === V4 ===
                "Pré YF Display": st.column_config.NumberColumn("Pré YF", format="$ %.2f"),
                "Pré Aff Display": st.column_config.NumberColumn("Pré Aff", format="$ %.2f"),
                "Achat $": st.column_config.NumberColumn(format="$ %.2f"),
                "Gain %": st.column_config.NumberColumn(format="%.1f %%"), "Gain $": st.column_config.NumberColumn(format="$ %.2f"),
                "Date Achat": st.column_config.DatetimeColumn(format="YYYY-MM-DD")
            }
        )

        col_save, col_exp = st.columns(2)
        with col_save:
            if st.button("💾 Sauvegarder les données (Google Sheets)"):
                with st.spinner("Synchronisation des données dynamiques en cours..."):
                    succes = sauvegarder_donnees_dans_sheets(df_live, 'Portefeuille BNC')
                    if succes: st.success("Données dynamiques mises à jour avec succès !")
                    else: st.error("Échec de l'écriture.")

    # --- ONGLET 2 : PROSPECTS CAD ---
    with tab2:
        col_min, col_max, _ = st.columns([1, 1, 2])
        min_cad = col_min.number_input("Min %", min_value=-100, max_value=500, value=25, step=5, key="min_cad")
        max_cad = col_max.number_input("Max %", min_value=-100, max_value=500, value=100, step=5, key="max_cad")

        df_prospects_cad = df_live_prospects[df_live_prospects['Devise'] == 'CAD']
        if "Pré G %" in df_prospects_cad.columns:
            df_prospects_cad = df_prospects_cad[(df_prospects_cad["Pré G %"].notna()) & (df_prospects_cad["Pré G %"] >= min_cad) & (df_prospects_cad["Pré G %"] <= max_cad)].sort_values(by="Pré G %", ascending=False)

        colonnes_a_afficher_pros = [c for c in colonnes_base_pros if c in df_prospects_cad.columns]
        config_description = config_largeur_description(df_prospects_cad, afficher_desc, px_par_char=6, largeur_min=80, largeur_max=320)

        st.dataframe(
            df_prospects_cad.style.apply(surligner_prospects, axis=1).map(couleur_var, subset=['Var %'] if afficher_var else []),
            use_container_width=False, hide_index=True, height=(len(df_prospects_cad) * 35) + 43,
            column_order=colonnes_a_afficher_pros,
            column_config={
                **config_description,   # === largeur Description ===
                "Priorité": st.column_config.TextColumn("⭐ Prio", width="small"),
                "Qtée": st.column_config.NumberColumn("Qtée", format="%d"),
                "Symbole": st.column_config.LinkColumn("Symbole", display_text=r"https://ca\.finance\.yahoo\.com/quote/(.*)"),
                "Pré G %": st.column_config.NumberColumn(format="%.1f %%"), "Prix $": st.column_config.NumberColumn(format="$ %.2f"),
                "Var %": st.column_config.NumberColumn(format="%.1f %%"), "Tendance": st.column_config.LineChartColumn("Tendance (5j)"),
                "Chaleur 52s": st.column_config.ProgressColumn("♨️ 52 sem.", format="%.0f %%", min_value=0, max_value=100),
                "Div %": st.column_config.NumberColumn("Div %", format="%.2f %%"),
                **config_nb_analystes,  # === V4 ===
                "Pré YF Display": st.column_config.NumberColumn("Pré YF", format="$ %.2f"),
                "Pré Aff Display": st.column_config.NumberColumn("Pré Aff", format="$ %.2f")
            }
        )

        if st.button("💾 Sauvegarder les données Prospects CAD (Sheets)"):
            with st.spinner("Écriture..."):
                if sauvegarder_donnees_dans_sheets(df_prospects_cad, 'Prospects'): st.success("Sheets mis à jour !")

    # --- ONGLET 3 : PROSPECTS US ---
    with tab3:
        col_min_us, col_max_us, _ = st.columns([1, 1, 2])
        min_us = col_min_us.number_input("Min %", min_value=-100, max_value=500, value=25, step=5, key="min_us")
        max_us = col_max_us.number_input("Max %", min_value=-100, max_value=500, value=100, step=5, key="max_us")

        df_prospects_usd = df_live_prospects[df_live_prospects['Devise'] == 'USD']
        if "Pré G %" in df_prospects_usd.columns:
            df_prospects_usd = df_prospects_usd[(df_prospects_usd["Pré G %"].notna()) & (df_prospects_usd["Pré G %"] >= min_us) & (df_prospects_usd["Pré G %"] <= max_us)].sort_values(by="Pré G %", ascending=False)

        colonnes_a_afficher_pros_us = [c for c in colonnes_base_pros if c in df_prospects_usd.columns]
        config_description = config_largeur_description(df_prospects_usd, afficher_desc, px_par_char=6, largeur_min=80, largeur_max=320)

        st.dataframe(
            df_prospects_usd.style.apply(surligner_prospects, axis=1).map(couleur_var, subset=['Var %'] if afficher_var else []),
            use_container_width=False, hide_index=True, height=(len(df_prospects_usd) * 35) + 43,
            column_order=colonnes_a_afficher_pros_us,
            column_config={
                **config_description,   # === largeur Description ===
                "Priorité": st.column_config.TextColumn("⭐ Prio", width="small"),
                "Qtée": st.column_config.NumberColumn("Qtée", format="%d"),
                "Symbole": st.column_config.LinkColumn("Symbole", display_text=r"https://ca\.finance\.yahoo\.com/quote/(.*)"),
                "Pré G %": st.column_config.NumberColumn(format="%.1f %%"), "Prix $": st.column_config.NumberColumn(format="$ %.2f"),
                "Var %": st.column_config.NumberColumn(format="%.1f %%"), "Tendance": st.column_config.LineChartColumn("Tendance (5j)"),
                "Chaleur 52s": st.column_config.ProgressColumn("♨️ 52 sem.", format="%.0f %%", min_value=0, max_value=100),
                "Div %": st.column_config.NumberColumn("Div %", format="%.2f %%"),
                **config_nb_analystes,  # === V4 ===
                "Pré YF Display": st.column_config.NumberColumn("Pré YF", format="$ %.2f"),
                "Pré Aff Display": st.column_config.NumberColumn("Pré Aff", format="$ %.2f")
            }
        )

        if st.button("💾 Sauvegarder les données Prospects US (Sheets)"):
            with st.spinner("Écriture..."):
                if sauvegarder_donnees_dans_sheets(df_prospects_usd, 'Prospects'): st.success("Sheets mis à jour !")

except Exception as e:
    st.error(f"Erreur interceptée : {type(e).__name__} - {e}")
    with st.expander("Détails techniques"):
        import traceback
        st.code(traceback.format_exc())
