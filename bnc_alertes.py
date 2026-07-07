"""
ALERTES BNC (hors app) : envoie un COURRIEL seulement s'il y a du NOUVEAU.

Quoi :
  - Portefeuille : titre passe en « Vendre » (potentiel restant <= 5 %),
                   variation du jour de +/- 5 % ou plus,
                   repli de plus de 15 % depuis le sommet 52 semaines.
  - Prospects    : « opportunité » = les DEUX cibles (Yahoo + Les Affaires) présentes,
                   concordantes (écart <= 20 %), et potentiel moyen >= 25 %.

Anti-bruit : l'état des alertes déjà envoyées est mémorisé dans un fichier JSON ;
on n'envoie que les NOUVELLES alertes (ou celles qui ont changé). Pas de nouveauté
= pas de courriel.

Prix : téléchargés en direct de Yahoo (yf.download groupé, robuste). Les cibles
(Pré YF / Pré Aff) viennent du Google Sheet (dernière exécution de l'app / synchro).

CONFIG COURRIEL (à créer une fois) : C:\\Users\\jfilt\\bnc_secrets\\bnc_alertes_config.json
  {
    "smtp_utilisateur": "jfilteau99@gmail.com",
    "smtp_mdp_application": "xxxx xxxx xxxx xxxx",
    "destinataire": "jfilteau99@gmail.com"
  }
  -> Gmail : le mot de passe est un « mot de passe d'application » (compte Google ->
     Sécurité -> validation en 2 étapes -> Mots de passe des applications).
  Sans ce fichier, le script journalise les alertes SANS envoyer de courriel.

À lancer avec le Python pythoncore (tâche planifiée BNC_Alertes, ~07h45).
"""

import os
import json
import smtplib
import traceback
from email.mime.text import MIMEText
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import gspread
import yfinance as yf

# ======================== CONFIGURATION ========================
CHEMIN_CRED = r"C:\Users\jfilt\bnc_secrets\compte_service.json"
CHEMIN_CONFIG = r"C:\Users\jfilt\bnc_secrets\bnc_alertes_config.json"
CHEMIN_ETAT = r"C:\Users\jfilt\bnc_secrets\bnc_alertes_etat.json"
CHEMIN_LOG = r"C:\Users\jfilt\My Drive\Actions\bnc_sync_log.txt"
NOM_GOOGLE_SHEET = "Action_2026-c_New"

SEUIL_VENDRE = 5.0        # potentiel restant (%) => Vendre
SEUIL_VARIATION = 5.0     # variation du jour (%) qui déclenche une alerte
SEUIL_BAISSE_SOMMET = 15.0  # repli (%) depuis le sommet 52 sem
SEUIL_OPPORTUNITE = 25.0  # potentiel moyen (%) mini pour une opportunité prospect
SEUIL_CONCORDANCE = 0.20  # écart max entre les 2 cibles (20 %)
# ===============================================================


def journal(message):
    horodatage = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S")
    ligne = f"[{horodatage}] [ALERTES] {message}"
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


def nombre(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "":
        return None
    for x in ('$', '%', chr(0xa0), chr(0x202f), ' '):
        s = s.replace(x, '')
    s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def entetes_map(entetes):
    return {' '.join(str(h).split()): i for i, h in enumerate(entetes)}


def lire_feuille(sh, nom):
    ws = sh.worksheet(nom)
    vals = ws.get_all_values()
    if not vals:
        return [], {}
    return vals[1:], entetes_map(vals[0])


def prix_yahoo(symboles):
    """Prix actuel, veille et sommet 52s par symbole (download groupé 1 an)."""
    out = {}
    if not symboles:
        return out
    try:
        data = yf.download(list(symboles), period="1y", interval="1d",
                           group_by="ticker", auto_adjust=True, progress=False, threads=True)
    except Exception as e:
        journal(f"yf.download en erreur : {type(e).__name__} - {e}")
        return out
    for s in symboles:
        try:
            d = data[s] if len(symboles) > 1 else data
            close = d['Close'].dropna()
            if len(close) >= 2:
                out[s] = {
                    'prix': float(close.iloc[-1]),
                    'veille': float(close.iloc[-2]),
                    'sommet': float(close.max()),
                }
        except Exception:
            continue
    return out


def construire_alertes(sh):
    alertes = {}   # cle unique -> texte (la cle sert à l'anti-bruit)

    # --- Portefeuille ---
    lignes, cols = lire_feuille(sh, "Portefeuille BNC")
    i_sym, i_yf, i_aff = cols.get("Symbole"), cols.get("Pré YF"), cols.get("Pré Aff")
    syms_port = [str(r[i_sym]).strip() for r in lignes
                 if i_sym is not None and len(r) > i_sym and str(r[i_sym]).strip() not in ("", "0")]
    marche = prix_yahoo(syms_port)

    for r in lignes:
        if i_sym is None or len(r) <= i_sym:
            continue
        sym = str(r[i_sym]).strip()
        if not sym or sym == "0" or sym not in marche:
            continue
        m = marche[sym]
        prix, veille, sommet = m['prix'], m['veille'], m['sommet']

        # cibles -> potentiel restant
        cibles = [c for c in (nombre(r[i_yf]) if i_yf is not None and len(r) > i_yf else None,
                              nombre(r[i_aff]) if i_aff is not None and len(r) > i_aff else None)
                  if c is not None and c > 0]
        if cibles and prix > 0:
            pot = (sum(cibles) / len(cibles) - prix) / prix * 100
            if pot <= SEUIL_VENDRE:
                alertes[f"vendre:{sym}"] = (f"[VENDRE] {sym} : potentiel restant {pot:+.1f} % "
                                            f"(prix {prix:.2f} $, cible {sum(cibles)/len(cibles):.2f} $)")
        if veille > 0:
            var = (prix - veille) / veille * 100
            if var >= SEUIL_VARIATION:
                alertes[f"hausse:{sym}:{datetime.now().strftime('%Y-%m-%d')}"] = \
                    f"[HAUSSE] {sym} : {var:+.1f} % aujourd'hui ({prix:.2f} $)"
            elif var <= -SEUIL_VARIATION:
                alertes[f"chute:{sym}:{datetime.now().strftime('%Y-%m-%d')}"] = \
                    f"[CHUTE] {sym} : {var:+.1f} % aujourd'hui ({prix:.2f} $)"
        if sommet > 0:
            repli = (prix - sommet) / sommet * 100
            if repli <= -SEUIL_BAISSE_SOMMET:
                alertes[f"repli:{sym}"] = (f"[REPLI] {sym} : {repli:.1f} % depuis son sommet 52 sem "
                                           f"({sommet:.2f} $ -> {prix:.2f} $) — protéger le gain ?")

    # --- Prospects : opportunités (2 cibles concordantes + gros potentiel) ---
    lignes_p, cols_p = lire_feuille(sh, "Prospects")
    j_sym, j_yf, j_aff, j_prix = (cols_p.get("Symbole"), cols_p.get("Pré YF"),
                                  cols_p.get("Pré Aff"), cols_p.get("Prix $"))
    for r in lignes_p:
        if j_sym is None or len(r) <= j_sym:
            continue
        sym = str(r[j_sym]).strip()
        if not sym or sym == "0":
            continue
        cy = nombre(r[j_yf]) if j_yf is not None and len(r) > j_yf else None
        ca = nombre(r[j_aff]) if j_aff is not None and len(r) > j_aff else None
        prix = nombre(r[j_prix]) if j_prix is not None and len(r) > j_prix else None
        if not (cy and ca and prix and cy > 0 and ca > 0 and prix > 0):
            continue
        if abs(cy - ca) / ((cy + ca) / 2) > SEUIL_CONCORDANCE:
            continue
        pot = ((cy + ca) / 2 - prix) / prix * 100
        if pot >= SEUIL_OPPORTUNITE:
            alertes[f"opportunite:{sym}"] = (f"[OPPORTUNITÉ] {sym} : +{pot:.0f} % de potentiel, "
                                             f"cibles Yahoo ({cy:.2f} $) et Affaires ({ca:.2f} $) concordantes")
    return alertes


def envoyer_courriel(sujet, corps):
    if not os.path.exists(CHEMIN_CONFIG):
        journal("Config courriel absente -> alertes journalisées seulement (voir docstring).")
        return False
    with open(CHEMIN_CONFIG, encoding="utf-8") as f:
        cfg = json.load(f)
    msg = MIMEText(corps, "plain", "utf-8")
    msg["Subject"] = sujet
    msg["From"] = cfg["smtp_utilisateur"]
    msg["To"] = cfg["destinataire"]
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as srv:
        srv.login(cfg["smtp_utilisateur"], cfg["smtp_mdp_application"].replace(" ", ""))
        srv.send_message(msg)
    return True


def main():
    if not os.path.exists(CHEMIN_CRED):
        journal(f"ERREUR : JSON compte de service introuvable : {CHEMIN_CRED}")
        return
    gc = gspread.service_account(filename=CHEMIN_CRED)
    sh = gc.open(NOM_GOOGLE_SHEET)

    alertes = construire_alertes(sh)

    # Anti-bruit : n'envoyer que les NOUVELLES alertes (clés absentes de l'état).
    etat = {}
    if os.path.exists(CHEMIN_ETAT):
        try:
            with open(CHEMIN_ETAT, encoding="utf-8") as f:
                etat = json.load(f)
        except Exception:
            etat = {}
    nouvelles = {k: v for k, v in alertes.items() if k not in etat}

    # État = alertes ACTIVES aujourd'hui (une alerte disparue puis revenue sera renvoyée).
    try:
        with open(CHEMIN_ETAT, "w", encoding="utf-8") as f:
            json.dump(alertes, f, ensure_ascii=False, indent=1)
    except Exception as e:
        journal(f"Etat non sauvegardé : {e}")

    if not nouvelles:
        journal(f"Aucune nouvelle alerte ({len(alertes)} active(s), toutes déjà signalées).")
        return

    corps = "Alertes BNC du " + datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M") + "\n\n"
    corps += "\n".join(sorted(nouvelles.values()))
    corps += "\n\n— bnc_alertes.py (seules les nouveautés sont envoyées)"
    for ligne in sorted(nouvelles.values()):
        journal(ligne)
    if envoyer_courriel(f"BNC : {len(nouvelles)} nouvelle(s) alerte(s)", corps):
        journal(f"Courriel envoyé ({len(nouvelles)} nouvelle(s) alerte(s)).")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        journal(f"ECHEC : {type(e).__name__} - {e}")
        journal(traceback.format_exc())
