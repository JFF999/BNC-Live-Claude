"""
Détection du dossier Google Drive local, indépendante de la LANGUE de Windows.

Google Drive nomme le dossier selon la langue du compte :
  - anglais  : C:\\Users\\<user>\\My Drive
  - français : C:\\Users\\<user>\\Mon Drive  (ou « Mon disque » selon les versions)
et certaines installations utilisent encore un lecteur monté (G:).

Utilisation :
    from chemins_bnc import dossier_actions
    CHEMIN_XLSX = os.path.join(dossier_actions(), "Action_2026-c_New.xlsx")
"""

import os

_NOMS_DRIVE = ("My Drive", "Mon Drive", "Mon disque")
_LECTEURS = ("G:", "H:", "I:")


def dossier_drive():
    """Renvoie le dossier racine Google Drive existant (anglais/français/lecteur)."""
    base = os.path.expanduser("~")
    candidats = [os.path.join(base, nom) for nom in _NOMS_DRIVE]
    candidats += [os.path.join(lettre + os.sep, nom) for lettre in _LECTEURS for nom in _NOMS_DRIVE]
    for c in candidats:
        if os.path.isdir(c):
            return c
    # Aucun trouvé : on renvoie le candidat par défaut (le message d'erreur du script
    # appelant affichera ce chemin, plus parlant qu'un plantage ici).
    return candidats[0]


def dossier_actions():
    return os.path.join(dossier_drive(), "Actions")
