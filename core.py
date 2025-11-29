import os
import time
from itertools import product
from string import ascii_lowercase, digits

def analyser_fichier(chemin, log_cb, progress_cb, max_brute_length=4):
    """
    Tente de craquer le mot de passe d'un fichier Word en utilisant :
      1) le dictionnaire défini dans Crack.charger_dictionnaire()
      2) une force brute (lettres minuscules + chiffres) jusqu'à max_brute_length
    - chemin: chemin du fichier à analyser
    - log_cb(message: str): callback pour les logs (affichage dans l'UI)
    - progress_cb(percent: int): callback pour la progression (0-100)
    Retourne True si mot de passe trouvé (et decryption réussie), False sinon.
    """
    try:
        import Crack
    except Exception as e:
        log_cb(f"[ERROR] Impossible d'importer Crack.py : {e}")
        progress_cb(0)
        return False

    if not os.path.isfile(chemin):
        log_cb(f"[ERROR] Le fichier n'existe pas : {chemin}")
        progress_cb(0)
        return False

    if not chemin.lower().endswith(('.doc', '.docx')):
        log_cb(f"[WARN] Le fichier sélectionné ne semble pas être un .doc/.docx : {chemin}")

    # Charger dictionnaire
    try:
        dictionnaire = Crack.charger_dictionnaire(getattr(Crack, 'chemin_dico', ''))
    except Exception as e:
        log_cb(f"[ERROR] Échec du chargement du dictionnaire : {e}")
        dictionnaire = []

    charset = ascii_lowercase + digits
    # Estimation du nombre total d'essais (attention aux valeurs très grandes)
    try:
        dict_count = len(dictionnaire)
        brute_total = sum(len(charset) ** l for l in range(1, max_brute_length + 1))
        total_attempts = max(1, dict_count + brute_total)
    except Exception:
        dict_count = len(dictionnaire)
        brute_total = 0
        total_attempts = max(1, dict_count)

    attempts_done = 0

    def update_progress():
        percent = int(attempts_done / total_attempts * 100) if total_attempts else 0
        progress_cb(min(max(percent, 0), 100))

    log_cb(f"[INFO] Début du craquage pour : {chemin}")
    # Passe dictionnaire
    for pw in dictionnaire:
        attempts_done += 1
        if attempts_done % 10 == 0:
            log_cb(f"[TRY][DICT] Tentative #{attempts_done}/{total_attempts} : '{pw}'")
        try:
            if Crack.essayer_mot_de_passe(chemin, pw):
                update_progress()
                log_cb(f"[FOUND] Mot de passe trouvé (dictionnaire) : {pw}")
                progress_cb(100)
                return True
        except Exception as e:
            log_cb(f"[ERROR] Exception en testant '{pw}' : {e}")
        if attempts_done % 5 == 0:
            update_progress()

    # Passe force brute
    log_cb("[INFO] Dictionnaire terminé — lancement de la force brute")
    for length in range(1, max_brute_length + 1):
        for combo in product(charset, repeat=length):
            pw = ''.join(combo)
            attempts_done += 1
            # Log allégé pour éviter le spam ; signale périodiquement
            if attempts_done % 200 == 0:
                log_cb(f"[TRY][BRUTE] Tentative #{attempts_done}/{total_attempts} (exemple) : '{pw}'")
            try:
                if Crack.essayer_mot_de_passe(chemin, pw):
                    update_progress()
                    log_cb(f"[FOUND] Mot de passe trouvé (force brute) : {pw}")
                    progress_cb(100)
                    return True
            except Exception as e:
                log_cb(f"[ERROR] Exception en testant brute '{pw}' : {e}")
            if attempts_done % 50 == 0:
                update_progress()

    update_progress()
    log_cb(f"[WARN] Mot de passe non trouvé après {attempts_done} tentatives.")
    return False
