import msoffcrypto
from itertools import product
from string import ascii_lowercase, digits
import logging
import time

print ("Let's go !")

# Précisez le chemin vers le fichier dictionnaire
chemin_dico = r'c:\Dico.txt'

# Configuration du logging pour les informations de débogage
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def charger_dictionnaire(chemin_dico):
    """Charge les mots de passe à partir d'un fichier dictionnaire."""
    try:
        with open(chemin_dico, 'r', encoding='utf-8') as fichier:
            mots_de_passe = [ligne.strip() for ligne in fichier]
        logging.info(f'{len(mots_de_passe)} mots de passe chargés du dictionnaire.')
        return mots_de_passe
    except FileNotFoundError:
        logging.error(f'Le fichier {chemin_dico} n\'a pas été trouvé.')
        return []

def force_brute(longueur_max):
    """Génère tous les mots de passe possibles jusqu'à une longueur maximale."""
    caracteres = ascii_lowercase + digits
    for longueur in range(1, longueur_max + 1):
        for tentative in product(caracteres, repeat=longueur):
            yield ''.join(tentative)

def essayer_mot_de_passe(chemin_word, mot_de_passe):
    """Tente de déchiffrer le fichier Word avec un mot de passe donné."""
    try:
        with open(chemin_word, 'rb') as fichier:
            fichier_office = msoffcrypto.OfficeFile(fichier)
            fichier_office.load_key(password=mot_de_passe)
            with open('temp_dechiffre.docx', 'wb') as fichier_dechiffre:
                fichier_office.decrypt(fichier_dechiffre)
            return True
    except msoffcrypto.exceptions.InvalidKeyError:
        logging.debug(f'Erreur de clé invalide avec le mot de passe : {mot_de_passe}')
        return False
    except Exception as e:
        logging.debug(f'Exception avec le mot de passe {mot_de_passe} : {e}')
        return False

def craquer_mot_de_passe_word(chemin_word, mots_de_passe, longueur_max_brute=4):
    """Tente de craquer le mot de passe du fichier Word en utilisant le dictionnaire et la force brute."""
    temps_debut = time.time()

    # Tente d'abord les mots de passe du dictionnaire
    for mot_de_passe in mots_de_passe:
        if essayer_mot_de_passe(chemin_word, mot_de_passe):
            temps_ecoule = time.time() - temps_debut
            logging.info(f'Mot de passe trouvé (dictionnaire) : {mot_de_passe} en {temps_ecoule:.2f} secondes')
            return mot_de_passe, temps_ecoule

    # Si aucun mot de passe du dictionnaire ne fonctionne, tente la force brute
    for mot_de_passe in force_brute(longueur_max_brute):
        if essayer_mot_de_passe(chemin_word, mot_de_passe):
            temps_ecoule = time.time() - temps_debut
            logging.info(f'Mot de passe trouvé (force brute) : {mot_de_passe} en {temps_ecoule:.2f} secondes')
            return mot_de_passe, temps_ecoule

    temps_ecoule = time.time() - temps_debut
    logging.warning(f'[-] Mot de passe non trouvé après {temps_ecoule:.2f} secondes')
    return None, temps_ecoule

if __name__ == "__main__":
    chemin_word = input("Veuillez entrer le chemin du fichier Word à craquer : ")
    # Convertir le chemin fourni en une chaîne brute
    chemin_word = fr'{chemin_word}'
    mots_de_passe = charger_dictionnaire(chemin_dico)
    if mots_de_passe:
        mot_de_passe_trouve, temps_ecoule = craquer_mot_de_passe_word(chemin_word, mots_de_passe)
        if mot_de_passe_trouve:
            print(f'Mot de passe trouvé : {mot_de_passe_trouve} en {temps_ecoule:.2f} secondes')
        else:
            print('Mot de passe non trouvé')
