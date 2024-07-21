# Craqueur-des-fichiers-Word

Le Craqueur des fichiers Word est un outil destiné à récupérer l'accès aux documents Microsoft Word protégés par mot de passe. Ce programme utilise une approche combinant l'utilisation d'un dictionnaire de mots de passe et une attaque par force brute pour tenter de déchiffrer le mot de passe d'un fichier Word.

Fonctionnalités

- Dictionnaire de mots de passe : Utilise un fichier texte contenant une liste de mots de passe potentiels.
- Force brute : Génère toutes les combinaisons possibles de mots de passe jusqu'à une longueur maximale spécifiée.
- Logging : Enregistre les informations de débogage et les résultats de la tentative de craquage.
Prérequis

- Python 3.12.x
- Bibliothèque `msoffcrypto`
- Fichier de dictionnaire (`Dico.txt`) contenant des mots de passe potentiels

Installation

1. Clonez ce dépôt sur votre machine locale :
    
    git clone https://github.com/Djochrist/Craqueur-des-fichiers-Word.git
    cd Craqueur-des-fichiers-Word

2. Installez les dépendances requises :
    pip install msoffcrypto

Utilisation

1. Préparer le fichier dictionnaire** :
    - Créez un fichier `Dico.txt` contenant une liste de mots de passe potentiels, chaque mot de passe sur une nouvelle ligne.

2. Modifier le chemin vers le fichier dictionnaire** :
    - Ouvrez le script et modifiez la variable `chemin_dico` pour y mettre le chemin absolu vers votre fichier `Dico.txt` :

3. Exécuter le script** :
    - Exécutez le script Python :
   
    - Lorsque le script est exécuté, il vous demandera de saisir le chemin du fichier Word sécurisé :
      
      Veuillez entrer le chemin du fichier Word à craquer : 
      
    - Entrez le chemin complet vers le fichier Word protégé par mot de passe.

4. Résultat :
    - Si le mot de passe est trouvé, il sera affiché dans la console avec le temps nécessaire pour le trouver.
    - Si aucun mot de passe n'est trouvé, un message d'échec sera affiché.

Exemple de sortie
Let's go !
Veuillez entrer le chemin du fichier Word à craquer : C:\chemin\vers\fichier_protege.docx
INFO: 1000 mots de passe chargés du dictionnaire.
INFO: Mot de passe trouvé (dictionnaire) : password123 en 2.35 secondes
Mot de passe trouvé : password123 en 2.35 secondes

