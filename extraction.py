"""
:author KAPCHE TEBU BAUDOUIN
:version 1.0
"""

from pdfminer.high_level import extract_text  # Module pdfminer pour extraire le texte à partir de fichiers PDF.
import io  # Module pour travailler avec des flux d'octets (bytes streams).
import pdfplumber  # Module pdfplumber pour extraire des données à partir de fichiers PDF.
import openpyxl  # Module openpyxl pour la manipulation de fichiers Excel.
import subprocess  # Module pour exécuter des commandes système.
import creation
import os
import main

excel=main.excel

def extract_cv(pdf_path):
    """
    Cette fonction extrait le texte à partir d'un fichier PDF et le retourne sous forme de liste de lignes de texte.

    :param pdf_path: Le chemin du fichier PDF à partir duquel extraire le texte.
    :type pdf_path: str
    :return: Une liste de lignes de texte extraites du PDF.
    :rtype: list[str]
    """
    text_en_tableau = []
    with open(pdf_path, 'rb') as pdf_file:
        text = extract_text(io.BytesIO(pdf_file.read()))
        text_en_tableau.extend(text.strip().split('\n'))
    return text_en_tableau

def extract_appel(pdf_path):
    """
    Cette fonction extrait le texte à partir de toutes les pages d'un fichier PDF et le retourne sous forme de liste de lignes de texte.

    :param pdf_path: Le chemin du fichier PDF à partir duquel extraire le texte.
    :type pdf_path: str
    :return: Une liste de lignes de texte extraites de toutes les pages du PDF.
    :rtype: list[str]
    """
    text_en_tableau = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            text_en_tableau.extend(text.strip().split('\n'))
    return text_en_tableau

def extract_competence():
    """
    Cette fonction extrait les compétences à partir d'un fichier Excel spécifié par la variable globale 'excel'.
    Elle renvoie une liste contenant les compétences extraites.

    :return: Une liste des compétences extraites.
    :rtype: list
    """
    while True:  
        try:
            global excel 

            competence = []  # Une liste pour stocker les compétences extraites
            
            classeur = openpyxl.load_workbook(excel)
            feuilles = ["competence"]

            for feuille_nom in feuilles:
                feuille = classeur[feuille_nom]

                for ligne in feuille.iter_rows():
                    for cellule in ligne:
                        competence.append(cellule.value)  # Ajouter la valeur de chaque cellule à la liste 'competence'
                        
            classeur.close()
            return competence

        except Exception as e:
            creation.creation_excel("krecrute.xlsx") 
            print(f"Une erreur s'est produite au niveau d'extract_competence: {str(e)}")

def chemin_document():
    """
    Cette fonction obtient le chemin du répertoire "Documents" de l'utilisateur courant sous Windows.

    :return: Le chemin du répertoire "Documents" de l'utilisateur.
    :rtype: str
    """
    # Exécutez la commande pour obtenir le chemin du répertoire "Documents"
    command = r'reg query "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Personal'
    result = subprocess.run(command, stdout=subprocess.PIPE, shell=True, text=True)

    # Analysez la sortie de la commande pour obtenir le chemin du répertoire "Documents"
    lines = result.stdout.strip().split('\n')
    documents_path_line = lines[-1]
    _, _, documents_path = documents_path_line.partition("    ")

    print("Chemin du répertoire Documents:", documents_path)
    position = documents_path.split("    ")  # Sépare la ligne en une liste ["Type", "REG_SZ", "Chemin"]
    print(position)
    chemin = position[-1].replace("\\", "/")  # Récupère le chemin et remplace les barres obliques inverses par des barres obliques normales
    if chemin.find("%USERPROFILE%") != -1:
        chemin = chemin.replace("/Documents", "")  # Si le chemin contient "%USERPROFILE%", retire "/Documents"
    return chemin

def obtenir_fichiers_dans_dossier(dossier):
    """
    Cette fonction renvoie la liste des noms de fichiers présents dans un dossier spécifié.

    :param dossier: Le chemin du dossier dans lequel vous souhaitez obtenir la liste des fichiers.
    :type dossier: str
    :return: Une liste contenant les noms de fichiers dans le dossier.
    :rtype: list[str]
    """
    return [fichier for fichier in os.listdir(dossier)]