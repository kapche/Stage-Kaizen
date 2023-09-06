"""
:author KAPCHE TEBU BAUDOUIN
:version 1.0
"""

import os 
import creation
import suppression
import aspose.words as aw  # Module Aspose.Words pour la manipulation de documents Word.
import insertion
import main 

Krecrute=main.krecrute
appel=main.appel
candidat=main.candidat
excel=main.excel
init_appel=main.init_appel
init_candidat=main.init_candidat
init_dossier=main.init_dossier
suppres_pdf=main.suppres_pdf
verif=main.verif

def verification_instance():
    """
    Cette fonction vérifie et configure l'instance de l'application en s'assurant que les répertoires et fichiers
    nécessaires existent et sont correctement configurés.
    """
    global krecrute
    global appel
    global candidat
    global excel
    global init_appel
    global init_candidat
    global init_dossier
    
    if not os.path.exists(krecrute):
        creation.creation_dossier(krecrute)
    
    if not os.path.exists(appel):
        # Si le répertoire "candidat" existe et est un dossier, supprimez son contenu
        if os.path.exists(candidat) and os.path.isdir(candidat):
            for element in os.listdir(candidat):
                element_path = os.path.join(candidat, element)
                if os.path.isfile(element_path):
                    os.remove(element_path)
                elif os.path.isdir(element_path):
                    os.rmdir(element_path)               
        creation.creation_dossier(appel)
        init_appel = []  
        init_candidat = {}  
        init_dossier = []
    
    if not os.path.exists(candidat):
        for cle, valeur in init_candidat.items():
            for item in init_candidat[cle]:
                item = suppression.suppression_extend(item)
                suppression.suppression_note(item)
        
        creation.creation_dossier(candidat)
        init_appel = []
        init_candidat = {}
        init_dossier = []
    
    if not os.path.exists(excel):
        creation.creation_excel("krecrute.xlsx")
        init_appel = []
        init_candidat = {}
        init_dossier = []
    
    for cle, valeur in init_candidat.items():
        dossier = candidat + "/" + cle
        print(dossier)
        if not os.path.exists(dossier):
            for item in init_candidat[cle]:
                item = suppression.suppression_extend(item)
                suppression.suppression_note(item)
            creation.creation_dossier(dossier)
            init_candidat[cle] = []  

def verification_extend(fiche):
    """
    Cette fonction vérifie l'extension d'un fichier. Si l'extension est .docx ou .doc, et elle le convertir en PDF.

    :param fiche: Le chemin du fichier à vérifier et, éventuellement, à convertir en PDF.
    :type fiche: str
    :return: Le nom du fichier PDF converti ou le nom du fichier d'origine si aucune conversion n'est nécessaire.
    :rtype: str
    """
    extension = os.path.splitext(fiche)
    global suppres_pdf

    
    if extension[1] == ".docx" or extension[1] == ".doc":
        
        fichier_doc = fiche.replace('/', "\\\\") 
        fichier_pdf = os.path.splitext(fichier_doc)[0] + ".pdf"

        while True:
            try:
                # Utilisez la bibliothèque Aspose.Words pour convertir le fichier .doc/.docx en PDF
                doc = aw.Document(fichier_doc)
                doc.save(fichier_pdf)
                break 

            except Exception as e:
                print(f"Une erreur s'est produite au niveau de la verification d'extention : {str(e)}")

        suppres_pdf = True  
        pdf = list(extension)
        pdf[1] = ".pdf"
        fichier_pdf_final = ""
        for item in pdf:
            fichier_pdf_final += item  

    return fichier_pdf_final

def verification_mot(profil, indexe, feuille, index, excel, competence, besoin, classeur, nom):
    """
    Cette fonction vérifie si un profil existe déjà dans la feuille "projet" d'un fichier Excel et, le cas échéant,
    appelle la fonction ajout_metier pour ajouter les informations à la feuille "metier".

    :param profil: Les informations du profil à vérifier.
    :type profil: list
    :param indexe: L'index de la colonne dans la feuille "projet" où chercher le nom du profil.
    :type indexe: int
    :param feuille: La feuille "projet" du classeur Excel.
    :type feuille: openpyxl.Worksheet
    :param index: L'index de la ligne actuelle dans la feuille "projet".
    :type index: int
    :param excel: Le chemin du fichier Excel.
    :type excel: str
    :param competence: Les compétences du profil.
    :type competence: list
    :param besoin: Les besoins du projet.
    :type besoin: list
    :param classeur: Le classeur Excel.
    :type classeur: openpyxl.Workbook
    :param nom: Le nom du dossier.
    :type nom: str
    """
    global verif  
    for ligne in feuille.iter_rows(min_col=indexe, max_col=indexe, values_only=True):
        if verif == False:
            break
        index += 1
        for cellule in ligne:
            if verif == False:  
                break
            if cellule is not None and isinstance(cellule, str) and profil[0].replace(" ", "").lower() in cellule.lower():
                # Vérifiez si le nom du profil (en minuscules et sans espaces) est dans la cellule
                insertion.ajout_metier(index, excel, competence, besoin, profil, classeur, feuille, nom)
                verif = False 