"""
:author KAPCHE TEBU BAUDOUIN
:version 1.0
"""

import os
import openpyxl  # Module openpyxl pour la manipulation de fichiers Excel.
import shutil  # Module pour la manipulation de fichiers et de répertoires.
import main 

suppres_pdf=main.suppres_pdf
init_candidat=main.init_candidat
candidat=main.candidat
note_garde=main.note_garde

def suppression_extend(fiche):
    """
    Cette fonction extrait la partie du nom de fichier d'un chemin de fichier sans son extension.

    :param fiche: Le chemin du fichier dont vous voulez supprimer l'extension.
    :type fiche: str
    :return: La partie du nom de fichier avant l'extension.
    :rtype: str
    """
    # Extraire l'extension du fichier en séparant le nom de fichier à l'aide de os.path.splitext()
    fichier_sans_extension = os.path.splitext(fiche)[0]

    return fichier_sans_extension 

def suppression_note(fichier):
    """
    Cette fonction supprime des notes dans un fichier Excel en fonction du nom de fichier donné.

    :param fichier: Le nom du fichier à rechercher et supprimer dans le fichier Excel.
    :type fichier: str
    """
    while True:
        try:
            global excel 
            
            classeur = openpyxl.load_workbook(excel)
            
            feuille = classeur['notation']
            
            index = 0
            supp = True
            
            for ligne in feuille.iter_rows(min_col=6, max_col=6, values_only=True):
                if supp == False:
                    break
                index += 1
                for cellule in ligne:
                    if supp == False:
                        break
                    if cellule is not None and isinstance(cellule, str) and fichier.lower() in cellule.lower():
                        # Supprimer la ligne correspondante dans la feuille "notation"
                        feuille.delete_rows(index)

            classeur.save(excel)
            classeur.close()
            
            break
        except Exception as e:
            print(f"Une erreur s'est produite au niveau de la suppression de note : {str(e)}")

def suppression_pdf(fiche):
    """
    Cette fonction supprime un fichier PDF spécifié s'il existe.

    :param fiche: Le chemin du fichier PDF à supprimer.
    :type fiche: str
    """
    global suppres_pdf  
    
    if os.path.exists(fiche):
        os.remove(fiche)  # Supprimer le fichier PDF
        suppres_pdf = False 
        print(f"Le fichier {fiche} a été supprimé avec succès.")

def suppression_appel(fichier):
    """
    Cette fonction gère la suppression d'appels en supprimant le dossier correspondant et en mettant à jour le dictionnaire "init_candidat".
    """ 
    global init_candidat
    
    fiche = suppression_extend(fichier)
    suppression_dossier(fiche)
    del init_candidat[fiche]

def suppression_dossier(dossier):
    """
    Cette fonction supprime un dossier et son contenu s'ils existent.

    :param dossier: Le nom du dossier à supprimer.
    :type dossier: str
    """
    global candidat  
    global init_candidat
    
    for item in init_candidat[dossier]:
        item = suppression_extend(item)  
        suppression_note(item) 
    
    fichier = candidat + "/" + dossier
    
    # Supprimer le dossier s'il existe
    if os.path.exists(fichier) and os.path.isdir(fichier):
        shutil.rmtree(fichier)
        print(f"Le dossier '{fichier}' a été supprimé.")
    else:
        print(f"Le dossier '{fichier}' n'existe pas.")
    
    print("Suppression terminée")

def suppression_cv(fichier):
    """
    Cette fonction gère la suppression d'un CV et de ses informations associées.
    """
    global note_garde

    note_sauvegarde= False
    for item in note_garde:
        if item == fichier:
            note_sauvegarde = True
    if note_sauvegarde:
        return
    fiche = suppression_extend(fichier)
    suppression_note(fiche)