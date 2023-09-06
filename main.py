"""
:author KAPCHE TEBU BAUDOUIN
:version 1.0
"""

import extraction
import creation
import creation
import time  # Module pour gérer le temps et les retards dans le code.
import verification
import insertion
import suppression
import pythoncom  # Module pour gérer la communication avec COM (Component Object Model) sur Windows.

if __name__ == "__main__":

    document = extraction.chemin_document()

    krecrute = document + "/Krecrute"
    appel = document + "/Krecrute/appel"
    candidat = document + "/Krecrute/candidat"

    creation.creation_dossier(krecrute)
    creation.creation_dossier(appel)
    creation.creation_dossier(candidat)

    creation.creation_excel("krecrute.xlsx")
    excel = krecrute + "/krecrute.xlsx"

    suppres_pdf = False
    verif = True

    init_appel = []
    init_dossier = []
    init_candidat = {}

    fichiers_ajoutes = []
    fichiers_supprimes = []

    note_garde = []

    while True:
        try:
            time.sleep(5)
            verification.verification_instance()
            fichiers_apres = extraction.obtenir_fichiers_dans_dossier(appel)

            # Comparer les listes pour détecter les ajouts et suppressions
            fichiers_ajoutes = [fichier for fichier in fichiers_apres if fichier not in init_appel]
            fichiers_supprimes = [fichier for fichier in init_appel if fichier not in fichiers_apres]
            init_appel = fichiers_apres

            for fichier in fichiers_ajoutes:
                insertion.ajout_appel(fichier)

            for fichier in fichiers_supprimes:
                suppression.suppression_appel(fichier)

            fichiers_ajoutes = []
            fichiers_supprimes = []

            for cle, valeur in init_candidat.items():
                dossier_candidature = candidat + "/" + cle
                fichiers_apres = extraction.obtenir_fichiers_dans_dossier(dossier_candidature)

                fichiers_ajoutes = [fichier for fichier in fichiers_apres if fichier not in init_candidat[cle]]
                fichiers_supprimes = [fichier for fichier in init_candidat[cle] if fichier not in fichiers_apres]

                init_candidat[cle] = fichiers_apres

                for fichier in fichiers_ajoutes:
                    insertion.ajout_cv(fichier,dossier_candidature)

                for fichier in fichiers_supprimes:
                    suppression.suppression_cv(fichier)
                    
            fichiers_ajoutes = []
            fichiers_supprimes = []

            pythoncom.CoUninitialize()

        except Exception as e:
            print("Une erreur s'est produite :", str(e))

        time.sleep(5)