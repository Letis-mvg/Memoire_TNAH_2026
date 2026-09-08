# =================================================================
# SCRIPT : Décompression des formats de fichiers
# =================================================================
#
# Avant de passer au tri des archives numériques d'exposition, nous
# effectuons leur décompression : c'est l'objectif du script.
#
# Une seule fonction exécute trois tâches distinctes et succéssives :
# - Scan du dossier racine et identification des archives compréssées
# à partir de leur extension
# - Extraction du contenu / décompréssion
# - Ajout de l'extrension ".traite" à la fin du nom de l'archive.
# (exemple : maquette.zip devient maquette.zip.traite )
#
# =================================================================

import os       # Pour l'interaction avec le système
import shutil   # Module de traitement des fichiers
import py7zr    # Pour la prise en compte de l'extension .7z
import rarfile  # Pour la prise en compte de l'extension .rar
import logging  # Création des journaux d'excécution
import config   # Import des paramètres de configuration


# =============================================
#  Conigaration des logs
# =============================================

os.makedirs(config.dossier_logs, exist_ok=True)

# Configuration général du journal

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        # Sauvegarde dans le fichier "decompression.log"
        logging.FileHandler(
            os.path.join(config.dossier_logs, "decompression.log"),
            encoding="utf-8"
        ),

        # Affichage dans le terminal
        logging.StreamHandler()
    ]
)

""" Cette fonction excécute la décompression de toutes les archive identifiées dans le dossier racine """

def deziper_archives():

    logging.info(" ===================================")
    logging.info("    DÉBUT DE LA DÉCOMPRESSION")
    logging.info(" ===================================")
    logging.info(f"Dossier analysé : {config.dossier_racine}")

    print(f" Début de l'exploration dans : {config.dossier_racine}")
    nb_archives =  0       # Nombre total d'archives détectés
    nb_extraits = 0        # Création du compteur d'archives extraites
    nb_erreurs = 0         # Création du compteur d'erreurs

    # Boucle pour parcourir automatiquement tous les dossiers et sous-dossiers
    # ( à travers l'utilisation du module os.walk )

    for root, dirs, files in os.walk(config.dossier_racine):

        # root va devenir le nom du dossier dans lequel le script se trouve à l'instant T

        for file in files:

            # Pour lui demander de rechercher dans chaque contenu de chaque fichier

            # Vérification et recherche des extensions d'archives

            # Vérifie si le nom du fichier se termine par les valeurs comprises dans les parenthèses

            if file.lower().endswith(tuple(config.EXTENSIONS_ARCHIVES)):
                nb_archives +=1
                chemin_complet = os.path.join(root, file)

                # Création du chemin complet du fichier en allant par le dossier racine

                logging.info(
                    f"Archive trouvée : {chemin_complet}"
                )

                # Demande au script de poursuivre l'excécution sans s'arrêter

                try:

                    extension = os.path.splitext(file)[1].lower()

                    if extension in (".zip", ".tar", ".gz"):

                        # On extrait dans le dossier actuel (root)

                        shutil.unpack_archive(
                            chemin_complet,
                            extract_dir=root
                        )

                    # -------------------
                    # Cas de .7z
                    # -------------------

                    elif extension == ".7z":
                        try:
                            with py7zr.SevenZipFile(chemin_complet,mode="r") as archive:
                                archive.extractall(path=root)
                        except Exception:
                            #Si l'excécution du script (par la bibliothèque py7zr) est face à un problème de droits d'accès, os.system excéute directement la commande dans dans le terminal Linux. Et extrat l'archive via le logiciel 7-Zip (-y répond automatiquement oui et >/dev/null masque le message)
                            os.system(f'7z x "{chemin_complet}" -o"{root}" -y > /dev/null')


                    # -----------------
                    # Cas du .rar
                    # -----------------

                    elif extension == ".rar":
                        try:
                            with rarfile.RarFile(chemin_complet) as archive:
                                archive.extractall(path=root)
                        except Exception:
                            #Si l'exyration bloque sur la permission d'erreur : unrar x c'est l'outil de Linux pour extraire les fichiers .rar. o+= forcement le remplacement si les fichiers existent déjà dans le dossier racine
                            os.system(f'unrar x -o+ "{chemin_complet}" -o"{root}/" > /dev/null')


                    else:

                        logging.warning(
                            f"Extension non gérée : {extension}"
                        )

                        continue

                    logging.info(
                        f"Archive exraite avec succès : {file}"
                    )

                    # --- OPTION SÉCURITÉ ---

                    # Une fois l'archive décompréssée et le contenu extrait,
                    # l'algorithme ajoute l'extension .traite pour la suite du traitement

                    os.rename(
                        chemin_complet,
                        chemin_complet + ".traite"
                    )

                    nb_extraits += 1

                    logging.info(
                        f"Archive renommée : {file}.traite"
                    )

                # Gestion des erreurs

                except Exception as e:

                    nb_erreurs += 1

                    logging.error(
                        f"Erreur lors de l'extraction de {file} : {e}"
                    )

    logging.info(f"Nombre total d'archives repérées : {nb_archives}")
    logging.info(f"Nombre total d'archives extraites : {nb_extraits}")
    logging.info(f"Nombre total d'erreurs d'extraction : {nb_erreurs}")
    logging.info("==== Fin de la décompréssion ====")

    print(f"\n Terminé ! {nb_extraits} archives ont été décompressées.")


if __name__ == "__main__":

    deziper_archives()