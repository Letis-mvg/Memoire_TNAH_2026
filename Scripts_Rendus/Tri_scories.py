# ====================================================================================================================================================
# SCRIPT : Tri et nettoyage des données du vrac
# ====================================================================================================================================================
# Présentation du script :
#
# L'objectif premier de ce script est d'analyser les résulter de l'export de CVS de l'outil archivistique Droid SH, afin
# afin de permettre une identification des données identifiées comme "non pérennes" :
#
# - fichiers cachés système (.DS_Store, thumb, fichiers temporaires, etc.)
# - fichier issus de la décompression de format (.traite)
# - fichiers vides
# - doublons identifiés par empreinte numérique MD5
# - dossiers devenu vide apre l'opération de tri.
#
# Les données détectées sont séparées des données saines et conduite dans un dossier d'éliminable, tout en conservabt leur provenance d'origine.
#
# ==================================================================================================================================================
 
# =====================================
# IMPORT DES MODULES
# =====================================
 
import pandas as pd     # Importation de pandas pour lire et manipuler le fichier Excel de référence
import os               # Importation du module système pour la gestion des dossiers et des chemins de fichiers
import shutil           # Module de traitement des fichiers
import logging          # Création des journaux d'excécution
import config           # Import des paramètres de configuration
 
# ==================================
# CONFIGURATION DU SYTEME DES LOGS
# ==================================
 
 
# Création d'un dossier logs s'il n'existe pas encore
os.makedirs(config.dossier_logs, exist_ok=True)
 
# Configuration général du journal
logging.basicConfig(
   level=logging.INFO,
   format="%(asctime)s | %(levelname)s | %(message)s",
   handlers=[
       # Sauvegarde dans le fichier "tri.log"
       logging.FileHandler(
           os.path.join(config.dossier_logs, "tri.log"),
           encoding="utf-8"
       ),
       # Affichage dans le terminal
       logging.StreamHandler()
   ]
)
 
 
# Traduction du language Windows vers celui de Linux (WSL) en ce qui concerne les chemins de fichiers
# Si le chemin est vide ou ne contient pas du texte, alors on ne le touche pas pour éviter de faire bugger le programme, remplace tous les antislashs par des slashs
 
 
def Conversion_chemin(chemin):
    if not chemin or pd.isna(chemin) or not isinstance(chemin,str):
        return ""
 
        # Normalisation des slash
    chemin = chemin.replace('\\', '/')
 
        # nettoyage des prefixes
    if chemin.lower().startswith('file:'):
        chemin = chemin.split('file:', 1)[1].lstrip('/')
 
        # Conversion de chemins
    if len (chemin) >=2 and chemin[1] == ':':
        drive = chemin[0].lower()
        chemin = f"/mnt/{drive}" + chemin[2:]
 
   
    return chemin
 

# ==============================
# FONCTION PRINCIPALE
# ==============================
 
'''
 Le but de cette fonction est de faire le tri  en passant par plusieurs étapes :
1- Identification et extraction des dossiers déjà vides avant le tri à parti de l'export droid
2- Lecture et harmonisation les données du csv
3- Détection des doublons, à partir de l'empreinte numérique (Hash) identifié dans la colonne "MD5_HASH" du fichier csv et création d'une liste noire ()
4- Préparer les compteurs pour avoir les chiffres à la fin du nettoyage
5- Déplacement des fichiers identidentifiés dans le dossier des éliminables
6- Nettoyage final des dossiers devenus vide et encore présente dans l'arborescence
'''
 
def Lancer_nettoyage():
   print("\n => Début de l'analyse des fichiers ...")  
 
   logging.info("=======================================")
   logging.info("DÉBUT DU PROCESSUS DE NETTOYAGE")
   logging.info("=======================================")  
 
   if not os.path.exists(config.csv_file):
       print(f"Erreur : Le fichier CSV '{config.csv_file}' est introuvable.")
       logging.error(f"Erreur : Le fichier CSV '{config.csv_file}' est introuvable.")   # On récupère le csv s'il existe, sinon on voit afficher une erreur
       return
   print(f"Fichier CSV trouvé : {config.csv_file}")
   logging.info(f"Fichier CSV trouvé : {config.csv_file}")
 
   # ===========================================================
   # 1- Identification des dossiers vides avant le début du tri
   # ===========================================================
 
   """
   Avant de faire une analyse du CSV, il est important de récupérer les dossiers pysiquement vides.
   Pourquoi?
   Tout simplement parce que l'export Droid peut reférencer uniquement les fichiers présents.
   Certains dossiers peuvent donc exister dans l'arborescence fichiers mais ne plus contenir aucun élémen.
   Ces dossiers sont déplacés dans :
        A_Eliminer/Dossiers_Vides/
   Ceci en conservant leur chemin relatif.
   """
   print("Recherche des dossiers vides avant analyse...")
   logging.info("Recherche des dossiers vides avant analyse")
   nb_dossiers_vides = 0
 
   for root, dirs, files in os.walk(config.dossier_racine, topdown=False):    # root = le dossier racine, dirs= les sous-dossiers à l'intérieur, files= les fichiers, os.walk( pour parcourir le dossier racine), topdown=False pour les dossiers imbriqués
       for dossier in dirs:     # Pour chaque dossier dans la liste des dossiers qu'on retrouve dans root (dossier racine) applique l'action suivante
           dossier_path = os.path.join(root, dossier)
           if os.path.exists(dossier_path) and not os.listdir(dossier_path):  # On vérifie si le dossier est réellement vide (listdir)
               # On calcule le chemin relatif pour garder la structure dans l'élimination
               relatif_dossier = os.path.relpath(dossier_path, config.dossier_racine)
               dest_path = os.path.join(config.dossier_elimination, "Dossiers_vides", relatif_dossier)
               os.makedirs(os.path.dirname(dest_path), exist_ok=True)
               try:
                   shutil.move(dossier_path, dest_path)
                   print(f"Dossier vide initial détecté et déplacé : {dossier}")
                   logging.info(f"Dossier vide déplacé : {dossier}")
                   nb_dossiers_vides += 1
               except Exception:
                   logging.exception(f"Erreur lors du déplacement du dossier vide : {dossier}")
   
                   logging.info(f"Le nettoyage final des dossiers vides est terminé, place à l'analyse suivante : {nb_dossiers_vides} dossiers déplacés")
 
   # ===============================================
   # 2- Lecture et harmonisation des données du csv
   # ===============================================
 
   print("Lecture de l'export CSV Droid...")
   logging.info("Lecture de l'export CSV Droid")
   try:
    # Essai en UTF-8 avant de verifier la compatibilité avec latin-1
    try:
       df= pd.read_csv(config.csv_file, sep=None, engine='python', encoding='UTF-8')
 
    except UnicodeDecodeError:
       df = pd.read_csv(config.csv_file, sep=None, engine='python', encoding='latin-1')
 
    df.columns = [c.upper() for c in df.columns] # Normalisation des noms des colonnes en majuscule
 
    print(f"CSV chargé avec succès : {len(df)} fichiers détectés.")
    logging.info(f"CSV chargé avec succès : {len(df)} fichiers détectés.")
 
   except FileNotFoundError:
       print(f"Le fichier CSV '{config.csv_file}' est introuvable.")
       logging.exception("Fichier csv introuvable.")
       return
   
 
   except Exception as e:
       print(f"Nous avons rencontré une erreur inattendue dans la lecture du fichier CSV : {e}")
       logging.exception("Erreur inattendue de lecture du fichier CSV.")
       return
 
   # =================================================================
   # 3- Identification des doublons à partir de l'empreinte numérique
   # =================================================================
   
   # Ici, l'algorithme vérifie dans la  colonne "MD5_HASH", s'il constate qu'à l'intérieur plusieurs fichiers ont la même empreinte, il garde le premier exemplaire et met le reste dans la liste noire "doublons-idx"
 
   if config.col_hash not in df.columns:
       print("Attention : La colonne MD5-HASH est manquante. Pas de détection des doublons.")
       logging.warning(
           "La colonne MD5-HASH is manquante dans le CSV."
           "Les doublons ne seront pas détectés."
       )
       doublons_idx = pd.Index([])
   else:
       df_valides = df.dropna(subset=[config.col_hash])   # df_valides ( c'est la copie du tableau qu'on fait pour éviter les cassures sur l'original), df.dropna (Drop Not Available = supprime les valeurs manquantes)
       doublons_idx = df_valides[df_valides.duplicated(subset=[config.col_hash], keep='first')].index     # Création de la liste noire des index considérés comme doublons
       print(f"Doublons détectés dans le CSV : {len(doublons_idx)}")
       logging.info(f"{len(doublons_idx)} doublons détectés.")
 
   # =============================
   # 4- Préparation des compteurs
   # =============================
 
   # Nombre total de fichiers présents dans l'export Droid
   nb_total_fichiers = len(df)
   # Nombre d'erreurs identifiés lors déplacement
   nb_erreurs_deplacement = 0
   # Compteur par catégorie
   nb_fichiers_sains, nb_doublons, nb_vides, nb_Fichiers_Systeme_TP, nb_archives_traitees = 0, 0, 0, 0, 0
   # Liste
   lignes_a_traiter = []     # Les compteurs sont à 0 dans un la variable vide : ligne_a_traiter
 
   nb_fichiers_sains = 0
 
   # =================================
   # 5- Tri et Diagnostic des données
   # =================================
 
   print("Parcours et analyse sémantique et physique de chaque ligne...")
   logging.info("Début de l'analyse des fichiers.")
 
   for index, ligne in df.iterrows():  # iterrows est utilisé pour parcourir ligne par ligne et appliquer les conditions
       
       chemin_win = str(ligne['FILE_PATH'])  # Traite le chemin comme un texte chaîne de caratères
       chemin_linux = Conversion_chemin(chemin_win)
 
 
       # Pour des besoins de sécurité
       if not chemin_linux:
        continue
 
       #  Prise en compte du suffixe .traite
       if not os.path.exists(chemin_linux) and os.path.exists(chemin_linux + ".traite"):
        chemin_linux = chemin_linux + ".traite"
        chemin_win = chemin_win + ".traite"
 
       nom = os.path.basename(chemin_linux)
       nom_minuscule = nom.lower()
       ext_fichier = os.path.splitext(nom_minuscule)[1]      # ext_excel = str(ligne.get('EXT', '')).lower()
 
       # On teste la longueur des extensions
 
       ext_sans_point = ext_fichier.lstrip('.')
 
       try :
        taille = int(ligne['SIZE'])
       except (ValueError, TypeError):
        taille = -1
 
       
       cat = None # C'est la variable qui contiendra le nom de sous-dossier
 
 
       # --- MOTEUR DE DÉTECTION ---
       # L'algoritme  à ce niveau passe au scan physique des fichiers. Il recupère et compte toutes les catérories de chacun d'eux.
       # Pour se faire il procède par filtre
 
 
       # ----------------------------------------------------------
       # Filtre 1 : Fichiers archivés (après traitement préalable)
       # ----------------------------------------------------------
 
       if ext_fichier in config.EXTENSIONS_ARCHIVES or nom_minuscule.endswith('.traite'):
        cat = "Archives_Traitees"
        nb_archives_traitees += 1
 
       # ------------------------------------------------
       # Filtre 2: Fichiers cachés et fichiers systhème
       # ------------------------------------------------
       # Si le nom du fichier commence par un des caractères identifié, Si le mot thumb ou ds-store est dans le nom du fichier
 
       elif  (nom_minuscule.startswith('.') or
           nom_minuscule.startswith('._') or
           nom_minuscule.startswith('~$') or
           'thumb' in nom_minuscule or
           'ds_store' in nom_minuscule or
           'dfont' in nom_minuscule or
           ext_fichier in config.EXTENSIONS_SYSTEME_TP or
           len(ext_sans_point) > 4): # Pour la détection des extensions anormales comme les ".Administration", ".Anef"
               
           cat = "Fichiers_Systeme_TP"
           nb_Fichiers_Systeme_TP += 1
 
       # -------------------------
       # Filtre 3: Fichiers vides
       # -------------------------
       elif taille == 0:
           cat = "Fichiers_Vides"
           nb_vides += 1
 
       # --------------------
       # Filtre 4: Doublons
       # --------------------
       elif index in doublons_idx:
           cat = "Doublons"
           nb_doublons += 1
 
       # --- Fichiers Sains ---
 
       # A la suite du tri par catégirie, le script récupère les fichiers considérés comme "sains" et les compte
       else:
           nb_fichiers_sains += 1
 
       # Si le fichier est une donnée à traiter, on l'ajoute à la  "ligne_a_traiter" pour les élements à éliminer
       if cat:
           lignes_a_traiter.append((index, chemin_win, chemin_linux, cat))
   
   # ========================
   #  AFFICHAGE DU BILAN
   # ========================
 
   print("\n" + "=" * 60)
   print("           BILAN DE L'ANALYSE ")
   print("=" * 60)
   print(f" Fichiers sains                 : {nb_fichiers_sains}")
   print(f" Doublons détectés              : {nb_doublons}")
   print(f" Fichiers système/ TP           : {nb_Fichiers_Systeme_TP}")
   print(f" Fichiers vides                 : {nb_vides}")
   print(f" Archives traitées              : {nb_archives_traitees}")
   print("-" * 60)
   print(f" TOTAL ANALYSÉ    : {nb_total_fichiers} fichiers")
   print(f" TOTAL À DÉPLACER : {len(lignes_a_traiter)} fichiers")
   print("=" * 60 + "\n")
 
   
 
   # =================
   # Confirmation
   # =================
 
   reponse = input("\n Voulez-vous exécuter le déplacement vers le dossier des éliminables ? (o/n) : ").lower()
   if reponse == 'o':
       print("\n Lancement du transfert physique vers le dossier des éliminables...")
       logging.info("Début du déplacement des fichiers.")
 
       # ---------------------------
       # Déplacement des fichiers
       # ---------------------------
 
       for idx, chemin_win, chemin_linux, cat in lignes_a_traiter:
           nom = os.path.basename(chemin_linux)
           # Vérification des fichiers .traite avant le déplacement
           if not os.path.exists(chemin_linux) and os.path.exists(chemin_linux + ".traite"):
               chemin_linux = chemin_linux + ".traite"
               chemin_win = chemin_win + ".traite"
           if os.path.exists(chemin_linux):
               # 1. Conversion propre des chemins pour WSL
               racine_linux = Conversion_chemin(config.dossier_racine)
               elimination_linux = Conversion_chemin(config.dossier_elimination)
               # 2. Calcul du chemin relatif avec des chemins 100% Linux
               relatif_fichier = os.path.relpath(chemin_linux, racine_linux)
               # 3. Construction de la destination finale
               dest = os.path.join(elimination_linux, cat, relatif_fichier)
               os.makedirs(os.path.dirname(dest), exist_ok=True)
               try:
                   shutil.move(chemin_linux, dest)  # Déplacement physique réel
                   print(f"[ACTION] Déplacé : {nom} -> Categorie: {cat}")
                   logging.info(f"Déplacé : {chemin_win} -> {cat}")
               except Exception as e:
                   print(f"[ERREUR] Échec de déplacement pour : {nom} : {e}")
                   logging.exception(f"Erreur lors du déplacement : {chemin_win} ")
                   nb_erreurs_deplacement += 1


# ===============================================================================
# 6- Vérification final de données troubles encore présente dans l'arborescence
# ===============================================================================
 
def Nettoyage_final(dossier_racine):
 
   """
   Cette fonction, à la suite de la fonction "Lancer_nettoyage" recense les dossiers vides et les envoie dans la catégorie "Dossiers_Vides"
   Ceci dans but de garantir la qualité des données conservées.
   Elle commence par inspecter les sous-dossiers à l'intérieur des grands dossiers, elle vérifie si les dossiers et sous dossiers qui restents sont réellement vides.
   Puis déplace les dossiers vides dans le dossier des éliminables.
   """
 
   print("\n[6/6] Détection finale des dossiers vides ...")
   logging.info("Début du nettoyage final des dossiers vides.")
   nb_dossiers_vides = 0 # compte je nombre de dossiers devenu vides à l'issue du tri
   # Parcours de l'arborescence du dossier source
   # tpdown=False permet de commencer la vérification par les dossiers les plus profonds
   for root, dirs, files in os.walk(config.dossier_racine, topdown=False):
       for dossier in dirs:
           chemin_complet = os.path.join(root, dossier)
           # On verifie :
           # - si le dossier est devenu vide après le tri des fichiers
           # - S'il ne contient plus de pièce
           if os.path.exists(chemin_complet) and not os.listdir(chemin_complet):
               relatif_dossier = os.path.relpath(chemin_complet, config.dossier_racine)
               dest_path = os.path.join(config.dossier_elimination, "Dossiers_vides", relatif_dossier)
               # Création de l'arborescence de destination
               os.makedirs(os.path.dirname(dest_path), exist_ok=True)
               try:
                   shutil.move(chemin_complet, dest_path)
                   print(f" OKAY : Dossier vide déplacé : {dossier}")
                   logging.info(f"Dossier devenu vide déplacé : {dossier}")
                   nb_dossiers_vides += 1
               except Exception:
                   logging.exception(f"Erreur lors du déplacement du dossier vide : {dossier}")
   print(f" ")
   logging.info(f"Nettoyage final terminé : {nb_dossiers_vides} dossiers déplacés.")
 
   # ============================
   #  ENREGISTREMENT DANS LE LOG
   # ============================
 
   logging.info("=" * 60 )
 
 
 
   
# --- Excécution du script---
if __name__ == "__main__":
  Lancer_nettoyage() # Lance dans un premier temps un premier nettoyage
  Nettoyage_final(config.dossier_racine) # Lance à la suite du premier nettoyage un nettoyage final pour identifier les dossiers restés vides après cette opération