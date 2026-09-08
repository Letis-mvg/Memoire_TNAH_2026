# =============================================================================
# SCRIPT : Transfert vers l'arborescence fichiers
# =============================================================================
#
# Pésentation : 
#
# Le but de ce script est d'effectuer le transfert des documents traités en amont 
# vers l'arborescence fichier. 
# 
# - Pour se faire, il recrée l'arborescence fichier
# - Scan les fichiers , en se basant sur un système de reconnaissance et de comparaison des mots_clés
# récupérés dans le plan de classement mais aussi dans le dictionnaire de synonymes créé.
#
# ==============================================================================






import os  # Importation du module système pour la gestion des dossiers et des chemins de fichiers
import shutil  # Importation du module pour copier et déplacer des fichiers sur le disque
import unicodedata  # Importation du module pour nettoyer et standardiser les caractères accentués
import pandas as pd  # Importation de pandas pour lire et manipuler le fichier Excel de référence
import re  # Importation des expressions regulières
from docx import Document  # Importation de python-docx pour extraire le texte des fichiers Word (.docx)
from pypdf import PdfReader  # Importation de pypdf pour extraire le texte des documents PDF natifs
import pdfplumber  # Importation de pdfplumber pour l'analyse fine des structures (tableaux/textes) des PDF
from pptx import Presentation  # Importation de python-pptx pour extraire le texte des présentations PowerPoint (.pptx)
import logging          # Création des journaux d'excécution
import config        # Import des paramètres de configuration



# ==================================
# CONFIGURATION DES LOGS
# ==================================

# configuration général du journal
logging.basicConfig(
   level=logging.INFO, 
   format="%(asctime)s | %(levelname)s | %(message)s",
   handlers=[
       # Sauvegarde dans le fichier "Arborescence.log"
       logging.FileHandler(
           os.path.join(config.dossier_logs, "Arborescence.log"),
           encoding="utf-8"
       ),
       # Affichage dans le terminal
       logging.StreamHandler()
   ]
)




  
        # Cette fonction convertie le chemin de fichiers de Windows en Linux (WSL)
        # Si le chemin est vide ou ne contient pas du texte, alors on ne le touche pas pour éviter de faire bugger le programme, 
        # remplace tous les antislashs par des slashs

def Conversion_chemin(chemin):
    if not chemin or pd.isna(chemin) or not isinstance(chemin,str):
        return ""

        # Transformation des slashs 
    chemin = chemin.replace('\\', '/')

        # nettoyage du préfixe "file"; l'algorithme transforme : file:///C:/Test en C:/Test
    if chemin.lower().startswith('file:'):
        chemin = chemin.split('file:', 1)[1].lstrip('/')

        # Conversion de chemins : il détecte les lettres du disque local (C:) et transforme le chemin en /mnt/c 
    if len (chemin) >=2 and chemin[1] == ':':
        drive = chemin[0].lower()
        chemin = f"/mnt/{drive}" + chemin[2:]

    # Et retourne le chemin compatible
    return chemin

# Fonction de pour la copie sécurisée des fichiers vers l'arborescence et des métadonnées qu'ils comportent; comme le nom de l'auteur
# les dates de création à travers l'utilisation de "copy2", si le serveur réseau bloque l'écriture des permissions, il récupère le fichier
# et applique ses dates d'origines


def copie_fichier(src, dest):

    # On vérifie si le dossier parent de destination existe
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    try:
        
        shutil.copy2(src, dest)
    except (PermissionError, OSError):
        shutil.copyfile(src, dest)
        try:
            st = os.stat(src)
            os.utime(dest,(st.st_atime, st.st_mtime))
        except Exception:
            pass 


# =========================================================================
#  FONCTIONS DE NORMALISATION TEXTUELLE 
# =========================================================================

 
     # La normalisation du texte se subdivise en trois parties :

     # 1- La fonction normaliser_texte :
     
     # - Vérifie que la valeur reçue est une chaîne de caractères; sinon il retourne une chaîne vide
     # - Normalise le texte de telle sorte à faciliter la comparaison


     # 2- La fonction extraire_mot_propre :

     # - Vérifie si le mot clé commence par un code numérique et qu'il contient un ou deux "_" et récupère le mot_clé

     # 3- La fonction verifier_mot_cle :

     # - Normalisation du contenu à analyser. Il fait la recherche directe.
     # - Si le mot_clé contien au moins trois caractères et qu'il est trouvé dans la cible il valide le match
     # - Il fait une recherche par synonyme_métiér


def normaliser_texte(texte):
 if not isinstance(texte, str):
     return ""
 texte_sans_accent = "".join(c for c in unicodedata.normalize("NFD", texte) if unicodedata.category(c) != "Mn") # Supprime les accents
 texte_nettoye = texte_sans_accent.replace("_", " ").replace("-", " ").replace("+", " ").replace(".", " ") # Remplace les séparateurs courants par des espaces afin de faciliter la comparaison (Pret_Des_Oeuvres.pdf = Pret Des Oeuvres.pdf)
 return " ".join(texte_nettoye.lower().split()).strip() # Met en miniscrule, supprime les espaces multiples et retire les espaces au début et à la fin (Prêt_Des_Oeuvres.pdf = pret des oeuvre.pdf)


def extraire_mot_propre(mot_cle):
 if mot_cle[0:2].isdigit() and "_" in mot_cle: # Vérifie si le mot clé commence par un code numérique et qu'il contient un "_" (05_securite)
     return mot_cle.split("_", 2)[-1] if mot_cle.count("_") >= 2 else mot_cle.split("_", 1)[-1] # S'il existe au moins deux "_" (05_securite_oeuvres), il récupère le dernier élément (oeuvres); au cas ou il y a un "_" il récupère le premier (pret)
 return mot_cle   # Si le mot clé n'a pas de code de l'arborescence, on le retourne tel qu'il est ( extraire_mot_propre(05_Graphisme) = Graphisme)

def verifier_mot_cle(mot_cle, nom_cible):
 cible_norm = normaliser_texte(nom_cible) # Nettoyage du mot_clé (suppression éventuelle du code de l'arborescence avant de le normaliser)
 mot_norm = normaliser_texte(extraire_mot_propre(mot_cle)) # Evite des confusions avec les mots au pluriel et au singulier
 if len(mot_norm) >= 3 and mot_norm in cible_norm:
     return True # Il fait la recherche directe sur nom de l'objet 
 for maitre_mot, synonymes in config.SYNONYMES_METIERS.items():   # Scan le dictionnaire de synonyme métier
     if mot_norm == normaliser_texte(maitre_mot) or mot_norm.rstrip("s") == normaliser_texte(maitre_mot).rstrip("s"): # Vérifie si le mot clé correspond au mot principal
         for syn in synonymes: # Recherche de tous les synonymes dans le texte
             if normaliser_texte(syn) in cible_norm: # Si le synonyme est trouvé, l'algoritme considère que le mot clé est trouvé.
                 return True
 return False  # Si aucun mot_clé ou synonyme n'est trouvé il passe à la suite du traitement

# =========================================================================
#  EXTRACTION MULTI-FORMATS
# =========================================================================


   # Cette fonction lit le contenu de chaque format  et extrait le texte avant de faire
   # une comparaison par score pour déterminer ou est ce que le document sera transféré.


def extraire_texte_fichier(chemin_fichier):
    _, ext = os.path.splitext(chemin_fichier.lower()) # Evite les confusions ou problèmes de majuscules et récupère l'extension

    print(f"[DEBUG] Extension détectée : {ext}") # Affiche le débug pour déterminer quel type de fichier est traité

    texte_extrait = "" 
    

    try:
        # Cas 1 : Fichiers images (elles sont ignorées)

        if ext in config.extensions_images:
            print("[DEBUG] Fichier image ignoré")
            return ""
        
        # Cas 2 : Fichiers texte
        # Il ouvre le fichier, fais les corrections des accents, évite de faire planter le script si les caractères sont illisibles

        if ext in ['.txt', '.csv']:
            print("[DEBUG] Lecture TXT/CSV")

            with open(chemin_fichier, 'r', encoding='utf-8', errors='ignore') as f:
                texte_extrait = f.read() # Fais la lecture complète du contenu
        
        # Cas 3 : Document word


        elif ext == '.docx':
            print("[DEBUG] Lecture DOCX")
            print(f"[DOCX] Traitement de : {os.path.basename(chemin_fichier)}")

            doc = Document(chemin_fichier) # Charge le document
            textes = [] # La liste temporaire qui recoit tout le texte

            # Pour les paragraphes

            for p in doc.paragraphs:
                textes.append(p.text)

            # Pour les tableaux

            for table in doc.tables:
                for row in table.rows:
                    for  cell in row.cells:
                        textes.append(cell.text)

            texte_extrait = "\n".join(textes) # Fusionne le texte extrait
            print(f"[DOCX] {os.path.basename(chemin_fichier)} -> {len(texte_extrait)} caractères")

        # Cas 4 : PowerPoint

        elif ext == '.pptx':
            print("[DEBUG] Lecture PPTX")

            prs = Presentation(chemin_fichier)
            textes_pptx = []

            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip(): # vérifie après le scan des diapositives qu'elles contiennent bien du texte.
                        textes_pptx.append(shape.text)

            texte_extrait = "\n".join(textes_pptx) 
       
       # Cas 5 : Tableurs
       
        elif ext in ['.ods', '.xlsx', '.xls']:
            print("[DEBUG] Lecture du tableur : {ext}")
            
            engine = 'odf' if ext == '.ods' else 'openpyxl' # Choix du moteur selon le format

            dict_dfs = pd.read_excel(chemin_fichier, sheet_name=None, engine=engine)
            textes_tableur = []

            for _, df_feuille in dict_dfs.items():
                contenu_texte = df_feuille.astype(str).values.flatten()
                textes_tableur.extend([t for t in contenu_texte if t != "nan" and t.strip()])

            texte_extrait = " ".join(textes_tableur)

        # Cas 6 : PDF
        elif ext == '.pdf':
            print("[DEBUG] Lecture PDF")

            with pdfplumber.open(chemin_fichier) as plumber:
                textes_pdf = []

                for page in plumber.pages:
                    txt = page.extract_text()
                    if txt:
                        textes_pdf.append(txt)

                texte_extrait = "\n".join(textes_pdf)

            if not texte_extrait.strip():
                reader = PdfReader(chemin_fichier)
                textes = [page.extract_text() for page in reader.pages if page.extract_text()]
                texte_extrait = "\n".join(textes)

 # Gestion des erreurs 
    except Exception as e:
        print(f"   [ERREUR LECTURE] Impossible de lire le contenu de {os.path.basename(chemin_fichier)} : {e}")
        logging.error(f"Impossible de lire le contenu de {chemin_fichier}: {e}")

    return texte_extrait

# =========================================================================
#  MOTEUR DE CALCUL DES SCORES 
# =========================================================================

def calculer_score_categorie(texte_contenu, nom_fichier, mots_cles_tries):
 scores = {}
 nom_fichier_norm = normaliser_texte(nom_fichier)
 texte_contenu_norm = normaliser_texte(texte_contenu)

 # Initialisation des catégories (score à 0)
 for mot_cle in mots_cles_tries:
     chemin_dest = mapping_ejection.get(mot_cle)
     if chemin_dest not in scores:
         scores[chemin_dest] = 0

 # -------------------------------------------
 # Match direct des mots-clés du reférentiel
 # -------------------------------------------

 for mot_cle in mots_cles_tries:
     chemin_dest = mapping_ejection.get(mot_cle)
     mot_propre = normaliser_texte(extraire_mot_propre(mot_cle))
     if len(mot_propre) >= config.LONGUEUR_MIN_MOT:
         if mot_propre in nom_fichier_norm:
             scores[chemin_dest] += config.SCORE_MOT_CLE_TITRE  # Bonus fort si le mot-clé exact est dans le titre
         elif mot_propre in texte_contenu_norm:
             scores[chemin_dest] += config.SCORE_MOT_CLE_CONTENU  # Dans le contenu

 # ----------------------------------------------- 
 # Match via le dictionnaire de synonymes métiers
 # -----------------------------------------------

 for maitre_mot, synonymes in config.SYNONYMES_METIERS.items():
     chemin_dest = None
     for k, v in mapping_ejection.items():
         if normaliser_texte(maitre_mot) in normaliser_texte(k):
             chemin_dest = v
             break
     if not chemin_dest:
         continue
     maitre_norm = normaliser_texte(maitre_mot)

     # Concept métier trouvé dans le nom du fichier
     if len(maitre_norm) >= config.LONGUEUR_MIN_MOT and (maitre_norm in nom_fichier_norm or nom_fichier_norm in maitre_norm):
         scores[chemin_dest] += config.SCORE_MAITRE_TITRE
  
     # recherche des synonymes
     for syn in synonymes:
         syn_norm = normaliser_texte(syn)
         if not syn_norm:
             continue

         # Synonyme complet    
         if syn_norm in nom_fichier_norm or nom_fichier_norm in syn_norm:
             scores[chemin_dest] += config.SCORE_SYNONYME

         # Synonyme composé de plusieurs mots    
         else:
             mots_syn = [m for m in syn_norm.split() if len(m) >= config.LONGUEUR_MIN_MOT]
             if mots_syn and all(m in nom_fichier_norm for m in mots_syn):
                 scores[chemin_dest] += config.SCORE_SYNONYME_COMPOSE


 
 # Journalisation technique des scores
         logging.debug(f"Scores calculés pour {nom_fichier} : {scores}")
         return scores



# =========================================================================
#  1. LECTURE ET CARTOGRAPHIE DE L'EXCEL
# =========================================================================

df = pd.read_excel(config.excel_ref, sheet_name=config.nom_onglet)
mapping_ejection = {}
liste_tous_les_dossiers = []
dossier_niv1 = dossier_niv2 = dossier_niv3 = ""
print(" Analyse de l'Excel et structuration des dossiers...")
for item in df[config.nom_colonne].dropna().astype(str).str.strip():
 if item[0:2].isdigit() and "_" in item:
     prefixe_num = "".join([c for c in item if c.isdigit() or c == "_"]).rstrip("_")
     nb_tirets = prefixe_num.count("_")
     if nb_tirets == 0:
         dossier_niv1 = item
         chemin_complet = dossier_niv1
         dossier_niv2 = dossier_niv3 = ""
     elif nb_tirets == 1:
         dossier_niv2 = item
         chemin_complet = os.path.join(dossier_niv1, dossier_niv2)
         dossier_niv3 = ""
     elif nb_tirets == 2:
         dossier_niv3 = item
         chemin_complet = os.path.join(dossier_niv1, dossier_niv2, dossier_niv3)
     liste_tous_les_dossiers.append(chemin_complet)
     mapping_ejection[item.lower()] = chemin_complet
 else:
     mots_de_la_ligne = [m.strip() for m in item.split(",")]
     dossier_actuel = os.path.join(dossier_niv1, dossier_niv2, dossier_niv3) if dossier_niv3 else (os.path.join(dossier_niv1, dossier_niv2) if dossier_niv2 else dossier_niv1)
     for mot in mots_de_la_ligne:
         if mot and len(mot) > 1:
             mapping_ejection[mot.lower()] = dossier_actuel

  # Fonction qui ignore les erreurs 

def ignore_permission_erreur(erreur):
    """ Cette fonction permet d'ignorer les erreurs de permissions lors du traitement à partir du serveur réseau """
    if isinstance(erreur, PermissionError):
        return

    

# ==================================================================
#  2. MOTEUR DE CLASSEMENT GÉNÉRIQUE ET DYNAMIQUE
# ==================================================================

def demarrer_ejection(dossier_racine, destination):
   dossier_racine = Conversion_chemin(dossier_racine)
   destination = Conversion_chemin(destination)
   nb_fichiers_transferes = 0
   dossiers_utilises = set()
   mots_cles_tries = sorted(mapping_ejection.keys(), key=len, reverse=True)
   dossier_iconographie = next(
       (
           d
           for d in liste_tous_les_dossiers
           if "iconographie" in d.lower() and "04" in d
       ),
       "04_Mise_en_espace"
   )
   print("\n Initialisation de l'arborescence cible...")
   dest_abs = os.path.abspath(config.destination)

   # --- CRÉATION DE TOUS LES SOUS DOSSIERS ---

   for dossier_officiel in liste_tous_les_dossiers:
       base_cible = (
           dossier_iconographie
           if "iconographie" in dossier_officiel.lower()
           else dossier_officiel
       )
       os.makedirs(os.path.join(config.destination, base_cible), exist_ok=True)
   fichiers_restants = []
   #config.extensions_images = list(config.extensions_images) + [".eps", ".ai", ".psd"]

   # ==========================================================================
   #  PHASE 1 : TRI DE SURFACE (TITRE, PHOTOTHÈQUE, PARENT PRÉCIS)
   # ==========================================================================

   print("\n [PHASE 1] Début du tri de surface...")
   for root, dirs, files in os.walk(config.dossier_racine, onerror=ignore_permission_erreur):
       if os.path.abspath(root) == dest_abs:
           continue
       if not files:
           continue
       cible_du_dossier_courant = None
       for mot_cle in mots_cles_tries:
           if verifier_mot_cle(mot_cle, root):
               cible_du_dossier_courant = mapping_ejection[mot_cle]
               break
       for file in files:
           src = os.path.join(root, file)
           if not os.path.exists(src):
               continue
           file_moved = False
           _, ext = os.path.splitext(file.lower())
           nom_fichier_norm = normaliser_texte(file)

           # --- CRITÉRE A : ANALYSE DU TITRE ---
           
           match_titre_trouve = False
           cible_titre = None
           for mot_cle in mots_cles_tries:
               mot_propre = normaliser_texte(extraire_mot_propre(mot_cle))
               if len(mot_propre) >= 3:
                   if len(mot_propre) == 3 and not re.search(r'\b' + re.escape(mot_propre) + r'\b', nom_fichier_norm):
                       continue
                   if mot_propre in nom_fichier_norm:
                       cible_titre = mapping_ejection[mot_cle]
                       match_titre_trouve = True
                       break
           if not match_titre_trouve:
               for maitre_mot, synonymes in config.SYNONYMES_METIERS.items():
                   maitre_norm = normaliser_texte(maitre_mot)
                   if len(maitre_norm) <= 3 and maitre_norm in nom_fichier_norm and not re.search(r'\b' + re.escape(maitre_norm) + r'\b', nom_fichier_norm):
                       continue
                   if len(maitre_norm) >= 3 and maitre_norm in nom_fichier_norm:
                       for k, v in mapping_ejection.items():
                           if maitre_norm in normaliser_texte(k):
                               cible_titre = v
                               match_titre_trouve = True
                               break
                   if match_titre_trouve: break                 
                   for syn in synonymes:
                       syn_norm = normaliser_texte(syn)
                       if syn_norm:
                           if len(syn_norm) <= 3 and syn_norm in nom_fichier_norm and not re.search(r'\b' + re.escape(syn_norm) + r'\b', nom_fichier_norm):
                               continue
                           if syn_norm in nom_fichier_norm:
                               for k, v in mapping_ejection.items():
                                   if maitre_norm in normaliser_texte(k):
                                       cible_titre = v
                                       match_titre_trouve = True
                                       break
                       if match_titre_trouve: break
                   if match_titre_trouve: break
           if match_titre_trouve and cible_titre:
               base_cible = dossier_iconographie if "iconographie" in cible_titre.lower() else cible_titre
               dest = os.path.join(config.destination, base_cible, file)
               if os.path.abspath(src).lower() != os.path.abspath(dest).lower():
                   try:
                       copie_fichier(src, dest)
                       nb_fichiers_transferes += 1
                       dossiers_utilises.add(base_cible)
                       print(f" [MATCH TITRE STRICT] {file} -> {base_cible}")
                       file_moved = True
                   except Exception as e:
                       print(f" [ERREUR MATCH TITRE] {e}")
               if file_moved: continue

           # --- CRITÉRE B : CODE PHOTOTHÈQUE ---

           nom_nettoye = file.replace("-", "_")
           prefixe_identifie = nom_nettoye.split("_")[0].upper()
           if prefixe_identifie in config.CODE_NOMMAGE_PHOTOTHEQUE:
               dest = os.path.join(config.destination, dossier_iconographie, file)
               if os.path.abspath(src).lower() != os.path.abspath(dest).lower():
                   try:
                       copie_fichier(src, dest)
                       nb_fichiers_transferes += 1
                       dossiers_utilises.add(dossier_iconographie)
                       print(f" [CODE PHOTOTHEQUE] {file} -> {dossier_iconographie}")
                       file_moved = True
                   except Exception as e:
                       print(f" [ERREUR PHOTOTHEQUE] {e}")
               if file_moved: continue

           # --- CRITÉRE C : HÉRITAGE DU DOSSIER PARENT ---

           if cible_du_dossier_courant:
               base_cible = dossier_iconographie if "iconographie" in cible_du_dossier_courant.lower() else cible_du_dossier_courant
               # Descente forcée si on hérite d'un dossier racine global (ex: "01_Concept")
               prefixe = base_cible.split("_")[0]
               if prefixe.isdigit():
                   # Cherche le premier sous-dossier officiel qui n'est pas la racine globale
                   sous_dossier_trouve = next((d for d in liste_tous_les_dossiers if d.startswith(prefixe) and d != base_cible and any(f"_{i}_" in d for i in range(10))), base_cible)
                   base_cible = sous_dossier_trouve
               dest = os.path.join(config.destination, base_cible, file)
               if os.path.abspath(src).lower() != os.path.abspath(dest).lower():
                   try:
                       copie_fichier(src, dest)
                       nb_fichiers_transferes += 1
                       dossiers_utilises.add(base_cible)
                       print(f" [PAR DOSSIER PARENT] {file} -> {base_cible}")
                       file_moved = True
                   except Exception as e:
                       print(f" [ERREUR PARENT] {e}")
               if file_moved: continue
           fichiers_restants.append((src, file, root, ext, cible_du_dossier_courant))

   # ==========================================================================
   #  PHASE 2 : ANALYSE SÉMANTIQUE ET PLAN DE SECOURS DYNAMIQUE
   # ==========================================================================

   print(f"\n [PHASE 2] Analyse sémantique et repli pour {len(fichiers_restants)} fichiers...")
   for src, file, root, ext, cible_parent_memorisee in fichiers_restants:
       if not os.path.exists(src):
           continue
       file_moved = False
       try:
           texte_extrait = extraire_texte_fichier(src)
       except Exception:
           texte_extrait = ""
       scores_fichier = calculer_score_categorie(texte_extrait, file, mots_cles_tries)
       if scores_fichier and max(scores_fichier.values()) > 0:
           cat_max = max(scores_fichier, key=scores_fichier.get)
           base_cible = dossier_iconographie if "iconographie" in cat_max.lower() else cat_max
           dest = os.path.join(config.destination, base_cible, file)
           if os.path.abspath(src).lower() != os.path.abspath(dest).lower():
               try:
                   copie_fichier(src, dest)
                   nb_fichiers_transferes += 1
                   dossiers_utilises.add(base_cible)
                   print(f" [MATCH SÉMANTIQUE INTERNE] {file} -> {base_cible}")
                   file_moved = True
               except Exception as e:
                   print(f" [ERREUR COPIE CRITÈRE] {e}")
           if file_moved: continue

       # --- STRATÉGIE DE REPLI IMAGES ---

       if ext in config.extensions_images:
           dest = os.path.join(config.destination, dossier_iconographie, file)
           if os.path.abspath(src).lower() != os.path.abspath(dest).lower():
               try:
                   copie_fichier(src, dest)
                   nb_fichiers_transferes += 1
                   dossiers_utilises.add(dossier_iconographie)
                   print(f" [GRAPHISME REPLI] {file} -> {dossier_iconographie}")
                   file_moved = True
               except Exception as e:
                   print(f" [ERREUR COPIE IMAGE] {e}")
           if file_moved: continue

       # --- PLAN DE SECOURS DYNAMIQUE ET AUTOMATIQUE ---

       if not file_moved:
           nom_dossier_actuel = os.path.basename(root)
           base_cible = cible_parent_memorisee if cible_parent_memorisee else nom_dossier_actuel
           # Descente forcée : si le plan de secours pointe sur une racine globale, on prend le premier sous-dossier numérique de ce préfixe
           prefixe = base_cible.split("_")[0]
           if prefixe.isdigit():
               sous_dossier_trouve = next((d for d in liste_tous_les_dossiers if d.startswith(prefixe) and d != base_cible and any(f"_{i}_" in d for i in range(10))), base_cible)
               base_cible = sous_dossier_trouve
           base_cible = dossier_iconographie if "iconographie" in base_cible.lower() else base_cible
           dest = os.path.join(config.destination, base_cible, file)

           try:
               if os.path.abspath(src).lower() == os.path.abspath(dest).lower():
                   print(f" [DÉJÀ EN PLACE] {file} est déjà dans le dossier final.")
                   continue
               copie_fichier(src, dest)
               nb_fichiers_transferes += 1
               dossiers_utilises.add(base_cible)

               print(f" [REPLI AUTOMATIQUE] {file} -> Classé dans : {base_cible}")

           except Exception as e:
               print(f" [ERREUR SECOURS FORCE] {e}")
               logging.exception(f" Erreur secours {e}")



   print("\n" + "=" * 50)
   print("           BILAN RESTRUCTURE ET NETTOYE")
   print("=" * 50)
   print(f"Fichiers transférés : {nb_fichiers_transferes}")
   print(f"Dossiers utilisés : {len(dossiers_utilises)}")
   print("=" * 50 + "\n")
 



if __name__ == "__main__":  # Point d'entrée standard assurant que le script est exécuté directement et non importé
 demarrer_ejection(config.dossier_racine, config.destination)  # Déclenche l'exécution globale du moteur d'éjection avec les dossiers paramétrés
 print("\n Opération terminée ! Les dossiers et fichiers ont été injectés.")  # Affiche le message final
 