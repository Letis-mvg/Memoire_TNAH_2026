# =========================================================================
#  SCRIPT DE CONFIGURATION ET PARAMETRES
# =========================================================================


"""

Ce fichier constitue le cerveau de notre logique d'automatisation. Il contient toute la gogique archivistique (synonymes métiers, règles de la photothèque du MAD et les extensions à éliminer)
Il ne comporte aucune donnée confidentielle de l'institution.

Utilisé dans les trois scripts :

- Dossier racine

Script de tri :

- Dossier des éliminables
- Fichier CSV des exports Droid
- Colonne "MD5_HASH"


Script de transfert dans l'arborescence :

- Dossier d'arborescence (destination)
- Reférentiel d'arborescence fichier
- Onglet "typologie mots clés"
- Colonne "Arborescence + mots clés"
- Synonymes métiers
- Code de nommage de la photothèque
- Extensions d'images

"""

import os # Importation du module système pour la gestion des dossiers et des chemins de fichiers

# --- CONFIGURATION DU SCRIPT ---

# Chemins des fichiers et des dossiers

dossier_racine = ""  # Dossier source d'exposition à traiter
col_hash = 'MD5_HASH'   # Identifiant de l'empreinte numérique (MD5_HASH) dans notre cas
csv_file = ""  # Fichier des exports Droid 
dossier_elimination = ""        # Dossier des éliminables

# Reférentiel arborescence

excel_ref = "Referentiel_EXPOS_ARBO_2024.xlsx"  # Définition du nom du fichier Excel contenant l'arborescence cible
nom_onglet = "TYPOLOGIE MOTS CLES"  # Nom précis de l'onglet à lire à l'intérieur du fichier Excel
nom_colonne = "Arborescence + mots clés"  # Nom de la colonne de l'Excel qui contient les structures et mots-clés
destination = ""  # Chemin du dossier de sortie où l'arborescence finale sera créée

# Liste globale des extensions d'images pour le filet de sécurité
extensions_images = ['.jpg', '.jpeg', '.png', '.tif', '.tiff']

# Extensions à éliminer (conforme au script de tri et le script de décompréssion)

EXTENSIONS_ARCHIVES = {'.zip', '.traite', '.rar', '.7z', '.tar', '.gz'}

# Extensions de fichiers système, temporaires et de polices
# '.tmp', '.bak', '.wbk' = Fichiers de sauvegarde
# '.ini', '.inf', '.dat' = Fichiers de configuration
# '.db', '.bin' = Fichiers caché
# '.aae', '.pc3', '.url' = Fichiers issu de recherche internet
# '.ttf', '.otf', '.ttc' = Polices

EXTENSIONS_SYSTEME_TP = {'.tmp', '.bak', '.wbk','.ini', '.inf', '.dat', '.db', '.bin', '.aae', '.pc3', '.url', '.ttf', '.otf', '.ttc','.abr','.bat','.css','.hqx','.bup','.crt','.cab'}

# Configuration de logs

dossier_logs = "logs"


# Dictionnaire de synonymes métiers

# Permet d'associer des variantes du vrac aux concepts officiels du fichier "excel_ref"
# Le but de ce dictionnaire est de servir de "traducteur sémantique". Les utilisateurs écrivent souvent des noms de fichiers
# avec des abréviations ou du jargon métier (ex: "pub", "sceno", "facture"). Ce dictionnaire permet au script de comprendre
# que si un fichier contient mot "pub", son vrai concept de rangement dans l'arborescence officielle est "publicite" par exemple.
 
SYNONYMES_METIERS = {

   "prets": ["prêteur", "prêteuse", "prêt", "emprunt", "dosssier de prêt", "dossier de travail", "lettre de prêt", "lettre prêteur", "convention", "liste des prêts", "fiche de prêt", "recolement", "oeuvres pretables", "lettre preteur", "prêtable", "oeuvre pretable"],
   "commissaire": ["commissariat", "curator", "curatorial", "comite scientifique"],
   "presse": ["dp", "dossier de presse", "revue presse", "press review", "legende cp", "legende dp", "communiqué de press", "press realease", "cp", "rc", "point presse", "visite presse"],
   "scenographie": ["scéno", "plan masse", "maquette", "ambiance", "museo", "museographie", "scénographe", "chantier", "implantation", "installation", "montage", "élévation", "demontage", "accrochage", "en travaux", "soclage", "socle", "vitrine", "lumiere", "eclairage"],
   "publicite": ["pub"],
   "eclairage": ["lumineuse"],
   "carton d'invitation": ["carton"],
   "catalogue": ["argu", "argumentaire", "argus", "catalogue", "chronologie catalogue", "jaquette", "légende", "ouvrage", "roman"],
   "retombées presse": ["insertion publicitaire", "advertising insertion"],
   "cartel": ["cartl", "carel"],
   "constats d'etat": ["condition report"],
   "promotion": ["medias", "media plan", "plan media", "couv", "rs", "campagne photo", "reseaux sociaux", "formulaire de demande de communication", "demande de communication", "newsletter", "contrat photo", "journal", "brochure", "flyer", "com", "comm", "abribus", "etiquette", "etiquettes", "communique"],
   "restauration": ["bon d'accord", "hygiene"],
   "bon de dépôt": ["dépôt"],
   "securite": ["alarme"],
   "inauguration": ["cocktail", "guest list", "invitation", "vernissage", "photographs"],
   "iconographie": ["crédit", "légendes photos", "bon pour accord", "photo reportage", "icono", "reportage photo"],
   "signaletique": ["panneau", "bache", "texte de salle", "texte salle", "legende", "titre de section", "signaletique", "SIGNALÉTIQUE", "signal", "etique", "conception graphique signalétique"],
   "assurance": ["constat d'etat", "valeur d'assurance", "police d'assurance", "sinistre", "police d'assurance"],
   "transport": ["convoyage", "convoyeur", "caisse de transport", "retour", "enlèvement", "emballage", "transport d'oeuvre", "transport international", "arrivée des oeuvres", "envoi des oeuvres", "transitaire", "livraison", "transporteur", "clou-a-clou"],
   "juridique": ["contrat", "accord", "convention", "avenant", "autorisation", "droits", "copyright", "droits d'auteur", "autorisation imi", "imi", "lettre", "signee", "signe"],
   "budget": ["facture", "devis", "depense", "cout", "financier", "compta", "subvention", "dépenses", "recettes", "bilan", "quantitatif"],
   "mecenat partenariat": ["sponsorship", "sponsor", "liste marque", "sponsoring", "marque", "dm", "dossier de mécénat", "fondation", "listing contacts", "remerciements", "mécanat", "collaboration", "attention", "mécène", "partenaire", "donateur", "donation", "contrat de partenariat", "accords", "mécénat", "solicitation", "sollicitation"],
   "graphisme": ["affiche", "visuel", "logo", "charte", "flyer", "identite visuelle", "banniere", "kakemono", "poster", "collage", "conception graphique"],
   "reportages photograpiques": ["reportage", "legende"],
   "pedagogie_mediation": ["infos", "présentation auteur", "présentation artiste"],
   "projet_scientifiques": ["chronologie", "bio", "introprojet", "proposition"],
   "action_culturelle": ["visite guidee", "atelier", "médiation", "scolaire", "conferences", "colloque", "évènementiel", "publics", "visiteurs", "audioguide"],
   "programmation": ["planning", "retroplanning", "calendrier", "échéancier", "reunion", "cr", "compte-rendu", "ordre du jour", "odj"]
}


# Dictionnaire de code de nommage des reproductions d'oeuvres de la photothèque 

# Ce dictionnaire permet d'identifier les code définit par la "Photothèque" dans le but d'identifier les services ou les entités des Arts Décoratifs dans les vues.
# Il identifie, analyse et récupère les vues ou images avant leur transfert dans l'arborescence fichiers.
# Le but étant d'associer un code  de lettre strict (A, J, P, B, D, C, M) trouvé au début du nom d'une image au sous-dossier 04_4_Iconographie de l'arborescence.
 
CODE_NOMMAGE_PHOTOTHEQUE = {
   "A": "iconographie",  # "A" : MAD - oeuvre "Arts décoratifs" (tous les départements)
   "J": "iconographie",  # "J" : Jouets
   "PH": "iconographie", # "PH": Photo - Divers
   "P": "iconographie",  # "P" : PUB- département Publicité
   "B": "iconographie",  # "B" : BIBLI -oeuvres de la bibliothèque du MAD
   "D": "iconographie",  # "D" : DOC - centre de documentation - documents
   "C": "iconographie",  # "C" : CAM - oeuvre du Musée Nissim de Camondo
   "M": "iconographie",  # "M" : MODE
}

# ----------- 
# Les scores
# -----------

# pour la prise en compte de la longueur minimum d'un mot 
LONGUEUR_MIN_MOT = 3 


# Critères de classement des scores

SCORE_MOT_CLE_TITRE = 90     # Mot clé présent dans le nom du fichier
SCORE_MOT_CLE_CONTENU = 30   # Mot clé présent dans le contenu
SCORE_MAITRE_TITRE = 60      # Concept métier présent dans le nom du fichier
SCORE_SYNONYME = 50          # Synonyme métier présent dans le nom du fichier
SCORE_SYNONYME_COMPOSE = 45  # Présence de tous les mots composé d'un synonyme (dans le contenu) 


