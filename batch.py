import pandas as pd
import unicodedata
import sys
import datetime

def nettoyer_cedex(valeur):
    if 'CEDEX' in valeur:
        valeur=valeur.replace('CEDEX','')
        valeur=valeur.strip()
    return valeur

def nettoyer_fichier(valeur):
    #suppression des caractères accentués
    valeur= unicodedata.normalize('NFKD', valeur)
    valeur = u"".join([c for c in valeur if not unicodedata.combining(c)])
    #les valeurs saisies dans le fichier liste doit être nettoyées et normalisées pour pouvoir être comparées
    if valeur is not None:
        #suppression des espaces avant et après la valeur
        valeur=valeur.strip()
        #passage de toutes les valeurs en caractères minuscules
        valeur=valeur.upper()
        #remplacement des - en espace
        if valeur.find('-'):
            valeur=valeur.replace('-',' ')
        if valeur.find('\''):
            valeur=valeur.replace('\'',' ')
    return valeur

def handleError(message):
    erreurFile.write(datetime.datetime.now().strftime("%d-%m-%Y %H:%M"))
    erreurFile.write('\n')
    erreurFile.write(message)
    erreurFile.close()
    sys.exit()
    
#écrire les erreurs dans un fichier erreur.txt
erreurFile = open('erreur.txt', 'w')
    
#ouverture du fichier tous numens
try:
    numens = pd.read_excel('tous_numen_simple.xlsx')
except FileNotFoundError as fnf_error:
    handleError('le fichier tous_numen_simple est introuvable \n')
    
#transformation en dataframe
df_numens=pd.DataFrame(numens)
#renommage des colonnes
try:
    df_numens.rename(columns={'Numen':'numen',"Nom d'usage":'nom_usage',"Nom patronymique":'nom_patronymique',"Prénom":"prenom","Discipline d'exercice Libellé":"discipline_exercice","Code du lieu d'affectation":"rne","Ville du lieu":"ville"},inplace=True)

    #suppression des lignes avec nom ou prénom manquants
    df_numens.dropna(axis=0,subset=['nom_usage','prenom','nom_patronymique'],inplace=True)
    #completion des valeurs na
    df_numens.discipline_exercice.fillna(value="NON RENSEIGNE",inplace=True)
    df_numens.rne.fillna(value="NON RENSEIGNE",inplace=True)
    df_numens.ville.fillna(value="NON RENSEIGNE",inplace=True)


    #nettoyage des colonnes : suppression des cedex, normalisation des noms et prenoms
    df_numens.loc[:,'ville']=df_numens.apply(lambda row:nettoyer_cedex(row['ville']),axis=1)
    df_numens.loc[:,'nom_usage']=df_numens.apply(lambda row:nettoyer_fichier(row['nom_usage']),axis=1)
    df_numens.loc[:,'nom_patronymique']=df_numens.apply(lambda row:nettoyer_fichier(row['nom_patronymique']),axis=1)
    df_numens.loc[:,'prenom']=df_numens.apply(lambda row:nettoyer_fichier(row['prenom']),axis=1)
except AttributeError:
    handleError('le nom des colonnes de tous_numen_simple ne correspond pas, reprendre les mêmes noms que dans le fichier modèle \n')

#ouverture du fichier liste à convoquer
try:
    convocations = pd.read_excel('liste.xlsx')
except FileNotFoundError as fnf_error:
    handleError('le fichier liste.xlsx est introuvable \n')
    
#transformation en dataframe
df_convocations=pd.DataFrame(convocations)
try:
    #renommage des colonnes
    df_convocations.rename(columns={'Nom':'nom',"Prenom":'prenom',"RNE":'rne',"Discipline":"discipline","Ville":"ville","Dispositif":"dispositif","Module":"module","Groupe":"groupe"},inplace=True)
    #vérification des colonnes optionnelles complétées
    discipline = True
    rne = True
    ville = True
    if df_convocations.discipline.isna().sum()==df_convocations.shape[0]:
        discipline=False
    if df_convocations.rne.isna().sum()==df_convocations.shape[0]:
        rne=False
    if df_convocations.ville.isna().sum()==df_convocations.shape[0]:
        ville=False
   #completer les valeurs non renseignées
    df_convocations.discipline.fillna(value="NON RENSEIGNE",inplace=True)
    df_convocations.rne.fillna(value="NON RENSEIGNE",inplace=True)
    df_convocations.ville.fillna(value="NON RENSEIGNE",inplace=True)
    #verification que toutes les lignes comportent bien un nom et un prénom
    verif_nom = df_convocations.nom.isna()
    verif_prenom = df_convocations.prenom.isna()
    df_non_traitees = df_convocations[verif_nom | verif_prenom]
    df_non_traitees.prenom.fillna(value="NON RENSEIGNE",inplace=True)
    df_non_traitees.nom.fillna(value="NON RENSEIGNE",inplace=True)
    #vérification des colonnes obligatoires dispositif, module et groupe
    if df_convocations.groupe.isna().sum()>0:
        handleError ('le numéro de groupe est obligatoire')
    if df_convocations.dispositif.isna().sum()>0:
        handleError('le numéro de dispositif est obligatoire')
    if df_convocations.module.isna().sum()>0:
        handleError('le numéro de module est obligatoire')
except AttributeError:
    handleError('le nom des colonnes de liste.xlsx ne correspond pas, reprendre les mêmes noms que dans le fichier modèle \n')

#ajouter les non traitees à une liste de dictionnaire
liste_non_traitees=df_non_traitees.to_dict('records')

#ne sont traitées que celles avec nom et prénom
df_traitees = df_convocations[~verif_nom & ~verif_prenom]

#nettoyage nom et prénom
df_traitees.loc[:,'nom']=df_traitees.apply(lambda row:nettoyer_fichier(row['nom']),axis=1)
df_traitees.loc[:,'prenom']=df_traitees.apply(lambda row:nettoyer_fichier(row['prenom']),axis=1)

#nettoyage des cedex 
df_traitees.loc[:, 'ville']=df_traitees.apply(lambda row:nettoyer_cedex(row['ville']),axis=1)

#transformation des df en liste de dictionnaires
liste_convocations=df_traitees.to_dict('records')
liste_numens = df_numens.to_dict('records')

liste_titre_colonne = []
personnes = []

#écrire les résultats dans un fichier ficCandidat.unl
ficCandidatFile = open('ficCandidat.unl', 'w')
#écrire les erreurs (personnes non trouvées et/ou personnes avec plusieurs identifications possibles
nontrouve_File = open('nontrouve.txt','w')
#écrire la liste des disciplines pour vérification (très important) pour éviter les erreurs d'identification
verificationFile = open('verification.txt','w')

global trouve
trouve=0
#résultat si tous les champs de la liste de personnes à convoquer concordent avec la liste des numens
resultat_fort = []

#boucle chercher les personnes   
for convocation in liste_convocations:

    #comparaison
    for numen in liste_numens:
            if (convocation['nom'] == numen['nom_usage'] or convocation['nom'] == numen['nom_patronymique']) and  convocation['prenom'] == numen['prenom']:
                if rne or ville or discipline:
                    if convocation['discipline'] == numen['discipline_exercice'] or convocation['rne'] == numen['rne'] or convocation['ville'] == numen['ville']:
                        resultat_fort.append(numen)
                else:
                     resultat_fort.append(numen)
           
        
            
    #écriture des résultats dans les fichiers non trouvés, vérification et ficCandidat.unl
    if len(resultat_fort)==1:
        for numen in resultat_fort:
            #ajout du numen dans ficCandidat
            candidature=f"{numen['numen']}|{str(convocation['dispositif'])}|{str(convocation['module'])}|{str(convocation['groupe']).rjust(2,'0')}|R|\n"
            ficCandidatFile.write(candidature)
            #ajout du numen dans verification
            verification=f"{numen['numen']} {numen['nom_usage']} {numen['prenom']} {numen['discipline_exercice']} \n"
            verificationFile.write(verification)

#cas ou plusieurs candidats
    if len(resultat_fort)>1 :
        ficCandidatFile.write(f"Plusieurs candidats possibles pour {convocation['nom']} {convocation['prenom']} \n")

        for numen in resultat_fort:
            candidature = f"{numen['numen']}|{str(convocation['dispositif'])}|{str(convocation['module'])}|{str(convocation['groupe']).rjust(2,'0')}|R| {numen['nom_usage']} {numen['prenom']} {numen['discipline']} {numen['rne']} {numen['ville']} \n"
            nontrouve_File.write(candidature)

#cas sans convocation trouvée
    if len(resultat_fort)==0:
        nontrouve_File.write('Candidat non trouvé \n')
        candidature = f"{convocation['nom']} {convocation['prenom']} {str(convocation ['discipline'])} {convocation['rne']} {convocation['ville']} \n"
        nontrouve_File.write(candidature)

    resultat_fort = []
    resultat_faible = []

#cas des convocations sans nom et prénom
if len(liste_non_traitees)>0:
    nontrouve_File = open('nontrouve.txt','a')
    for convocation in liste_non_traitees:
        if convocation['nom'] =="NON RENSEIGNE":
            nontrouve_File.write('Candidat non trouvé car pas de nom indiqué \n')
            candidature = f"{convocation['nom']} {convocation['prenom']} \n"
            nontrouve_File.write(candidature)
        if convocation['prenom']=="NON RENSEIGNE":
            nontrouve_File.write('Candidat non trouvé car pas de prénom indiqué \n')
            candidature = f"{convocation['nom']} {convocation['prenom']} \n"
            nontrouve_File.write(candidature)


ficCandidatFile.close()
nontrouve_File.close()
verificationFile.close()
erreurFile.close()
