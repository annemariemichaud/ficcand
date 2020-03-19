import pandas as pd
import unicodedata


def nettoyer_cedex(valeur):
    valeur=valeur.strip()
    if 'CEDEX'in valeur:
        valeur=valeur[:-6]
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

#ouverture du fichier tous numens
numens = pd.read_excel('tous_numen_simple.xlsx')
#transformation en dataframe
df_numens=pd.DataFrame(numens)
#renommage des colonnes
df_numens.rename(columns={'Numen':'numen',"Nom d'usage":'nom_usage',"Nom patronymique":'nom_patronymique',"Prénom":"prenom","Discipline d'exercice Libellé":"discipline_exercice","Code du lieu d'affectation":"rne","Ville du lieu":"ville"},inplace=True)
#suppression des lignes avec nom ou prénom manquants
df_numens.dropna(axis=0,subset=['nom_usage','prenom','nom_patronymique'],inplace=True)
#completion des valeurs na
df_numens.discipline_exercice.fillna(value="NON RENSEIGNE",inplace=True)
df_numens.rne.fillna(value="NON RENSEIGNE",inplace=True)
df_numens.ville.fillna(value="NON RENSEIGNE",inplace=True)
#nettoyage des colonnes : suppression des cedex, normalisation des noms et prenoms
df_numens['ville']=df_numens.apply(lambda row:nettoyer_cedex(row['ville']),axis=1)
df_numens['nom_usage']=df_numens.apply(lambda row:nettoyer_fichier(row['nom_usage']),axis=1)
df_numens['nom_patronymique']=df_numens.apply(lambda row:nettoyer_fichier(row['nom_patronymique']),axis=1)
df_numens['prenom']=df_numens.apply(lambda row:nettoyer_fichier(row['prenom']),axis=1)


#ouverture du fichier liste à convoquer
convocations = pd.read_excel('liste.xlsx')
#transformation en dataframe
df_convocations=pd.DataFrame(convocations)
#renommage des colonnes
df_convocations.rename(columns={'Nom':'nom',"Prenom":'prenom',"RNE":'rne',"Discipline":"discipline","Ville":"ville","Dispositif":"dispositif","Module":"module","Groupe":"groupe"},inplace=True)

#verification que toutes les lignes comportent bien un nom et un prénom
verif_nom = df_convocations.nom.isna()
verif_prenom = df_convocations.prenom.isna()
df_non_traitees = df_convocations[verif_nom | verif_prenom]

#TODO identifier les non traitees
#TODO ajouter les non traitees au fichier non trouvés

#ne sont traitées que celles avec nom et prénom
df_traitees = df_convocations[~verif_nom & ~verif_prenom]

#nettoyage nom et prénom
df_traitees['nom']=df_traitees.apply(lambda row:nettoyer_fichier(row['nom']),axis=1)
df_traitees['prenom']=df_traitees.apply(lambda row:nettoyer_fichier(row['prenom']),axis=1)

#nettoyage des cedex 
df_traitees['ville']=df_traitees.apply(lambda row:nettoyer_cedex(row['ville']),axis=1)

#TODO transformation des df en liste de dictionnaires

liste_titre_colonne = []
personnes = []

#écrire les résultats dans un fichier ficCandidat.unl
ficCandidatFile = open('ficCandidat.unl', 'w')
#écrire les erreurs (personnes non trouvées et/ou personnes avec plusieurs identifications possibles
erreurFile = open('erreur.txt','w')
#écrire la liste des disciplines pour vérification (très important) pour éviter les erreurs d'identification
verificationFile = open('verification.txt','w')

#récupération des titres des colonnes du fichier liste
for i in range(1,onglet.max_column+1):
    liste_titre_colonne.append(onglet.cell(row=1,column=i).value)
#récupération de la liste des personnes à convoquer
for i in range(2,onglet.max_row+1):
    personne = {}
    for j in range(0,onglet.max_column):
        personne[liste_titre_colonne[j]]=onglet.cell(row=i,column=j+1).value
    personnes.append(personne)


#nettoyer le fichier de données
for ligne in personnes:
    for cle in ligne.keys():
        if cle in ('Nom','Prenom','Discipline','Ville','RNE'):
            ligne[cle]=nettoyer_fichier(ligne[cle])


global trouve
trouve=0
#résultat si tous les champs de la liste de personnes à convoquer concordent avec la liste des numens
resultat_fort = []
##résultat si certains champs de la liste de personnes à convoquer concordent avec la liste des numens
resultat_faible = []

#chercher les personnes   
for ligne in personnes:

    #selon deux listes resultats faibles ou forts
    for numen in numens:
        if (numen['Nom usage'] is not None and numen['Nom patronymique'] is not None) and numen['Prenom'] is not None:
            if (ligne['Nom'] == numen['Nom usage'] or ligne['Nom'] == numen['Nom patronymique']) and  ligne['Prenom'] == numen['Prenom']:
                if numen['Discipline'] is not None and ligne['Discipline'] is not None :
                    if ligne['Discipline'] in numen['Discipline']:
                        resultat_fort.append(numen)
                if ligne['RNE'] is not None and numen['RNE'] is not None:
                    if ligne['RNE'] == numen['RNE']and numen not in resultat_fort:
                        resultat_fort.append(numen)
                if ligne['Ville'] is not None and numen['Ville'] is not None:
                    if ligne['Ville'] in numen['Ville'] and numen not in resultat_fort:
                        resultat_fort.append(numen)
                if len(resultat_fort)==0:
                    resultat_faible.append(numen)

            if (ligne['Nom'] in numen['Nom usage'] or ligne['Nom'] in numen['Nom patronymique'] or numen['Nom usage'] in ligne['Nom'] or numen['Nom patronymique'] in ligne['Nom'] ) and  (ligne['Prenom'] in numen['Prenom'] or numen['Prenom'] in ligne['Prenom']):
                if numen['Discipline'] is not None and ligne['Discipline'] is not None :
                    if ligne['Discipline'] in numen['Discipline']:
                        resultat_faible.append(numen)
                if ligne['RNE'] is not None and numen['RNE'] is not None:
                    if ligne['RNE'] == numen['RNE']and numen not in resultat_faible:
                        resultat_faible.append(numen)
                if ligne['Ville'] is not None and numen['Ville'] is not None:
                    if ligne['Ville'] in numen['Ville'] and numen not in resultat_faible:
                        resultat_faible.append(numen)
                if len(resultat_fort)==0 and len(resultat_faible)==0:
                    resultat_faible.append(numen)
                    
    #écriture des résultats dans les fichiers erreurs, vérification et ficCandidat.unl
    if len(resultat_fort)==1:
        for numen in resultat_fort:
            candidature = numen['Numen'] + '|' + str(ligne['Dispositif']) + '|' + str(ligne['Module']) + '|' + str(ligne['Groupe']).rjust(2,'0') + '|R|\n'
            if numen['Discipline'] is None:
                numen['Discipline'] = 'sans specialite'
            verification = numen['Numen'] + ' ' + numen['Nom usage'] + ' ' + ' ' + numen['Prenom'] + ' ' +  numen['Discipline'] + '\n'
            verificationFile.write(verification)
            ficCandidatFile.write(candidature)

    if len(resultat_faible)==1 and len(resultat_fort)==0:
        for numen in resultat_faible:
            candidature = numen['Numen'] + '|' + str(ligne['Dispositif']) + '|' + str(ligne['Module']) + '|' + str(ligne['Groupe']).rjust(2,'0') + '|R|\n'
            if numen['Discipline'] is None:
                numen['Discipline'] = 'sans specialite'
            verification = numen['Numen'] + ' ' + numen['Nom usage'] + ' ' + ' ' + numen['Prenom'] + ' ' +  numen['Discipline'] + '\n'
            verificationFile.write(verification)
            ficCandidatFile.write(candidature)

    if len(resultat_fort)>1 :
        erreurFile.write('Plusieurs candidats possibles pour ' + ligne['Nom'] + ' ' + ligne['Prenom'] + '\n')

        for numen in resultat_fort:
            candidature = numen['Numen'] + '|' + str(ligne['Dispositif']) + '|'
            candidature = candidature + str(ligne['Module']) + '|'
            if numen['Discipline'] is None:
                numen['Discipline'] = 'sans specialite'
            candidature = candidature + str(ligne['Groupe']).rjust(2,'0') + '|R|' + ' ' + numen['Nom usage'] + ' ' + numen['Prenom'] + ' ' + numen['Discipline'] + ' ' + numen['RNE'] + ' ' + numen['Ville'] + '\n'
            erreurFile.write(candidature)

    if len(resultat_faible)>1 and len(resultat_fort)==0:
        erreurFile.write('Plusieurs candidats possibles pour ' + ligne['Nom'] + ' ' + ligne['Prenom'] + '\n')

        for numen in resultat_faible:
            candidature = numen['Numen'] + '|' + str(ligne['Dispositif']) + '|'
            candidature = candidature + str(ligne['Module']) + '|'
            if numen['Discipline'] is None:
                numen['Discipline'] = 'sans specialite'
            candidature = candidature + str(ligne['Groupe']).rjust(2,'0') + '|R|' + ' ' + numen['Nom usage'] + ' ' + numen['Prenom'] + ' ' + numen['Discipline'] + ' ' + numen['RNE'] + ' ' + numen['Ville'] + '\n'
            erreurFile.write(candidature)

    if len(resultat_fort)==0 and len(resultat_faible)==0:
        erreurFile.write('Candidat non trouve \n')
        candidature = ligne['Nom'] + ' ' + ligne['Prenom']
        if ligne ['Discipline'] is None :
            ligne['Discipline'] = 'sans specialite'
        candidature = candidature + ' ' + ligne ['Discipline']
        if ligne['RNE'] is  None:
            ligne['RNE'] = 'Sans RNE'
        candidature = candidature    + ' ' + ligne['RNE']
        if ligne['Ville'] is None:
            ligne['Ville'] = 'sans ville'
        candidature = candidature + ' ' + ligne['Ville']
        candidature = candidature + '\n'
        erreurFile.write(candidature)

    resultat_fort = []
    resultat_faible = []




ficCandidatFile.close()
erreurFile.close()
verificationFile.close()
