import openpyxl
import unicodedata
import re
from data import numens as numens

def nettoyer_fichier(valeur):
    if valeur is not None:
        valeur=unicodedata.normalize('NFKD',valeur).encode('ASCII', 'ignore').decode('ASCII')
        valeur=valeur.strip()
        valeur=valeur.upper()
        if valeur.find('-'):
            valeur=valeur.replace('-',' ')
    return valeur

wb = openpyxl.load_workbook('liste.xlsx')

onglet = wb.active

liste_titre_colonne = []

personnes = []

#ecrire les resultas dans un fichier
ficCandidatFile = open('ficCandidat.unl', 'w')
erreurFile = open('erreur.txt','w')
verificationFile = open('verification.txt','w')
#recuperation donnees liste à traiter
for i in range(1,onglet.max_column+1):
    liste_titre_colonne.append(onglet.cell(row=1,column=i).value)

for i in range(2,onglet.max_row+1):
    personne = {}
    for j in range(0,onglet.max_column):
        personne[liste_titre_colonne[j]]=onglet.cell(row=i,column=j+1).value
    personnes.append(personne)


#nettoyer le fichier de donnees
for ligne in personnes:
    for cle in ligne.keys():
        if cle in ('Nom','Prenom','Discipline','Ville','RNE'):
            ligne[cle]=nettoyer_fichier(ligne[cle])

global trouve
trouve=0
resultat_fort = []
resultat_faible = []
#chercher les personnes

    
for ligne in personnes:

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
