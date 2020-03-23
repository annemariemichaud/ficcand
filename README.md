# ficcand

## Objectifs du programme : 
L'utilisation du batch ficcand dans Gaia permet de retenir une liste de personnes.
Ce batch nécessite de connaître les numens de personnes à convoquer.
Lorsque ces numens ne sont pas connus, le programme ficcand permet de retrouver les numens des personnes à convoquer.

## Principe de fonctionnement
Ce programme compare une liste de personnes à convoquer au format xlsx à la liste des personnels d'une académie extraite de Gaia par une requête BI4.
Il identifie chaque personne à convoquer en cherchant les correspondances entre :
* le nom d'usage (obligatoire)
* le nom patronymique (obligatoire)
* la discipline d'exercice (optionnel)
* le rne du lieu d'affectation (optionnel)
* la ville du lieu d'affectation (optionnel).

Le programme produit 4 types de fichiers text en sortie :

* le fichier erreur.txt qui enregistre toutes les erreurs d'exécution (par exemple une liste sans prénom)
* le fichier ficCandidat.unl qui est le fichier à uploader sur l'interface Gaia EBG
* le fichier nontrouve.txt qui liste les personnes qui n'ont pas été identifiées par le programme 
* le fichier verification.txt qui permet de vérifier les disciplines d'exercice des personnes convoquées. 






