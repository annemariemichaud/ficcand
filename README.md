# ficcand

## Objectifs du programme : 
L'utilisation du batch ficcand dans Gaia permet de retenir une liste de personnes.
Ce batch nécessite de connaître les numens des personnes à convoquer.
Lorsque ces numens ne sont pas connus, le programme ficcand permet de retrouver les numens des personnes à convoquer.

## Principe de fonctionnement
Ce programme compare une liste de personnes à convoquer au format xlsx à la liste des personnels d'une académie extraite de Gaia par une requête BI4.
Il identifie chaque personne à convoquer en cherchant les correspondances entre :
* le nom d'usage (obligatoire)
* le prénom (obligatoire)
* la discipline d'exercice (optionnel)
* le rne du lieu d'affectation (optionnel)
* la ville du lieu d'affectation (optionnel).

Le programme produit 4 types de fichiers texte en sortie :

* le fichier erreur.txt qui enregistre toutes les erreurs d'exécution (par exemple une liste sans prénom)
* le fichier ficCandidat.unl qui est le fichier à uploader sur l'interface Gaia EBG
* le fichier nontrouve.txt qui liste les personnes qui n'ont pas été identifiées par le programme 
* le fichier verification.txt qui permet de vérifier les disciplines d'exercice des personnes convoquées. 

## Utilisation

* 1ère étape : extraire la liste des numens des personnels de l'académie via requête BI4. La liste des champs à utiliser est donnée dans le fichier tous_numen_simple_modele.
L'ordre des colonnes doit impérativement être respecté. La requête doit être filtrée dans BI4 via le code position : C101, C102, C103 pour les personnes en activité. 

* 2ème étape : récupérer l'exécutable ficcand.exe qui se trouve dans le dossier build (à venir). 

* 3ème étape : récupérer la liste des personnes à convoquer en utilisant le fichier modèle : liste_modèle. Attention, l'odre des colonnes doit être respectée.

* 4ème étape : tous ces fichiers doivent se trouver dans le même répertoire. Il suffit de lancer ficcand.exe et de patienter. Les 4 fichiers textes de sortie vont être produits.

* 5ème étape : uploader le fichier ficCandidat.unl sur la plateforme ebg.

## Limites d'utilisation 

* Le programme peut retenir par erreur des homonymes. C'est particulièrement le cas lorsque seul les nom et prénom sont indiqués. Il faut alors impérativement lire le fichier verification.txt.
Par exemple, si seuls des enseignants du second degré doivent être convoqués, la mention "sans spécialité" doit vous alerter. Il s'agit alors soit de personnels administratifs, soit d'enseignants du 1er degré.

* Le programme ne peut pas trouver des personnes qui ne se trouvent pas dans le fichier tous_numen_simple. Il faudra donc probablement faire différents essais de filtres avec différents codes position.

## Licence

Ce programme est fourni sous la licence GNU GENERAL PUBLIC LICENSE V3. 








