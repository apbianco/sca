# Prend en entree un fichier CSV avec les colonnes suivantes:
#
#   Nom,Prenom,Ville,Annee,Sexe,Categorie,Age,Groupe
#
# Et sort un fichier sav avec le format suivant:
#
# EQU%d|Sex|Nom|Prenom|Annee||||||||||||Groupe|Ville|||Categorie|||||||||||||
#
# pour inportation directe dans ski ffs.
#
# Utilisation: a partir de l'invite de commandes. input.csv c'est le
# fichier Excel/Trix sauvegarde en .csv et /tmp/output.sav c'est le
# fichier a ingurgiter dans ski ffs.
#
# $ python3 championnat_pm.y 

import sys
import csv

class SAVEntry:
    def __init__(self, name, surname, sex, yob, group, city, category):
        assert name and surname and yob
        assert group and city and category
        if sex == 'H':
            sex = 'M'
        assert sex == 'M' or sex == 'F', "sex={0}".format(sex)
        self.name = name
        self.surname = surname
        self.sex = sex
        self.yob = yob
        self.group = group
        self.city = city
        self.category = category

    def Str(self, rank):
        return ('EQU{0}|{1}|{2}|{3}|{4}|||||||||||'
                '{5}|{6}||||{7}|||||||||||||'.format(
                    rank, self.sex, self.name, self.surname, self.yob,
                    self.city, self.group, self.category))

        
def GenerateData(input_filename, output_filename):
    data = []
    with open(input_filename, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for row in csv_reader:
            if not row["Groupe"]:
                print("Skipping {0} {1} ({2}) as non racer".format(
                    row["Nom"], row["Prénom"], row["Ville"]))
                continue
            data.append(SAVEntry(name=row["Nom"],
                                 surname=row["Prénom"],
                                 city=row["Ville"],
                                 yob=row["Année de Naissance"],
                                 sex=row["Sexe"],
                                category=row["Cat"],
                                 group=row["Groupe"]))
    index = 1
    with open(output_filename, mode='w') as handle:
        for elt in data:
            handle.write("{0}\n".format(elt.Str(index)))
            index += 1
        

data = GenerateData('input.csv', '/tmp/output.sav')
sys.exit(0)
