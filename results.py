#

import csv
import sys

###############################################################################
# Key extraction
###############################################################################

def ListCities(data):
    cities = set()
    for elt in data:
        cities.add(elt.city)
    return list(cities)


def ListCategories(data):
    categories = set()
    for elt in data:
        categories.add(elt.category)
    return list(categories)

###############################################################################
# Sorting and tallying total times
###############################################################################

def SortByTot(data):
    return sorted(data, key=lambda x: x.tot)


def TotalTimeHS(data, cutoff):
    tot = 0
    if len(data) < cutoff:
        return -1
    for elt in data[0:cutoff]:
        tot += elt.tot
    return tot

###############################################################################
# Sorting by bib
###############################################################################

def SortByBib(data):
    return sorted(data, key=lambda x: x.bib)

###############################################################################
# Data selection according to criteria
###############################################################################

def SelectByCity(data, city):
    return [elt for elt in data if elt.city == city]


def SelectByGroup(data, group):
    return [elt for elt in data if elt.group == group]


def SelectByCityAndGroup(data, city, group):
    return [elt for elt in data if elt.city == city and elt.group == group]


def SelectBySex(data, sex):
    assert sex in ('M', 'F')
    return [elt for elt in data if elt.sex == sex]


def SelectByCatAndSex(data, cat, sex):
    assert sex in ('M', 'F')
    return [elt for elt in data if elt.sex == sex and elt.category == cat]

###############################################################################
# Time stamps management
###############################################################################

def TimeNA(timestamp):
    if not timestamp:
        return True
    return str(timestamp)[0] not in [
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']


def HSToTime(hundreds):
    minutes=int((hundreds/(100*60))%60)
    seconds=int((hundreds/100)%60)
    hundreds = int(hundreds%100)
    if minutes:
        return "{0}:{1:02d}.{2:02d}".format(minutes, seconds, hundreds)
    else:
        return "{0:02d}.{1:02d}".format(seconds, hundreds)


def TimeToHS(timestamp):
    minutes = 0
    if ':' in timestamp:
        values = timestamp.split(':')
        assert len(values) == 2, "timestamp={0}".format(timestamp)
        minutes = int(values[0])
        timestamp = values[1]
    values = timestamp.split('.')
    assert len(values) == 2, "timestamp={0}".format(timestamp)
    return minutes*60*100 + int(values[0])*100 + int(values[1])

###############################################################################
# Class definitions
###############################################################################

class ResultEntry:
    def __init__(self, bib, name, sex, category, yob, city, group, tot):
        assert name and bib and sex and category
        assert yob and city and group and tot
        self.bib = int(bib)
        self.name = name
        self.sex = sex
        self.category = category
        self.yob = yob
        self.city = city
        self.group = group
        self.tot = tot

    def Print(self, rank, print_header=True):
        if rank <= 1 and print_header:
            print("Rang Dos Nom                          Sexe "
                  "Cat    Annee Ville                Grp       Tot\n")
        if TimeNA(self.tot):
            tot = self.tot
        else:
            tot = HSToTime(self.tot)
        print("{0:<4d} {1:<3d} {2:<28} {3:<4} {4:<6} "
              "{5:<5} {6:<20} {7:<4}  {8:>7}".format(
            rank, self.bib, self.name, self.sex, self.category,
            self.yob, self.city, self.group, tot))


class ResultSet:
    def __init__(self, entries, length, tot):
        assert entries and len(entries) and tot
        assert length
        assert len(entries) >= length
        self.entries = entries[0:length]
        self.tot = tot

    def Print(self, print_header):
        index = 1
        for elt in self.entries:
            elt.Print(index, print_header)
            index += 1
        print("                                                        "
              "                     -------------")
        print("                                                        "
              "                     "
              "TOT:  {0:<7}".format(HSToTime(self.tot)))
        print()


def ReadInput(file, only_dnf=False):
    data = []
    fields = ['Bib', 'Name', 'Sex', 'Category',
              'Yob', 'City', 'Group', 'M1', 'M2', 'Tot']
    with open(file, 'r') as handle:
        csv_reader = csv.DictReader(handle, delimiter='\t', fieldnames=fields)
        for row in csv_reader:
            if only_dnf:
                if TimeNA(row['Tot']):
                    # Find which lap the disqualification happened
                    lap = ' ?'
                    if TimeNA(row['M1']):
                        lap = ' M1'
                    if TimeNA(row['M2']):
                        lap = ' M2'
                    data.append(ResultEntry(row['Bib'], row['Name'], row['Sex'],
                                            row['Category'], row['Yob'], row['City'],
                                            row['Group'], row['Tot']+lap))
                else:
                    continue
            else:
                if TimeNA(row['Tot']):
                    print("⚠️  {0}, bib={1}. Total time: {2}, skipped".format(
                        row['Name'], row['Bib'], row['Tot']))
                    continue
                data.append(ResultEntry(row['Bib'], row['Name'], row['Sex'],
                                        row['Category'], row['Yob'], row['City'],
                                        row['Group'], TimeToHS(row['Tot'])))
    return data

###############################################################################
# Printing
######################################################################

def PrintResultEntry(entries):
    index = 1
    for elt in entries:
        elt.Print(index)
        index += 1

def PrintIndexedCutoff(data, cutoff):
    index = 1
    if len(data) == 0:
        print('*** VIDE\n')
    
    for elt in data:
        elt.Print(index)
        if index == cutoff and elt != data[-1]:
            print()
        index += 1

def PrintHeader(msg):
    padding = 89
    print('-'*padding)
    print(msg.center(padding))
    print('-'*padding)
        
###############################################################################
# Scratch M/F
###############################################################################

def Scratch(data, header, what, cutoff):
    PrintHeader(header)
    PrintIndexedCutoff(SortByTot(SelectBySex(data, what)), cutoff)


def ScratchF(data):
    Scratch(data, 'SCRATCH FEMMES', 'F', 3)


def ScratchM(data):
    Scratch(data, 'SCRATCH HOMMES', 'M', 3)


###############################################################################
# By cities and teams (FIXME)
###############################################################################

def ByCities(data):
    PrintHeader('PAR VILLES')
    cities = ListCities(data)
    cities_threshold = 3
    result_set = []
    for city in cities:
        # Select a city and sort all results
        sorted_select = SortByTot(SelectByCity(data, city))
        # For that city, compute the total time for the cities_threshold-th first entries
        total = TotalTimeHS(sorted_select, cities_threshold)
        if total != -1:
            result_set.append(ResultSet(sorted_select, cities_threshold, total))

    # Sort the selected cities by total time, ascending and print the results.
    sorted_final = SortByTot(result_set)
    print_header = True
    for elt in sorted_final:
        elt.Print(print_header)
        print_header = False


def ByPM(data):
    PrintHeader('PAR POLICES MUNICIPALES')
    group = 'PM'
    cities = ListCities(data)
    pm_threshold = 3
    result_set = []
    for city in cities:
        # Select a city/group and sort all results
        sorted_select = SortByTot(SelectByCityAndGroup(data, city, group))
        # For that city/group, compute the total time for the pm_threshold-th
        # first entries
        total = TotalTimeHS(sorted_select, pm_threshold)
        if total != -1:
            result_set.append(ResultSet(sorted_select, pm_threshold, total))

    # Sort the selected cities by total time, ascending and print the results.
    sorted_final = SortByTot(result_set)
    print_header = True
    for elt in sorted_final:
        elt.Print(print_header)
        print_header = False

###############################################################################
# By categories
###############################################################################

def ByCategories(data):
    categories = ListCategories(data)
    cutoff = 3
    for category in ['SENIOR', 'M1', 'M2', 'M3', 'M4', 'M5', 'V1', 'V2', 'V3', 'SNOW']:
        if not category in categories:
            continue
        selection = SortByTot(SelectByCatAndSex(data, category, 'M'))
        if not selection:
            continue
        PrintHeader('{0} HOMMES'.format(category))
        PrintIndexedCutoff(selection, cutoff)
        selection = SortByTot(SelectByCatAndSex(data, category, 'F'))
        if not selection:
            continue
        PrintHeader('{0} FEMMES'.format(category))
        PrintIndexedCutoff(selection, cutoff)

assert TimeNA('')
assert TimeNA('Dsq')
assert not TimeNA('12.22')
assert not TimeNA('1:12.22')
assert TimeToHS('13.30') == 1330
assert TimeToHS('1:23.65') == 8365
assert TimeToHS('2:13.30') == 13330
assert HSToTime(TimeToHS('13.30')) == '13.30'
assert HSToTime(TimeToHS('1:23.65')) == '1:23.65'
assert HSToTime(TimeToHS('2:13.30')) == '2:13.30'
assert HSToTime(TimeToHS('1:05.23')) == '1:05.23'
assert HSToTime(TimeToHS('1:05.03')) == '1:05.03'

data = ReadInput('/Users/apetitbianco/Downloads/EDITION.txt')
PrintHeader('Qualifies rentrant dans les classements')
PrintResultEntry(SortByBib(data))

dnf = ReadInput('/Users/apetitbianco/Downloads/EDITION.txt', only_dnf=True)
PrintHeader('DNF')
PrintResultEntry(SortByBib(dnf))

# ByCities(data)
# ByPM(data)
# ScratchM(data)
# ScratchF(data)
# ByCategories(data)
