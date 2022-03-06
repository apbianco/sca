import sys

def ReadRace(filename):
  try:
    with open(filename, 'r') as f:
      return f.readlines()
  except:
    print('Can not read file {0}, exiting...'.format(filename))
    sys.exit(1)

def PrintLine(line):
  print('Name={0} {1}, Club={2}, Equipe={3}'.format(line[5], line[4], line[9], line[12]))

def ModifyRace(race):
  num_racer = 0
  new_race = []
  for line in race:
    split_line = line.split('|')
    # Identify the race header as a line with 30+ entries with CALEND as the 6th entry
    if len(split_line) > 30 and split_line[5] == 'CALEND':
      split_line[8] = split_line[8] + ' (Equipes)'
    # Identify a registration as a line with 40+ entries that starts with a license number
    if len(split_line) > 40 and split_line[1][0:3] == 'FFS':
      split_line[12] = 'Team ' + split_line[9]
      # PrintLine(split_line)
      num_racer += 1
    print('|'.join(split_line).rstrip())

r = ReadRace('equipe.sav')
ModifyRace(r)
