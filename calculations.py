''' This file contains all the functons that parse the json returned from d3up.
`cell` refers to the target cell in the excel sheet to write to. '''

from collections import namedtuple


Stat = namedtuple('Stat', 'cell, value')

# Mappings of json constants to cells in the monk toolbox
# d3: cell: expressed as percent
DPS = {'dexterity': ('B2', False), 
       'elite-damage': ('B3', True),
       'demon-damage': ('B4', True),
       'attack-speed-incs': ('B5', True),
       'attack-speed-incs-dw': ('', False),
       'critical-hit': ('B6', True),
       'critical-hit-damage': ('B7', True),
       'bonus-elemental-percent': ('B8', True),
       'plus-lightning-damage-skills': ('B9', True),
       'mk-fists-of-thunder': ('B10', True),
       'mk-sweeping-wind': ('B11', True),
       'plus-max-damage': ('F6', False),
       'plus-min-damage': ('E6', False),
       'ring1-max': ('F4', False),
       'ring1-min': ('E4', False),
       'ring2-max': ('F5', False),
       'ring2-min': ('E5', False)}

MAINHAND = {'dps-speed-mh': ('B17', False),
            'dps-mh-min': ('B20', False),
            'dps-mh-max': ('B21', False),
            'dps-mh-real-min': ('', False),
            'dps-mh-real-max': ('', False)}

OFFHAND = {'dps-speed-oh': ('B29', False),
           'dps-oh-min': ('B32', False),
           'dps-oh-max': ('B33', False),
           'dps-oh-real-min': ('', False),
           'dps-oh-real-max': ('', False)}

SUSTAIN = {'life-regen': ('D35', False),
           'life-steal': ('D32', True),
           'life-hit': ('D29', False),
           'spirit-spent-life': ('D38', False)}

EHP = {'armor': ('J28', False),
       'resist-all': ('J29', False),
       'vitality': ('J30', False),
       'plus-life': ('J31', True)}

# cell = x - y, as defined by the following mapping:
# d3: x, y, cell 
SUBTRACTS = {'dps-mh-elem-max': ('dps-mh-max', 'dps-mh-real-max', 'B23'),
             'dps-mh-elem-min': ('dps-mh-min', 'dps-mh-real-min', 'B22'),
             'dps-oh-elem-max': ('dps-oh-max', 'dps-oh-real-max', 'B35'),
             'dps-oh-elem-min': ('dps-oh-min', 'dps-oh-real-min', 'B34')}

def get_data(jsondict):
    return jsondict['stats'], jsondict['actives'], jsondict['passives']

def get_values(stats):
    data = {}
    for group in (DPS, MAINHAND, OFFHAND, SUSTAIN, EHP):
        for stat, (cell, percent) in sorted(group.items()):
            value = stats.get(stat, 0)
            if percent:
                value = float(value)/100.0
            data[stat] = Stat(cell, value)

    for stat, (total, real, cell) in SUBTRACTS.items():
        data[stat] = Stat(cell, stats[total] - stats[real])

    return data