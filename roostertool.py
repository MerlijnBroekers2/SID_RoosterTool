import pandas as pd
import numpy as np
import gurobipy as gp
# import xlsxwriter module
import xlsxwriter
from xlwt import Workbook
import xlwt
import math

### INPUT ###
make_library = True
make_pulse = False
make_fellow = False
make_echo = False

max_uren = 100
zondag_quota = 20
### INPUT ###

df = pd.read_excel(r'Beschikbaarheid.xlsx', sheet_name='Hele Team')
availability = pd.read_excel(
    r'Beschikbaarheid.xlsx', header=1, sheet_name='Hele Team')
poules = ['Poule Library']


def calc_hour(hours, poule):
    if hours == '-' or hours == '' or hours == ' ' or isinstance(hours, float):
        return 0
    try:
        hour_start = int(hours[0:2])
        min_start = int(hours[3:5])
        hour_finish = int(hours[6:8])
        min_finish = int(hours[9:11])
    except:
        raise ValueError('hours not specified correctly:' +
                         str(hours), ' in: '+poule)
    go = True
    min = 0
    while go:
        min_start += 1
        min += 1
        if min_start == 60:
            hour_start += 1
            min_start = 0
            if hour_start == 24:
                hour_start = 0
        if min_start == min_finish and hour_start == hour_finish:
            return min/60


# Haal alle uren op die in shifts zitten. Dit is nodig om de hoeveelheid uren eerlijk te verdelen. De uren zijn opgeslagen in de dictionary "shifts".
shifts = dict()
shifts['time'] = dict()
shifts['hours'] = dict()
shifts['persons'] = dict()
shifts['type'] = []
shifts['dag'] = []
shifts['datum'] = []

for index, row in availability.iterrows():
    for poule in poules:
        if poule not in shifts['time']:
            shifts['time'][poule] = []
            shifts['hours'][poule] = []
            shifts['persons'][poule] = []
        if isinstance(row['Type'], str):
            shifts['time'][poule].append(row[poule])
            shifts['hours'][poule].append(calc_hour(row[poule], poule))
            if math.isclose(calc_hour(row[poule], poule), 0):
                shifts['persons'][poule].append(0)
            else:
                if poule == 'Poule Library':
                    shifts['persons'][poule].append(int(row['Benodigd (lib)']))
                else:
                    shifts['persons'][poule].append(1)
    if isinstance(row['Type'], str):
        shifts['type'].append(row['Type'])
        if isinstance(row['Dag'], str):
            dag = row['Dag']
        shifts['dag'].append(dag)
        if isinstance(row['Dag'], str):
            datum = row['Datum']
        shifts['datum'].append(datum)
        # if isinstance(row['Benodigd (lib)'], str):
        #     datum = row['Benodigd (lib)']
        # shifts['Benodigd (lib)'].append(datum)


# Function om te vinden wie als eerste persoon van een poule in de excel lijst staat.
start_persoons = dict()
for index, row in df.iterrows():
    for poule in poules:
        start_persoons[row[poule]] = poule
    break

# De dictionary "beschikbaarheid" bevat alle beschikbaarheid van medewerkers in lijsten. De lijnen hieronder zetten alle mensen in deze dictionary.
all_persoons = set()
beschikbaarheid = dict()
for index, row in df.iterrows():
    poule = None
    for persoon in row:
        if persoon in start_persoons:
            poule = start_persoons[persoon]
            beschikbaarheid[poule] = dict()
        if poule is not None:
            beschikbaarheid[poule][persoon] = []
            all_persoons.add(persoon)

# Zet alle beschikbaarheid in de lijst.
for index, row in availability.iterrows():
    for poule in poules:
        for persoon in beschikbaarheid[poule]:
            if not isinstance(row[persoon], float):
                if row[persoon] == 'j' or row[persoon] == 'J' or row[persoon] == 'x' or row[persoon] == 'X':
                    beschikbaarheid[poule][persoon].append(1)
                else:
                    beschikbaarheid[poule][persoon].append(0)
            else:
                beschikbaarheid[poule][persoon].append(0)

# Calc totoal availability per shift.
total_available = dict()
for poule in poules:
    total_available[poule] = [0] * len(shifts['hours'][poule])
    for shift in range(len(shifts['hours'][poule])):
        available = 0
        for availableLIST in beschikbaarheid[poule].values():
            available += availableLIST[shift]
        total_available[poule][shift] = available

# Het kan zijn dat er niemand beschikbaar is voor een shift. Daarom wordt de fictieve persoon NoOne toegevoegd. Deze persoon is altijd beschikbaar, als niemand beschikbaar is.
for poule in poules:
    beschikbaarheid[poule]['NoOne_1'] = [0] * len(shifts['hours'][poule])
    beschikbaarheid[poule]['NoOne_2'] = [0] * len(shifts['hours'][poule])
    for shift in range(len(shifts['hours'][poule])):
        if shifts['persons'][poule][shift] >= 1:
            beschikbaarheid[poule]['NoOne_1'][shift] = 1
        if shifts['persons'][poule][shift] == 2:
            beschikbaarheid[poule]['NoOne_2'][shift] = 1

# Bekeren de totale beschikbaarheid per poule
opgegeven_uren = dict()
opgegeven_uren['totaal'] = dict()
opgegeven_uren['niet-zondag'] = dict()
for poule in poules:
    opgegeven_uren['totaal'][poule] = dict()
    opgegeven_uren['totaal'][poule]['totaal'] = 0
    opgegeven_uren['niet-zondag'][poule] = dict()
    opgegeven_uren['niet-zondag'][poule]['totaal'] = 0
    for persoon, beschikbaarLIJST in beschikbaarheid[poule].items():
        opgegeven_uren['totaal'][poule][persoon] = 0
        opgegeven_uren['niet-zondag'][poule][persoon] = 0
        for shift_nr, shift_length in enumerate(shifts['hours'][poule]):
            opgegeven_uren['totaal'][poule]['totaal'] += shift_length * \
                beschikbaarLIJST[shift_nr]
            opgegeven_uren['totaal'][poule][persoon] += shift_length * \
                beschikbaarLIJST[shift_nr]
            if shifts['dag'][shift_nr] != 'Zondag':
                opgegeven_uren['niet-zondag'][poule]['totaal'] += shift_length * \
                    beschikbaarLIJST[shift_nr]
                opgegeven_uren['niet-zondag'][poule][persoon] += shift_length * \
                    beschikbaarLIJST[shift_nr]

# bereken het aandeel opgegeven uren van elk persoon.
aandeel = dict()
aandeel['opgegeven'] = dict()
for poule in poules:
    aandeel['opgegeven'][poule] = dict()
    for persoon, uren in opgegeven_uren['totaal'][poule].items():
        aandeel['opgegeven'][poule][persoon] = uren / \
            opgegeven_uren['totaal'][poule]['totaal']

# Maar voor elke poule een rooster.
for poule in poules:
    # Sla de poule over als er geen rooster voor gemaakt moet worden.
    if not make_pulse and poule == 'Poule Pulse':
        continue
    elif not make_fellow and poule == 'Poule Fellowship':
        continue
    elif not make_library and poule == 'Poule Library':
        continue
    elif not make_echo and poule == 'Poule Echo':
        continue

    '''Initiate model'''
    # Open het optimisatie model
    m = gp.Model()
    m.setParam('TimeLimit', 30)
    '''Define variables'''
    # De "A" variable is binairy (0 of 1). De variable is 1 als de persoon de shift krijgt. de variable is 0 als de persoon de shift niet krijgt.
    A = dict()
    for persoon in beschikbaarheid[poule]:
        for shift, hours in enumerate(shifts['hours'][poule]):
            A[(persoon, shift)] = m.addVar(vtype=gp.GRB.BINARY,
                                           name='A(' + persoon + ',' + str(shift) + ')')

    ''''Set objective function'''
    # De objective (het doel) is het minimaliseren van het verschil in het aandeel opgegeven uren en het aandeel gekregen uren. Dit verschil is ook wel de "error".
    error = dict()
    M = dict()
    objective = 0
    for persoon in beschikbaarheid[poule]:
        if persoon == 'NoOne_1':
            # # Use very high penalty for assigning this person to a shift.
            for shift, hours in enumerate(shifts['hours'][poule]):
                objective += 1000 * A[(persoon, shift)]
        elif persoon == 'NoOne_2':
            # # Use an even higher penalty for assigning this person to a shift.
            for shift, hours in enumerate(shifts['hours'][poule]):
                objective += 2000 * A[(persoon, shift)]
        else:
            error[persoon] = m.addVar(
                ub=np.inf, lb=-np.inf, name='error('+persoon+')')
            objective += error[persoon] * error[persoon]
    m.setObjective(objective)

    '''Constraint 1'''
    # Deze constraint is toegevoegd zodat de error berekend wordt.
    aandeel['gekregen'] = dict()
    aandeel['gekregen'][poule] = dict()
    for persoon in beschikbaarheid[poule]:
        if persoon == 'NoOne_1' or persoon == 'NoOne_2':
            continue
        add = 0
        for shift, hours in enumerate(shifts['hours'][poule]):
            add += A[(persoon, shift)] * hours
        aandeel['gekregen'][poule][persoon] = add / sum(shifts['hours'][poule])
        m.addConstr(error[persoon], gp.GRB.EQUAL, aandeel['opgegeven'][poule][persoon] -
                    aandeel['gekregen'][poule][persoon], name='C1: Error of ' + persoon)

    '''Constraint 2'''
    # Deze constraint zorgt er voor dat er per shift, precies 1 persoon wordt aangewezen.
    for shift, hours in enumerate(shifts['hours'][poule]):
        add = 0
        for persoon in beschikbaarheid[poule]:
            add += A[(persoon, shift)]
        m.addConstr(add, gp.GRB.EQUAL, shifts['persons'][poule][shift],
                    name='C2: 1 persoon per shift for shift: ' + str(shift))

    '''Constraint 3'''
    # Deze constraint zorgt ervoor dat iemand geen shift wordt toegewezen als hij/zij niet beschikbaar is.
    for persoon, beschikbaar in beschikbaarheid[poule].items():
        for shift, hours in enumerate(shifts['hours'][poule]):
            m.addConstr(A[(persoon, shift)], gp.GRB.LESS_EQUAL, beschikbaar[shift],
                        name='C3: ' + persoon + ' available for shift ' + str(shift))

    '''Constraint 4'''
    # Deze constraint zorgt ervoor dat mensen niet zowel savonds als de volgende ochtend kunnen werken.
    for persoon, beschikbaar in beschikbaarheid[poule].items():
        if persoon == 'NoOne_1' or persoon == 'NoOne_2':
            continue
        prev_type = 'None'
        prev_shift = -1
        for shift, type in enumerate(shifts['type']):
            if prev_type == 'Avond' and type == 'Ochtend':
                m.addConstr(A[(persoon, shift)] + A[(persoon, prev_shift)], gp.GRB.LESS_EQUAL, 1,
                            name='C4: ' + persoon + ' not allowed at shift ' + str(prev_shift) + ' and ' + str(shift))
            prev_type = type
            prev_shift = shift

    '''Constraint 5'''
    # Deze constraint zorgt ervoor dat mensen, wanneer zij s'avonds werken, niet ook sochtends (en smiddags) kunnen werken.
    for persoon, beschikbaar in beschikbaarheid[poule].items():
        if persoon == 'NoOne_1' or persoon == 'NoOne_2':
            continue
        prev_type = 'None'
        prev_shift = -1
        prev_prev_type = 'None'
        prev_prev_shift = -2
        for shift, type in enumerate(shifts['type']):
            if type == 'Avond':
                if prev_type == 'Middag' or prev_type == 'Ochtend':
                    m.addConstr(A[(persoon, prev_shift)] + A[(persoon, shift)], gp.GRB.LESS_EQUAL, 1, name='C5.1: ' +
                                persoon + ' not allowed on midday shift ' + str(prev_shift) + ' and evening shift ' + str(shift))
                if prev_prev_type == 'Ochtend':
                    m.addConstr(A[(persoon, prev_prev_shift)] + A[(persoon, shift)], gp.GRB.LESS_EQUAL, 1, name='C5.2: ' +
                                persoon + ' not allowed on morning shift ' + str(prev_prev_shift) + ' and evening shift ' + str(shift))
            prev_prev_type = prev_type
            prev_prev_shift = prev_shift
            prev_type = type
            prev_shift = shift

    '''Constraint 6'''
    # Geef mensen geen zondag shifts als quata niet wordt gehaald.
    for persoon, beschikbaar in beschikbaarheid[poule].items():
        if persoon == 'NoOne_1' or persoon == 'NoOne_2':
            continue
        if opgegeven_uren['niet-zondag'][poule][persoon] < zondag_quota:
            for shift, dag in enumerate(shifts['dag']):
                if dag == 'Zondag':
                    m.addConstr(A[(persoon, shift)], gp.GRB.EQUAL, 0, name='C6: ' +
                                persoon + ' not allowed on sunday at shift: ' + str(shift))

    '''Constraint 7'''
    # Geef mensen maar maximaal 1 zondag shift
    for persoon in beschikbaarheid[poule]:
        if persoon == 'NoOne_1' or persoon == 'NoOne_2':
            continue
        add = 0
        for shift, hours in enumerate(shifts['hours'][poule]):
            if shifts['dag'][shift] == 'Zondag':
                add += A[(persoon, shift)]
        m.addConstr(add, gp.GRB.LESS_EQUAL, 1,
                    name='C7: 1 zondag shift voor: ' + str(persoon))

    '''Constraint 8'''
    # Maximaal aantal uren per persoon.
    for persoon in beschikbaarheid[poule]:
        if persoon == 'NoOne_1' or persoon == 'NoOne_2':
            continue
        add = 0
        for shift, hours in enumerate(shifts['hours'][poule]):
            add += shifts['hours'][poule][shift] * A[(persoon, shift)]
        m.addConstr(add, gp.GRB.LESS_EQUAL, max_uren,
                    name='C8: maximaal ' + str(max_uren) + ' voor: ' + str(persoon))

    m.optimize()
    print(m.display())

    # workbook = xlsxwriter.Workbook('Rooster van '+ poule +'.xlsx')
    # worksheet = workbook.add_worksheet()

    workbook = Workbook()
    sheet = workbook.add_sheet('Sheet 1')

    # schrijf dienst dagen en tijden
    prev_dag = 'Start'
    row = 0
    column = 1

    col_width = [(2, 15), (3, 15), (4, 20), (6, 15),
                 (7, 20), (9, 15), (10, 20)]
    for col, width in col_width:
        sheet.col(col).width = 256 * width

    for shift, dag in enumerate(shifts['dag']):

        datum = str(shifts['datum'][shift].day) + '/' + \
            str(shifts['datum'][shift].month)
        time = shifts['time'][poule][shift]
        if dag != prev_dag or dag == 'Start':
            column = 1
            # new dag
            row += 1
            if prev_dag == 'Zondag' and dag == 'Maandag':
                row += 1
            sheet.write(row, column, datum)
            column += 1
            sheet.write(row, column, dag)
        column += 1
        sheet.write(row, column, time)
        column += 1
        # determine name who works:
        names = []
        number_needed = shifts['persons'][poule][shift]
        for persoon in beschikbaarheid[poule]:
            if math.isclose(A[(persoon, shift)].x, 1):
                names.append(persoon)
        for i in range(number_needed):
            if names[i] == 'NoOne_1':
                st = xlwt.easyxf('pattern: pattern solid;')
                if total_available[poule][shift] == 0:
                    st.pattern.pattern_fore_colour = 45
                    name = 'Niemand beschikbaar'
                else:
                    st.pattern.pattern_fore_colour = 43
                    name = 'Onhaalbaar'
                sheet.write(row, column, name, st)
            elif names[i] == 'NoOne_2':
                st = xlwt.easyxf('pattern: pattern solid;')
                if total_available[poule][shift] == 1:
                    st.pattern.pattern_fore_colour = 45
                    name = 'Niemand beschikbaar'
                else:
                    st.pattern.pattern_fore_colour = 43
                    name = 'Onhaalbaar'
                sheet.write(row, column, name, st)
            else:
                sheet.write(row, column, names[i])
            column += 1
        if number_needed == 2:
            column = column - 1
        prev_dag = dag

    workbook.save('Rooster_van_' + poule + '.xlsx')
