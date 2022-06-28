import os
import docx
import json
import time

cfg = None
with open('config.json', 'rb') as file:
    cfg = file.read().decode('utf-8')
    cfg = json.loads(cfg)


allTableNames = []
for t in cfg['tables']:
    allTableNames.append(t['friendlyName'])

state = None
try:
    with open('state.json', 'rb') as file:
        state = file.read().decode('utf-8')
        state = json.loads(state)
except FileNotFoundError:
    state = {k: {'cursor': 0} for k in allTableNames}

includedTables = set(allTableNames)
userInput = ''
while not '/' in userInput:
    i = 0
    for tn in allTableNames:
        print(f'{"+" if tn in includedTables else "-"}{i} {tn}')
        i-=-1
    try:
        userInput = input('Оберіть потрібні таблиці (наприклад: "-1-4/" означає викреслити таблиці з номерами 1 та 4 і підтвердити):\n')
    except KeyboardInterrupt:
        i = 0
        for tn in allTableNames:
            print(f'{"+" if tn in includedTables else "-"}{i} {tn}')
            i-=-1
        break
    i = 0
    while i<len(userInput):
        if userInput[i] == '+' or userInput[i] == '-':
            j = i+1
            temp = ''
            while j<len(userInput) and  userInput[j] >= '0' and userInput[j] <= '9':
                temp = temp+userInput[j]
                j-=-1
            try:
                temp = int(temp)
                name = allTableNames[temp]
                if userInput[i] == '+': includedTables.add(allTableNames[temp])
                if userInput[i] == '-': includedTables.remove(allTableNames[temp])
            except: pass
        i-=-1


shared_columns = {}
unique_columns = {}
for t in cfg['tables']:
    tableName = t['friendlyName']
    if not tableName in includedTables: continue

    for c in t['columns'].keys():
        type =  t['columns'][c]['type'] if 'type' in  t['columns'][c] else None
        if type == None or type == 'shared':
            shared_columns[c] = ''
            t['columns'][c]['type'] = 'shared' # do not tolerate None. Will use below
        elif type == 'unique':
            if not tableName in unique_columns:
                unique_columns[tableName] = {}
            unique_columns[tableName][c] = ''

def fill_all_columns(all: list):
    i = 0
    while i < len(all):
        d,k,c= all[i]
        try:
            v = input(f'Введіть значення "{k}" в "{c}":\n')
            d[k] = v
        except KeyboardInterrupt:
            i -= 2
        i-=-1
        if i<0:
            i = 0

        

allColumns = [(shared_columns,key,'загальні') for key in shared_columns.keys()]
for t in unique_columns.keys():
    allColumns += [(unique_columns[t], key, t) for key in unique_columns[t].keys()]
tiedPhoneNumbers = {}
keyForTiedPhoneNumbers = 'Номери через кому'
allColumns.append((tiedPhoneNumbers, keyForTiedPhoneNumbers, 'Телефони родичів'))
fill_all_columns(allColumns)


print("Перевірте порядок аркушів: ")
i = 1
for tn in allTableNames:
    print(f'{i}: {tn}')
    i-=-1
input("Перевірте порядок та натисніть Enter:")


served = []
doc:docx.Document = docx.Document('blank.docx')
tableIndex = -1

for tn in allTableNames:
    served.append(tn)
    tableIndex-=-1
    if not tn in includedTables:
        continue

    table = None
    for t in cfg['tables']:
        if t['friendlyName'] == tn:
            table = t
            break
    cursor = state[tn]['cursor']
    if cursor == 0:
        i = ''
        while not 'Y' in i and not 'y' in i and not 'т' in i and not 'Т' in i:
            i = input(f"Аркуш '{tn}' заповнено. Роздрукуйте і вкладіть наступний, введіть 'так' та натисніть Enter.")
    wtable = doc.tables[tableIndex]
    origHeight = wtable.rows[cursor+1].height
    row = wtable.row_cells(cursor+1) # +1 because first row is heading
    columnsOrderedNames = table['columnsOrder']
    unique_col_values = unique_columns[tn]
    for i in range(len(columnsOrderedNames)):
        column = table['columns'][columnsOrderedNames[i]]
        if column['type'] == 'autofill':
            row[i].text = str(cursor+1) #+1 for 1-based
        elif column['type'] == 'empty': pass
        else:
            row[i].text = shared_columns[columnsOrderedNames[i]] if columnsOrderedNames[i] in shared_columns else unique_col_values[columnsOrderedNames[i]]
    if wtable.rows[cursor+1].height != origHeight:
        print(f'Warning: height breach in {tn} table')
    cursor = ((cursor) + 1) % table['tableCount']
    state[tn]['cursor'] = cursor

while True:
    try:
        doc.save('temp.docx')
        break
    except PermissionError:
        time.sleep(0.2)
        continue
os.startfile('temp.docx', 'print')
time.sleep(10)


print(f'Served: {str(served)}')

state = json.dumps(state, ensure_ascii=False, indent=4)
with open('state.json', 'wb') as file:
    state = state.encode('utf-8')
    file.write(state)

for k in unique_columns.keys():
    shared_columns[k] = unique_columns[k]
shared_columns['TIED_NUMBERS'] = f'{tiedPhoneNumbers[keyForTiedPhoneNumbers]},{shared_columns["Телефон"]}'
with open('log.txt', 'ab+') as file:
    file.write(f'{json.dumps(shared_columns, ensure_ascii=False)},\n'.encode('utf-8'))