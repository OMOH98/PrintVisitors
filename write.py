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
includedTables = set(allTableNames)
userInput = ''
while not '/' in userInput:
    i = 0
    for tn in allTableNames:
        print(f'{"+" if tn in includedTables else "-"}{i} {tn}')
        i-=-1
    userInput = input('Type a string without delimiters, +1 to include second table, -0 to exclude first one, etc:\n')
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


# shared_columns = {}
# unique_columns = {}
# for t in cfg['tables']:
#     tableName = t['friendlyName']
#     if not tableName in includedTables: continue

#     for c in t['columns'].keys():
#         type =  t['columns'][c]['type'] if 'type' in  t['columns'][c] else None
#         if type == None or type == 'shared':
#             shared_columns[c] = ''
#             t['columns'][c]['type'] = 'shared' # do not tolerate None. Will use below
#         elif type == 'unique':
#             if not tableName in unique_columns:
#                 unique_columns[tableName] = {}
#             unique_columns[tableName][c] = ''


# def fillColumn(d: dict, k:str, n = 'shared'):
#     v = input(f'Enter value for column "{k}" in "{n}":\n')
#     d[k] = v

# for k in shared_columns.keys():
#     fillColumn(shared_columns, k)
# for t in unique_columns.keys():
#     for k in unique_columns[t].keys():
#         fillColumn(unique_columns[t], k, t)

# shared_columns = {'name': 'Gleb', 'passport': 'TT255251', 'category': 'VNPO', 'phone': '0671141091'}
# unique_columns = {'На продуктові набори': {'issue': 'D1'}, 'На памперси': {'issue': 'XXL'}}
shared_columns = {'name': 'Gleb K', 'passport': 'TT255251', 'category': 'NVPO', 'phone': '0671141091', 'address': 'улица Пушкина, дом Колотушкина, кв зефира, №5.'}
unique_columns = {'На продуктові набори': {'issue': 'Д1'}, 'На білизну': {'issue': '1 біл. 2 рушн.'}, 'На дитяче харчування і памперси': {'issue': '4*Б'}, 'На дорослі памперси': {'issue': 'ХХL'}, 'На набори посуду': {'issue': '1'}, 'На набори гігієни': {'issue': '1'}}
print("Check tables order: ")
i = 1
for tn in allTableNames:
    print(f'{i}: {tn}')
    i-=-1
input("Check order and hit enter to proceed")
served = []
SLEEP_TIME = 9
for tn in allTableNames:
    served.append(tn)
    if not tn in includedTables:
        os.startfile('blank.docx', 'print')
        time.sleep(SLEEP_TIME)
        continue

    table = None
    for t in cfg['tables']:
        if t['friendlyName'] == tn:
            table = t
            break
    cursor = table['cursor']
    if cursor == 0:
        i = ''
        while not 'Y' in i and not 'y' in i:
            i = input(f"One table seems to be completed. Confirm that you changed '{tn}' (type yes): ")
    doc:docx.Document = docx.Document(table['file'])
    wtable = doc.tables[0]
    row = wtable.row_cells(table['cursor']+1) # +1 because first row is heading
    columnsOrderedNames = table['columnsOrder']
    unique_col_values = unique_columns[tn]
    for i in range(len(columnsOrderedNames)):
        column = table['columns'][columnsOrderedNames[i]]
        if column['type'] == 'autofill':
            row[i].text = str(cursor+1) #+1 for 1-based
        elif column['type'] == 'empty': pass
        else:
            row[i].text = shared_columns[columnsOrderedNames[i]] if columnsOrderedNames[i] in shared_columns else unique_col_values[columnsOrderedNames[i]]
    table['cursor'] = ((table['cursor']) + 1) % table['tableCount']
    while True:
        try:
            doc.save('temp.docx')
            break
        except PermissionError:
            time.sleep(0.2)
            continue
    os.startfile('temp.docx', 'print')
    time.sleep(SLEEP_TIME)

print(f'Served: {str(served)}')

cfg = json.dumps(cfg, ensure_ascii=False, indent=4)
with open('config.json', 'wb') as file:
    cfg = cfg.encode('utf-8')
    file.write(cfg)


print(shared_columns)
print(unique_columns)