import gspread

planilha = input("Digite o nome da planilha: ").lower()

list_insert = []
def sheet_request():
    gc = gspread.service_account(filename='service_account.json')
    sh = gc.open('Solicitacoes')
    return sh.sheet1.get_all_records()

def sheet_import(planilha):
    gc = gspread.service_account(filename='service_account.json')
    sh = gc.open(planilha)
    return sh.sheet1.get_all_records()

def sheet_finally():
    gc = gspread.service_account(filename='service_account.json')
    sh = gc.open('list_final')
    return sh.sheet1.get_all_records()

def write_sheet(x):
    gc = gspread.service_account(filename='service_account.json')
    sh = gc.open('list_final')
    work = sh.sheet1
    work.append_row()




sheet_req = sheet_request()

sheet_impor = sheet_import(planilha)
sheet_final = sheet_finally()

for res in sheet_req:
    for res2 in sheet_impor:
        if res['cpf'] == res2['cpf'] and res['Email'] == res2['Email']:
            list_insert.append(res)

i = 0
for item in list_insert:
    for res in sheet_final:
        if item['cpf'] != sheet_final['cpf'] and item['Email'] != sheet_final['Email']:




