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
    return sh.sheet1.get_all_values()


def write_sheet(x:list):
    gc = gspread.service_account(filename='service_account.json')
    sh = gc.open('list_final')
    work = sh.sheet1
    work.append_row(x)

sheet_request = sheet_request()
sheet_import = sheet_import(planilha)
sheet_finally = sheet_finally()

def transfer_data():

    for res in sheet_request:
        for res2 in sheet_import:
            if res['Email'] == res2['Email']:
                list_insert.append(res)

    lenght = len(sheet_finally)

    for item in list_insert:
        i = 0
        for res in sheet_finally:
            if item['Email'] in res:
                print('Já existente')
            else:
                i += 1

        if i == lenght:
            write_sheet([item['Nome'], item['Email'], item['cpf'], item['telefone']])



transfer_data()
print(sheet_finally)
