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

def write_sheet(x: list):
    gc = gspread.service_account(filename='service_account.json')
    sh = gc.open('list_final')
    work = sh.sheet1
    work.append_row(x)

sheet_request = sheet_request()
sheet_import = sheet_import(planilha)
sheet_finally = sheet_finally()

def transfer_data():

    for res in sheet_request:
        if len(str(res['cpf'])) == 11:
            cpf_replace = f'{res["cpf"][0:3]}.{res["cpf"][3:6]}.{res["cpf"][6:9]}-{res["cpf"][9:]}'
            res['cpf'] = cpf_replace

        elif len(str(res['cpf'])) == 10:
            res['cpf'] = '0'+str(res['cpf'])
            cpf_replace = f'{res["cpf"][0:3]}.{res["cpf"][3:6]}.{res["cpf"][6:9]}-{res["cpf"][9:]}'
            res['cpf'] = cpf_replace

        elif len(str(res['cpf'])) == 9:
            res['cpf'] = '00'+str(res['cpf'])
            cpf_replace = f'{res["cpf"][0:3]}.{res["cpf"][3:6]}.{res["cpf"][6:9]}-{res["cpf"][9:]}'
            res['cpf'] = cpf_replace


        for res2 in sheet_import:
            if len(str(res2['cpf'])) == 11:
                cpf_replace = f'{res2["cpf"][0:3]}.{res2["cpf"][3:6]}.{res2["cpf"][6:9]}-{res2["cpf"][9:]}'
                res2['cpf'] = cpf_replace

            elif len(str(res2['cpf'])) == 10:
                res2['cpf'] = '0' +str(res['cpf'])
                cpf_replace = f'{res2["cpf"][0:3]}.{res2["cpf"][3:6]}.{res2["cpf"][6:9]}-{res2["cpf"][9:]}'
                res2['cpf'] = cpf_replace

            elif len(str(res2['cpf'])) == 9:
                res2['cpf'] = '00' + str(res['cpf'])
                cpf_replace = f'{res2["cpf"][0:3]}.{res2["cpf"][3:6]}.{res2["cpf"][6:9]}-{res2["cpf"][9:]}'
                res2['cpf'] = cpf_replace

            if res['Email'] == res2['Email'] or res['cpf'] == res2['cpf']:
                list_insert.append(res)

    lenght = len(sheet_finally)

    for item in list_insert:
        i = 0
        for res in sheet_finally:
            if item['Email'] in res or item['cpf'] in res:
                print(item['Email'],'Já existente')
            else:
                i += 1

        if i == lenght:
            write_sheet([item['Nome'], item['Email'], item['cpf'], item['telefone']])

transfer_data()
