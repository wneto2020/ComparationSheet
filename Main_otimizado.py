import gspread

class ComparationSheet:

    def __init__(self, planilha):
        self.request = self.sheet_request('Solicitacoes', 'dict')
        self.imp = self.sheet_request(planilha, 'dict')
        self.final = self.sheet_request('planilha intermed', 'lista')
        self.send = self.sheet_request('planilha intermed', 'dict')
        self.list = []
        self.lenght = 0
        self.cont = 2

    def sheet_request(self, sheet, option):
        gc = gspread.service_account(filename='service_account.json')
        sh = gc.open(sheet)
        if option == 'lista':
            return sh.sheet1.get_all_values()

        elif option == 'dict':
            return sh.sheet1.get_all_records()

    def write_sheet(self, sheet, x: list):
        gc = gspread.service_account(filename='service_account.json')
        sh = gc.open(sheet)
        work = sh.sheet1
        work.append_row(x)

    def update_status(self, x):
        gc = gspread.service_account(filename='service_account.json')
        sh = gc.open('planilha intermed')
        work = sh.sheet1
        work.update(f'E{str(x)}', 'Enviado')

    def treatment_data(self, sheet):
        if len(str(sheet['cpf'])) == 11:
            sheet['cpf'] = str(sheet['cpf'])
            cpf_replace = f'{sheet["cpf"][0:3]}.{sheet["cpf"][3:6]}.{sheet["cpf"][6:9]}-{sheet["cpf"][9:]}'
            sheet['cpf'] = cpf_replace
            return sheet['cpf']

        elif len(str(sheet['cpf'])) == 10:
            sheet['cpf'] = '0' + str(sheet['cpf'])
            cpf_replace = f'{sheet["cpf"][0:3]}.{sheet["cpf"][3:6]}.{sheet["cpf"][6:9]}-{sheet["cpf"][9:]}'
            sheet['cpf'] = cpf_replace
            return sheet['cpf']

        elif len(str(sheet['cpf'])) == 9:
            sheet['cpf'] = '00' + str(sheet['cpf'])
            cpf_replace = f'{sheet["cpf"][0:3]}.{sheet["cpf"][3:6]}.{sheet["cpf"][6:9]}-{sheet["cpf"][9:]}'
            sheet['cpf'] = cpf_replace
            return sheet['cpf']

        else:
            return sheet['cpf']

    def transfer_data(self):
        for data in self.request:
            if str(data['cpf']) ==11 or str(data['cpf']) == 10 or str(data['cpf']) == 9:
                data['cpf'] = self.treatment_data(data)

            for data2 in self.imp:
                if str(data['cpf']) == 11 or str(data['cpf']) == 10 or str(data['cpf']) == 9:
                    data2['cpf'] = self.treatment_data(data2)

                if data['Email'] == data2['Email'] or data['cpf'] == data2['cpf']:
                    self.list.append(data)

        for item in self.list:
            print(item)
            for res in self.final:
                if item['Email'] in res or item['cpf'] in res:
                    print(item['Email'], 'Já existente')

                else:
                    self.lenght += 1
            if self.lenght == len(self.list):
                self.write_sheet('planilha intermed', [item['Nome'], item['Email'], item['cpf'], item['telefone']])

    def ultimate_sheet(self):
        for res in self.send:
            if res['situacao'] == 'Enviado':
                self.cont += 1

            elif res['situacao'] != 'Enviado':
                self.update_status(str(self.cont))
                self.write_sheet('Planilha de envio', [res['Nome'], res['Email'], res['cpf'], res['telefone']])
                self.cont += 1


if __name__ == "__main__":
    transfer = ComparationSheet('planilha importada')
    transfer.transfer_data()
    transfer.ultimate_sheet()
