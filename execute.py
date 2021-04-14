import gspread

class AutomationSheet:
    def __init__(self, importation, solicitation, median):
        self.solicitation = self.request_sheet(solicitation, 'dict')
        self.importation = self.request_sheet(importation, 'dict')
        self.median_list = self.request_sheet(median, 'list')
        self.median_dict = self.request_sheet(median, 'dict')
        self.list = []
        self.median = median
        self.lenght = 2

    @staticmethod
    def request_sheet(sheet, condition=None):
        gc = gspread.service_account('service_account.json')
        sh = gc.open(sheet)
        if condition == 'list':
            return sh.sheet1.get_all_values()
        elif condition == 'dict':
            return sh.sheet1.get_all_records()
        else:
            return sh

    def write(self, sheet, add: list):
        write_sheet = self.request_sheet(sheet)
        work = write_sheet.sheet1
        work.append_row(add)

    def update(self, sheet, x):
        update_sheet = self.request_sheet(sheet)
        update = update_sheet.sheet1
        update.update(str(f'D{x}'), 'Enviado')

    def write_comparison(self):
        for person in self.solicitation:
            for person_auth in self.importation:
                if person['Nome Completo (Para Certificado)'] == person_auth['Nome'] or person['Endereço de e-mail'] == person_auth['Email']:
                    self.list.append(person)
        return self.list

    def start(self):
        list_comparison = self.write_comparison()
        for item in list_comparison:
            i = 0
            for res in self.median_list:
                if (item['Nome Completo (Para Certificado)'] or item['Endereço de e-mail']) in res:
                    continue
                else:
                    i += 1
            if i == len(self.median_list):
                self.write(self.median, [item['Nome Completo (Para Certificado)'], item['Endereço de e-mail'],
                                         item['CPF no Formato XXX.XXX.XXX-XX (Para Certificado)']])

    def update_cpf(self, sheet):
        i = 2
        update_sheet = self.request_sheet(sheet)
        update = update_sheet.sheet1
        sh = update.get_all_records()
        for item in sh:
            if len(str(item['cpf'])) == 11:
                cpf = str(item['cpf'])
                replace = f'{cpf[0:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
                update.update(str(f'C{i}'), replace)

            if len(str(item['cpf'])) == 10:
                cpf = "0" + str(item['cpf'])
                replace = f'{cpf[0:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
                update.update(str(f'C{i}'), replace)

            if len(str(item['cpf'])) == 9:
                cpf = "00" + str(item['cpf'])
                replace = f'{cpf[0:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
                update.update(str(f'C{i}'), replace)
            i += 1

    def update_situation(self, send, median):
        for item in self.median_dict:
            if item['situacao'] == 'Enviado':
                print("enviado")

            elif item['situacao'] != 'Enviado':
                self.write(send, [item['Nome'], item['Email'], item['cpf']])
                self.update(median, self.lenght)

            self.lenght += 1

if __name__ == '__main__':
    planilha_import = input('Nome da planilha importada (sem extensão): ')
    planilha_request = input('Nome da planilha de solicitação: ')
    planilha_intermed = input('Nome da planilha intermediadora: ')
    planilha_send = input('Planilha para enviar os dados: ')

    test = AutomationSheet(planilha_import, planilha_request, planilha_intermed)
    test.start()
    test.update_situation(planilha_send, planilha_intermed)
    test.update_cpf(planilha_send)
