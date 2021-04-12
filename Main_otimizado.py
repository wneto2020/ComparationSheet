import gspread

class AutomationSheet:
    def __init__(self):
        self.solicitation = self.request_sheet('Solicitacoes', 'dict')
        self.importation = self.request_sheet('planilha importada', 'dict')
        self.median_list = self.request_sheet('planilha intermed', 'list')
        self.median_dict = self.request_sheet('planilha intermed', 'dict')
        self.list = []
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
        update.update(str(f'E{x}'), 'Enviado')

    @staticmethod
    def treatment(dictionary: dict):
        cpf = str(dictionary['cpf'])
        if len(cpf) == 11:
            cpf = f'{cpf[0:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
            dictionary['cpf'] = cpf
            return dictionary['cpf']

        elif len(cpf) == 10:
            cpf = f'0{cpf[0:2]}.{cpf[2:5]}.{cpf[5:8]}-{cpf[8:]}'
            dictionary['cpf'] = cpf
            return dictionary['cpf']

        elif len(cpf) == 9:
            cpf = f'00{cpf[0:1]}.{cpf[1:4]}.{cpf[4:7]}-{cpf[7:]}'
            dictionary['cpf'] = cpf
            return dictionary['cpf']
        else:
            return dictionary['cpf']

    def write_comparison(self):
        for person in self.solicitation:
            person['cpf'] = self.treatment(person)

            for person_auth in self.importation:
                person_auth['cpf'] = self.treatment(person_auth)

                if person['Email'] == person_auth['Email'] or person['cpf'] == person_auth['cpf']:
                    self.list.append(person)
        return self.list

    def start(self):
        list_comparison = self.write_comparison()
        for item in list_comparison:
            i = 0
            for res in self.median_list:
                if (item['Email'] or item['cpf']) in res:
                    continue
                else:
                    i += 1
            if i == len(self.median_list):
                self.write('planilha intermed', [v for v in item.values()])

    def update_situation(self):
        for item in self.median_dict:
            if item['situacao'] == 'Enviado':
                print("enviado")

            elif item['situacao'] != 'Enviado':
                self.write('Planilha de envio', [v for v in item.values()])
                self.update('planilha intermed', self.lenght)

            self.lenght += 1

if __name__ == '__main__':
    test = AutomationSheet()
    test.start()
    test.update_situation()
